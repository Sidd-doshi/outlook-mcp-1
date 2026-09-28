import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphRequestRaw } from "../graph.js";
import { sanitizeFileList } from "../sanitize.js";
import type { Env } from "../types.js";
import { DANGEROUS_FILENAME_EXT, MAX_LARGE_UPLOAD_BYTES, pathSchema, uploadBytes } from "./files.js";

// Reading attachments on received email. Two tools:
//   list_email_attachments — what's attached (name, type, size), no bytes.
//   save_email_attachment  — copy one attachment straight into OneDrive.
// Bytes never travel through the conversation: the agent saves the file, then
// reads it with its own tooling (e.g. rclone + pdftotext) or attaches it onward
// with onedrive_path.

const messageIdSchema = z.string().min(1).max(512);

type RawAttachment = {
	id?: string;
	name?: string;
	contentType?: string;
	size?: number;
	isInline?: boolean;
	"@odata.type"?: string;
};

export type AttachmentSummary = {
	id: string;
	name: string;
	content_type: string | null;
	size: number | null;
	is_inline: boolean;
	kind: "file" | "item" | "reference";
};

function kindOf(odataType: string | undefined): AttachmentSummary["kind"] {
	if (odataType?.endsWith("itemAttachment")) return "item";
	if (odataType?.endsWith("referenceAttachment")) return "reference";
	return "file";
}

// Characters OneDrive rejects in a file name, plus ones our path schema forbids.
// Attachment names come from whoever sent the email, so they're cleaned, not trusted.
export function safeFileName(name: string): string {
	const cleaned = name
		.replace(/[\\/:*?"<>|#%\u0000-\u001f]/g, "-")
		.replace(/[^\w \-.,()'!&+@$]/g, "-")
		.replace(/\s+/g, " ")
		.replace(/^[.\s]+|[.\s]+$/g, "")
		.slice(0, 200);
	return cleaned || "attachment";
}

export async function listEmailAttachmentsImpl(
	env: Env,
	args: { message_id: string },
): Promise<{ success: true; message_id: string; attachments: AttachmentSummary[] }> {
	const data = (await graphGet(
		env,
		`/me/messages/${encodeURIComponent(args.message_id)}/attachments`,
		{ $select: "id,name,contentType,size,isInline" },
	)) as { value?: RawAttachment[] };
	const attachments = (data.value ?? []).map((a) => ({
		id: String(a.id ?? ""),
		name: String(a.name ?? ""),
		content_type: a.contentType ?? null,
		size: a.size ?? null,
		is_inline: !!a.isInline,
		kind: kindOf(a["@odata.type"]),
	}));
	return { success: true, message_id: args.message_id, attachments };
}

// Pick the attachment to save: by id, by name (case-insensitive exact), or the
// only non-inline file attachment. Ambiguity is an error that lists the choices,
// never a guess.
export function chooseAttachment(
	list: AttachmentSummary[],
	args: { attachment_id?: string; attachment_name?: string },
): AttachmentSummary {
	const describe = (xs: AttachmentSummary[]) =>
		xs.map((a) => `"${a.name}" (${a.content_type ?? "unknown type"}, id ${a.id})`).join("; ");
	if (args.attachment_id) {
		const hit = list.find((a) => a.id === args.attachment_id);
		if (!hit) throw ToolError.validation(`No attachment with id ${args.attachment_id}. Attachments: ${describe(list) || "none"}.`);
		return hit;
	}
	if (args.attachment_name) {
		const want = args.attachment_name.toLowerCase();
		const hits = list.filter((a) => a.name.toLowerCase() === want);
		if (hits.length === 1) return hits[0]!;
		if (hits.length > 1) throw ToolError.validation(`More than one attachment is called "${args.attachment_name}". Pass attachment_id instead: ${describe(hits)}.`);
		throw ToolError.validation(`No attachment called "${args.attachment_name}". Attachments: ${describe(list) || "none"}.`);
	}
	const files = list.filter((a) => a.kind === "file" && !a.is_inline);
	if (files.length === 1) return files[0]!;
	if (files.length === 0) throw ToolError.validation(`This email has no file attachments to save. Attachments: ${describe(list) || "none"}.`);
	throw ToolError.validation(`This email has ${files.length} attachments. Say which one with attachment_name or attachment_id: ${describe(files)}.`);
}

export async function saveEmailAttachmentImpl(
	env: Env,
	args: {
		message_id: string;
		destination_folder: string;
		attachment_id?: string;
		attachment_name?: string;
		file_name?: string;
		conflict_behavior?: "rename" | "replace" | "fail";
	},
): Promise<unknown> {
	const { attachments } = await listEmailAttachmentsImpl(env, { message_id: args.message_id });
	const chosen = chooseAttachment(attachments, args);

	if (chosen.kind !== "file") {
		throw ToolError.validation(
			chosen.kind === "item"
				? `"${chosen.name}" is an attached email or calendar item, not a file, so it can't be saved to OneDrive.`
				: `"${chosen.name}" is a link to a cloud file, not an attached file. Open the link's source instead.`,
		);
	}
	if ((chosen.size ?? 0) > MAX_LARGE_UPLOAD_BYTES) {
		throw ToolError.validation(
			`"${chosen.name}" is ${((chosen.size ?? 0) / 1024 / 1024).toFixed(1)} MB. The limit is ${MAX_LARGE_UPLOAD_BYTES / 1024 / 1024} MB.`,
		);
	}

	const fileName = safeFileName(args.file_name ?? chosen.name);
	if (DANGEROUS_FILENAME_EXT.test(fileName)) {
		throw ToolError.validation(`Won't save "${fileName}": that file type isn't allowed.`);
	}
	const folder = args.destination_folder.replace(/\/+$/, "");
	const itemPath = `${folder}/${fileName}`;

	const response = await graphRequestRaw(
		env,
		`/me/messages/${encodeURIComponent(args.message_id)}/attachments/${encodeURIComponent(chosen.id)}/$value`,
	);
	const bytes = new Uint8Array(await response.arrayBuffer());
	if (bytes.byteLength === 0) {
		throw new ToolError({
			userMessage: `The attachment "${chosen.name}" came back empty.`,
			internalMessage: "attachment $value returned 0 bytes",
			upstreamName: "Graph",
		});
	}

	const mime = (chosen.content_type ?? "application/octet-stream").toLowerCase();
	const item = await uploadBytes(env, itemPath, bytes, mime, args.conflict_behavior ?? "rename");
	return {
		success: true,
		attachment: { id: chosen.id, name: chosen.name, content_type: chosen.content_type, size: bytes.byteLength },
		saved_to: itemPath,
		file: sanitizeFileList([item])[0],
	};
}

export const attachmentTools = defineTools<Env>({
	list_email_attachments: {
		description:
			"List the attachments on a received email: each one's id, name, content type, size, whether it's inline (e.g. a signature logo), and kind (file, item = an attached email or event, reference = a cloud link). Doesn't return the contents. Use save_email_attachment to put a file into OneDrive.",
		schema: z.object({ message_id: messageIdSchema }),
		handler: (env, args) => listEmailAttachmentsImpl(env, args),
	},

	save_email_attachment: {
		description:
			'Save a file attached to a received email straight into OneDrive, without passing it through the conversation. Choose the attachment with attachment_name (exact, case-insensitive) or attachment_id; if the email has exactly one non-inline file attachment you can omit both. destination_folder is a OneDrive folder path (it must exist; create it with create_folder). file_name defaults to the attachment\'s own name, cleaned of characters OneDrive rejects. conflict_behavior defaults to "rename", so nothing is overwritten. Up to 25 MB. Returns the saved OneDrive path and driveItem.',
		schema: z.object({
			message_id: messageIdSchema,
			destination_folder: pathSchema,
			attachment_id: z.string().min(1).max(1024).optional(),
			attachment_name: z.string().min(1).max(255).optional(),
			file_name: z.string().min(1).max(255).optional(),
			conflict_behavior: z.enum(["rename", "replace", "fail"]).optional(),
		}),
		handler: (env, args) => saveEmailAttachmentImpl(env, args),
	},
});
