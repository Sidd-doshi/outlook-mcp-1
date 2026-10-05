import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import {
	graphGet,
	graphPost,
	graphPostAccepted,
	graphPutBinary,
	graphRequestRaw,
} from "../graph.js";
import { sanitizeFileList } from "../sanitize.js";
import type { Env } from "../types.js";
import { encodeOneDrivePath } from "./_shared.js";

// OneDrive paths: alphanumerics, spaces, dots, dashes, underscores, slashes,
// parentheses, apostrophes, and a few common safe punctuation marks. Disallows
// query/fragment chars (`?`, `#`, `&`, `=`) and control chars to prevent path
// injection. Path is encoded per-segment before interpolation regardless.
export const pathSchema = z
	.string()
	.min(1)
	.max(1024)
	.regex(
		/^[\w \-.,()'!&+@$\/]+$/,
		"path may contain letters, numbers, spaces, and . - _ , ( ) ' ! & + @ $ /",
	);

// ── Upload limits + MIME allowlist ────────────────────────────────────────────

// Microsoft Graph's "simple upload" (PUT to :/content) supports up to 4 MB.
// We cap slightly below to leave headroom for transport overhead.
export const MAX_UPLOAD_BYTES = 4 * 1024 * 1024;

// Uploads go to the user's own OneDrive (no third-party delivery), so the
// allowlist is broader than the email-attachment one — it includes plain-text
// formats the caller is likely to need (vCard, iCalendar, Markdown, XML). It
// still rejects executables and dangerous MIME types.
export const UPLOAD_ALLOWED_MIME_TYPES = new Set([
	// Documents
	"application/pdf",
	"application/msword",
	"application/vnd.ms-excel",
	"application/vnd.ms-powerpoint",
	"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	"application/vnd.openxmlformats-officedocument.presentationml.presentation",
	// Images (excludes SVG — XML, can carry <script>)
	"image/png",
	"image/jpeg",
	"image/jpg",
	"image/gif",
	"image/webp",
	// Text formats
	"text/plain",
	"text/csv",
	"text/html",
	"text/markdown",
	"text/vcard",
	"text/x-vcard",
	"text/calendar",
	"application/json",
	"application/xml",
	"text/xml",
	// Archives
	"application/zip",
]);

// Executable extensions blocked regardless of declared MIME type.
export const DANGEROUS_FILENAME_EXT = /\.(exe|bat|cmd|scr|msi|dll|ps1|vbs|com|cpl|jar|app)$/i;

// ── Pure helpers (exported for unit tests) ────────────────────────────────────

// Encode bytes as standard base64, in chunks so large files don't blow the
// argument limit of String.fromCharCode.
export function bytesToBase64(bytes: Uint8Array): string {
	let binary = "";
	const CHUNK = 0x8000;
	for (let i = 0; i < bytes.length; i += CHUNK) {
		binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
	}
	return btoa(binary);
}

// Decode standard base64 to bytes. Throws ToolError on invalid input rather
// than letting atob's DOMException leak out as an unhandled error.
export function base64ToBytes(b64: string): Uint8Array {
	let binary: string;
	try {
		binary = atob(b64);
	} catch {
		throw ToolError.validation("content_base64 is not valid base64");
	}
	const bytes = new Uint8Array(binary.length);
	for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
	return bytes;
}

// ── Tool implementations ──────────────────────────────────────────────────────

async function listFilesImpl(
	env: Env,
	args: { folder_path?: string; count?: number },
): Promise<unknown> {
	const count = args.count ?? 25;
	const basePath = args.folder_path
		? `/me/drive/root:/${encodeOneDrivePath(args.folder_path)}:/children`
		: `/me/drive/root/children`;

	const data = (await graphGet(env, basePath, {
		$select: "id,name,size,file,folder,lastModifiedDateTime,webUrl",
		$top: count,
	})) as { value: unknown[] };
	return sanitizeFileList(data.value);
}

async function getFileInfoImpl(env: Env, args: { item_path: string }): Promise<unknown> {
	const data = (await graphGet(env, `/me/drive/root:/${encodeOneDrivePath(args.item_path)}`, {
		$select: "id,name,size,file,folder,lastModifiedDateTime,webUrl",
	})) as {
		name?: string;
		size?: number;
		file?: { mimeType?: string };
		folder?: unknown;
		lastModifiedDateTime?: string;
	};
	if (data.folder) {
		throw ToolError.validation(`"${args.item_path}" is a folder, not a file.`);
	}
	const mimeType = data.file?.mimeType ?? null;
	return {
		success: true,
		name: data.name ?? null,
		size: data.size ?? null,
		mime_type: mimeType,
		is_video: typeof mimeType === "string" && mimeType.startsWith("video/"),
		last_modified: data.lastModifiedDateTime ?? null,
	};
}

async function shareFileImpl(
	env: Env,
	args: {
		item_path: string;
		link_type?: "view" | "edit";
		scope?: "anonymous" | "organization";
	},
): Promise<unknown> {
	const data = (await graphPost(
		env,
		`/me/drive/root:/${encodeOneDrivePath(args.item_path)}:/createLink`,
		{ type: args.link_type ?? "view", scope: args.scope ?? "anonymous" },
	)) as { link: { webUrl: string; type: string; scope: string } };

	return {
		success: true,
		url: data.link.webUrl,
		type: data.link.type,
		scope: data.link.scope,
	};
}

async function uploadOneDriveFileImpl(
	env: Env,
	args: {
		item_path: string;
		content_base64: string;
		content_type: string;
		conflict_behavior?: "rename" | "replace" | "fail";
	},
): Promise<unknown> {
	const mime = args.content_type.toLowerCase();
	if (!UPLOAD_ALLOWED_MIME_TYPES.has(mime)) {
		throw ToolError.validation(
			`MIME type not allowed: ${mime}. Allowed types: ${Array.from(
				UPLOAD_ALLOWED_MIME_TYPES,
			).join(", ")}`,
		);
	}

	const bytes = base64ToBytes(args.content_base64);
	if (bytes.byteLength === 0) {
		throw ToolError.validation("upload body is empty");
	}
	if (bytes.byteLength > MAX_UPLOAD_BYTES) {
		throw ToolError.validation(
			`File is ${(bytes.byteLength / 1024 / 1024).toFixed(1)} MB. Max is ${
				MAX_UPLOAD_BYTES / 1024 / 1024
			} MB for simple upload. Use upload_large_file instead.`,
		);
	}
	if (DANGEROUS_FILENAME_EXT.test(args.item_path)) {
		throw ToolError.validation(`Filename extension not allowed: ${args.item_path}`);
	}

	const encodedPath = encodeOneDrivePath(args.item_path);
	// `@microsoft.graph.conflictBehavior` is a Graph-defined query param.
	// URLSearchParams handles the URL-encoding of the `@` and `.` correctly.
	const query = new URLSearchParams({
		"@microsoft.graph.conflictBehavior": args.conflict_behavior ?? "replace",
	});
	const path = `/me/drive/root:/${encodedPath}:/content?${query.toString()}`;

	const data = (await graphPutBinary(env, path, bytes, mime)) as {
		id?: string;
		name?: string;
		size?: number;
		webUrl?: string;
		file?: { mimeType?: string };
		folder?: unknown;
		lastModifiedDateTime?: string;
	};

	return { success: true, file: sanitizeFileList([data])[0] };
}

// ── Document tools: large upload, download, PDF conversion, folders, copy ─────

// Upload sessions take chunks in multiples of 320 KiB; Graph recommends 5–10 MiB.
// 10 × 320 KiB ≈ 3.1 MiB keeps each Worker subrequest body small.
export const UPLOAD_CHUNK_BYTES = 10 * 320 * 1024;
// Largest file upload_large_file accepts. The bytes arrive as base64 in the tool
// call, so this is bounded by what a sensible MCP message can carry.
export const MAX_LARGE_UPLOAD_BYTES = 25 * 1024 * 1024;
// download_onedrive_file returns bytes inline as base64, which costs context.
export const DEFAULT_DOWNLOAD_BYTES = 5 * 1024 * 1024;
export const MAX_DOWNLOAD_BYTES = 10 * 1024 * 1024;

// Source formats Microsoft Graph can render to PDF (`/content?format=pdf`).
export const PDF_CONVERTIBLE_EXT = new Set([
	"doc", "docx", "dot", "dotx", "dotm", "odt", "rtf",
	"ppt", "pptx", "pps", "ppsx", "odp",
	"xls", "xlsx", "xlsm", "ods",
	"htm", "html", "md", "markdown", "eml", "msg", "epub", "tif", "tiff",
]);

const TEXT_EXT = /\.(txt|md|markdown|csv|json|xml|html?|vcf|ics)$/i;

type DriveItem = {
	id?: string;
	name?: string;
	size?: number;
	webUrl?: string;
	file?: { mimeType?: string };
	folder?: { childCount?: number };
	lastModifiedDateTime?: string;
	parentReference?: { driveId?: string; id?: string; path?: string };
};

type Conflict = "rename" | "replace" | "fail";

// "87 KB", "4.2 MB": readable at every size (a 50 KB limit shouldn't read "0.0 MB").
export function formatBytes(n: number): string {
	if (n < 1024 * 1024) return `${Math.max(1, Math.round(n / 1024))} KB`;
	return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

export function extensionOf(path: string): string {
	const name = path.split("/").pop() ?? "";
	const dot = name.lastIndexOf(".");
	return dot > 0 ? name.slice(dot + 1).toLowerCase() : "";
}

// "Clients/Acme/2026-09-22 Acme - 1-Page Roadmap.docx" → same path, ".pdf".
export function pdfPathFor(itemPath: string): string {
	const slash = itemPath.lastIndexOf("/");
	const dir = slash >= 0 ? itemPath.slice(0, slash + 1) : "";
	const name = itemPath.slice(slash + 1);
	const dot = name.lastIndexOf(".");
	return `${dir}${dot > 0 ? name.slice(0, dot) : name}.pdf`;
}

function assertSafeUploadTarget(itemPath: string, mime: string): void {
	if (!UPLOAD_ALLOWED_MIME_TYPES.has(mime)) {
		throw ToolError.validation(
			`MIME type not allowed: ${mime}. Allowed types: ${Array.from(UPLOAD_ALLOWED_MIME_TYPES).join(", ")}`,
		);
	}
	if (DANGEROUS_FILENAME_EXT.test(itemPath)) {
		throw ToolError.validation(`Filename extension not allowed: ${itemPath}`);
	}
}

function assertHttps(url: string, what: string): void {
	let parsed: URL;
	try {
		parsed = new URL(url);
	} catch {
		throw new ToolError({
			userMessage: `Outlook returned a malformed ${what}.`,
			internalMessage: `Unparseable ${what}: ${url}`,
			upstreamName: "Graph",
		});
	}
	if (parsed.protocol !== "https:") {
		throw new ToolError({
			userMessage: `Outlook returned an insecure ${what}.`,
			internalMessage: `Non-https ${what}: ${parsed.origin}`,
			upstreamName: "Graph",
		});
	}
}

// Upload through a Graph upload session, in UPLOAD_CHUNK_BYTES pieces. The
// session URL is pre-authenticated, so chunk PUTs carry no bearer token.
export async function uploadViaSession(
	env: Env,
	itemPath: string,
	bytes: Uint8Array,
	conflict: Conflict,
): Promise<DriveItem> {
	const session = (await graphPost(
		env,
		`/me/drive/root:/${encodeOneDrivePath(itemPath)}:/createUploadSession`,
		{ item: { "@microsoft.graph.conflictBehavior": conflict } },
	)) as { uploadUrl?: string };
	if (!session.uploadUrl) {
		throw new ToolError({
			userMessage: "Outlook didn't return an upload session.",
			internalMessage: "createUploadSession response had no uploadUrl",
			upstreamName: "Graph",
		});
	}
	assertHttps(session.uploadUrl, "upload URL");

	const total = bytes.byteLength;
	for (let start = 0; start < total; start += UPLOAD_CHUNK_BYTES) {
		const end = Math.min(start + UPLOAD_CHUNK_BYTES, total);
		const response = await fetch(session.uploadUrl, {
			method: "PUT",
			headers: { "Content-Range": `bytes ${start}-${end - 1}/${total}` },
			body: bytes.subarray(start, end),
		});
		if (!response.ok) {
			// Free the partial upload rather than leave it for Graph to expire.
			await fetch(session.uploadUrl, { method: "DELETE" }).catch(() => undefined);
			throw new ToolError({
				userMessage: `Outlook error ${response.status} while uploading ${itemPath} (at byte ${start} of ${total}).`,
				internalMessage: `Upload session chunk ${start}-${end - 1}/${total} → ${response.status}`,
				status: response.status,
				upstreamName: "Graph",
			});
		}
		if (end === total) return (await response.json()) as DriveItem;
	}
	throw ToolError.validation("upload body is empty");
}

// Simple upload up to 4 MB, upload session above that.
export async function uploadBytes(
	env: Env,
	itemPath: string,
	bytes: Uint8Array,
	mime: string,
	conflict: Conflict,
): Promise<DriveItem> {
	if (bytes.byteLength <= MAX_UPLOAD_BYTES) {
		const query = new URLSearchParams({ "@microsoft.graph.conflictBehavior": conflict });
		return (await graphPutBinary(
			env,
			`/me/drive/root:/${encodeOneDrivePath(itemPath)}:/content?${query.toString()}`,
			bytes,
			mime,
		)) as DriveItem;
	}
	return uploadViaSession(env, itemPath, bytes, conflict);
}

export async function uploadLargeFileImpl(
	env: Env,
	args: { item_path: string; content_base64: string; content_type: string; conflict_behavior?: Conflict },
): Promise<unknown> {
	const mime = args.content_type.toLowerCase();
	assertSafeUploadTarget(args.item_path, mime);
	const bytes = base64ToBytes(args.content_base64);
	if (bytes.byteLength === 0) throw ToolError.validation("upload body is empty");
	if (bytes.byteLength > MAX_LARGE_UPLOAD_BYTES) {
		throw ToolError.validation(
			`File is ${(bytes.byteLength / 1024 / 1024).toFixed(1)} MB. Max is ${MAX_LARGE_UPLOAD_BYTES / 1024 / 1024} MB.`,
		);
	}
	const item = await uploadViaSession(env, args.item_path, bytes, args.conflict_behavior ?? "replace");
	return { success: true, file: sanitizeFileList([item])[0] };
}

export async function downloadOneDriveFileImpl(
	env: Env,
	args: { item_path: string; max_bytes?: number },
): Promise<unknown> {
	const limit = args.max_bytes ?? DEFAULT_DOWNLOAD_BYTES;
	const encoded = encodeOneDrivePath(args.item_path);
	const info = (await graphGet(env, `/me/drive/root:/${encoded}`, {
		$select: "id,name,size,file,folder",
	})) as DriveItem;
	if (info.folder) {
		throw ToolError.validation(`"${args.item_path}" is a folder, not a file. Use list_files instead.`);
	}
	if ((info.size ?? 0) > limit) {
		throw ToolError.validation(
			`"${args.item_path}" is ${formatBytes(info.size ?? 0)}, over the ${formatBytes(limit)} limit. Raise max_bytes (up to ${formatBytes(MAX_DOWNLOAD_BYTES)}) or attach it to an email with onedrive_path instead.`,
		);
	}

	const response = await graphRequestRaw(env, `/me/drive/root:/${encoded}:/content`);
	const bytes = new Uint8Array(await response.arrayBuffer());
	if (bytes.byteLength > limit) {
		throw ToolError.validation(`"${args.item_path}" is larger than the ${formatBytes(limit)} limit.`);
	}

	const mime = info.file?.mimeType ?? "application/octet-stream";
	const isText =
		mime.startsWith("text/") ||
		mime === "application/json" ||
		mime === "application/xml" ||
		TEXT_EXT.test(args.item_path);
	return {
		success: true,
		name: info.name ?? null,
		size: bytes.byteLength,
		mime_type: mime,
		content_base64: bytesToBase64(bytes),
		...(isText ? { text: new TextDecoder().decode(bytes) } : {}),
	};
}

export async function convertToPdfImpl(
	env: Env,
	args: { item_path: string; output_path?: string; conflict_behavior?: Conflict },
): Promise<unknown> {
	const ext = extensionOf(args.item_path);
	if (!PDF_CONVERTIBLE_EXT.has(ext)) {
		throw ToolError.validation(
			`Can't convert ".${ext || "(none)"}" to PDF. Supported: ${Array.from(PDF_CONVERTIBLE_EXT).join(", ")}.`,
		);
	}
	const outputPath = args.output_path ?? pdfPathFor(args.item_path);
	if (extensionOf(outputPath) !== "pdf") {
		throw ToolError.validation(`output_path must end in .pdf (got "${outputPath}").`);
	}

	// Graph renders the document with Microsoft's own engine (same fonts and
	// layout as Word) and redirects to the PDF bytes.
	const response = await graphRequestRaw(
		env,
		`/me/drive/root:/${encodeOneDrivePath(args.item_path)}:/content?format=pdf`,
	);
	const bytes = new Uint8Array(await response.arrayBuffer());
	if (bytes.byteLength === 0) {
		throw new ToolError({
			userMessage: `OneDrive returned an empty PDF for ${args.item_path}.`,
			internalMessage: "format=pdf returned 0 bytes",
			upstreamName: "Graph",
		});
	}

	const item = await uploadBytes(env, outputPath, bytes, "application/pdf", args.conflict_behavior ?? "replace");
	return { success: true, source: args.item_path, pdf: sanitizeFileList([item])[0] };
}

export async function createFolderImpl(env: Env, args: { folder_path: string }): Promise<unknown> {
	const segments = args.folder_path.split("/").filter((s) => s.length > 0);
	if (segments.length === 0) throw ToolError.validation("folder_path is empty");

	const created: string[] = [];
	let parent = "";
	let item: DriveItem = {};
	for (const segment of segments) {
		const here = parent ? `${parent}/${segment}` : segment;
		try {
			item = (await graphGet(env, `/me/drive/root:/${encodeOneDrivePath(here)}`, {
				$select: "id,name,folder,webUrl,lastModifiedDateTime",
			})) as DriveItem;
			if (!item.folder) {
				throw ToolError.validation(`"${here}" already exists and is a file, not a folder.`);
			}
		} catch (e) {
			if (!(e instanceof ToolError) || e.status !== 404) throw e;
			const parentChildren = parent
				? `/me/drive/root:/${encodeOneDrivePath(parent)}:/children`
				: "/me/drive/root/children";
			item = (await graphPost(env, parentChildren, {
				name: segment,
				folder: {},
				"@microsoft.graph.conflictBehavior": "fail",
			})) as DriveItem;
			created.push(here);
		}
		parent = here;
	}

	return {
		success: true,
		folder: sanitizeFileList([item])[0],
		path: segments.join("/"),
		created,
		already_existed: created.length === 0,
	};
}

export async function copyItemImpl(
	env: Env,
	args: { source_path: string; destination_folder: string; new_name?: string; conflict_behavior?: Conflict },
	opts: { pollIntervalMs?: number; maxPolls?: number } = {},
): Promise<unknown> {
	const dest = (await graphGet(env, `/me/drive/root:/${encodeOneDrivePath(args.destination_folder)}`, {
		$select: "id,folder,parentReference",
	})) as DriveItem;
	if (!dest.folder) {
		throw ToolError.validation(`"${args.destination_folder}" is not a folder.`);
	}
	const driveId = dest.parentReference?.driveId;

	const query = new URLSearchParams({
		"@microsoft.graph.conflictBehavior": args.conflict_behavior ?? "rename",
	});
	const { location } = await graphPostAccepted(
		env,
		`/me/drive/root:/${encodeOneDrivePath(args.source_path)}:/copy?${query.toString()}`,
		{
			parentReference: { ...(driveId ? { driveId } : {}), id: dest.id },
			...(args.new_name ? { name: args.new_name } : {}),
		},
	);

	// Copy is asynchronous. The monitor URL is pre-authenticated; poll it briefly
	// so the common case (a small document) comes back complete.
	if (!location) return { success: true, status: "accepted" };
	assertHttps(location, "copy monitor URL");
	const interval = opts.pollIntervalMs ?? 1000;
	const maxPolls = opts.maxPolls ?? 10;
	for (let i = 0; i < maxPolls; i++) {
		if (interval > 0) await new Promise((r) => setTimeout(r, interval));
		const res = await fetch(location);
		if (!res.ok) break;
		const status = (await res.json()) as { status?: string; resourceId?: string; percentageComplete?: number };
		if (status.status === "completed" && status.resourceId) {
			const item = (await graphGet(env, `/me/drive/items/${encodeURIComponent(status.resourceId)}`, {
				$select: "id,name,size,file,folder,lastModifiedDateTime,webUrl",
			})) as DriveItem;
			return { success: true, status: "completed", file: sanitizeFileList([item])[0] };
		}
		if (status.status === "failed") {
			throw new ToolError({
				userMessage: `OneDrive couldn't copy ${args.source_path}.`,
				internalMessage: `copy monitor reported failed: ${JSON.stringify(status)}`,
				upstreamName: "Graph",
			});
		}
	}
	return {
		success: true,
		status: "in_progress",
		note: "The copy was accepted and is still running. Check the destination with list_files shortly.",
	};
}

// ── Tool definitions ──────────────────────────────────────────────────────────

export const filesTools = defineTools<Env>({
	list_files: {
		description: "List files and folders in OneDrive.",
		schema: z.object({
			folder_path: pathSchema.optional(),
			count: z.number().int().min(1).max(200).optional(),
		}),
		handler: (env, args) => listFilesImpl(env, args),
	},

	get_onedrive_file_info: {
		description:
			"Get metadata for a single OneDrive file by path (name, size in bytes, MIME type, whether it is a video, last-modified). Use this to check a file exists and inspect its type/size before downloading it — e.g. to confirm a video and decide whether it needs converting. Does NOT return the file contents.",
		schema: z.object({ item_path: pathSchema }),
		handler: (env, args) => getFileInfoImpl(env, args),
	},

	share_file: {
		description:
			'Create a shareable link for a file or folder in OneDrive. Default scope is "anonymous" — anyone with the link can access. Use "organization" to restrict to org members only.',
		schema: z.object({
			item_path: pathSchema,
			link_type: z.enum(["view", "edit"]).optional(),
			scope: z.enum(["anonymous", "organization"]).optional(),
		}),
		handler: (env, args) => shareFileImpl(env, args),
	},

	upload_onedrive_file: {
		description:
			'Upload a new file to OneDrive via Microsoft Graph simple upload. Pass file bytes as base64 in content_base64 and the MIME type in content_type. Max 4 MB. conflict_behavior defaults to "replace"; use "rename" to auto-suffix `(1)` on collision or "fail" to error. Returns the uploaded driveItem metadata (id, name, size, webUrl) so the caller can immediately share it via share_file — useful for any host-then-share workflow (e.g. generating a .vcf in your own folder and returning a download link).',
		schema: z.object({
			item_path: pathSchema,
			content_base64: z.string().min(1).max(8 * 1024 * 1024),
			content_type: z.string().min(1).max(128),
			conflict_behavior: z.enum(["rename", "replace", "fail"]).optional(),
		}),
		handler: (env, args) => uploadOneDriveFileImpl(env, args),
	},
	upload_large_file: {
		description:
			'Upload a file of any size up to 25 MB to OneDrive via a Graph upload session (chunked). Use this instead of upload_onedrive_file when the file is over 4 MB. Same inputs: item_path, content_base64, content_type, optional conflict_behavior (default "replace").',
		schema: z.object({
			item_path: pathSchema,
			content_base64: z.string().min(1).max(36 * 1024 * 1024),
			content_type: z.string().min(1).max(128),
			conflict_behavior: z.enum(["rename", "replace", "fail"]).optional(),
		}),
		handler: (env, args) => uploadLargeFileImpl(env, args),
	},

	download_onedrive_file: {
		description:
			"Download a OneDrive file's contents by path. Returns content_base64 plus, for text formats (txt, md, csv, json, xml, html, vcf, ics), a decoded `text` field. Default limit 5 MB (max_bytes up to 10 MB). To send a file to someone, attach it with onedrive_path on an email instead of downloading it.",
		schema: z.object({
			item_path: pathSchema,
			max_bytes: z.number().int().min(1).max(MAX_DOWNLOAD_BYTES).optional(),
		}),
		handler: (env, args) => downloadOneDriveFileImpl(env, args),
	},

	convert_to_pdf: {
		description:
			'Convert a OneDrive document (docx, pptx, xlsx, rtf, odt, html, md and more) to PDF using Microsoft\'s own renderer, so fonts and layout match Word, and save the PDF to OneDrive. output_path defaults to the same folder and name with a .pdf extension. conflict_behavior defaults to "replace". Returns the PDF\'s driveItem, ready to attach to an email with onedrive_path.',
		schema: z.object({
			item_path: pathSchema,
			output_path: pathSchema.optional(),
			conflict_behavior: z.enum(["rename", "replace", "fail"]).optional(),
		}),
		handler: (env, args) => convertToPdfImpl(env, args),
	},

	create_folder: {
		description:
			'Create a OneDrive folder path, including any missing parent folders (like mkdir -p). Safe to repeat: existing folders are left alone. Returns the folder plus which levels were created, e.g. "Rise Advisory/Clients/Acme/03 Roadmap".',
		schema: z.object({ folder_path: pathSchema }),
		handler: (env, args) => createFolderImpl(env, args),
	},

	copy_item: {
		description:
			'Copy a OneDrive file or folder into another folder, optionally under a new name. conflict_behavior defaults to "rename" (adds a suffix rather than overwrite); use "replace" to overwrite or "fail" to error. Waits briefly for the copy to finish and returns the new item; very large copies may report status "in_progress".',
		schema: z.object({
			source_path: pathSchema,
			destination_folder: pathSchema,
			new_name: z.string().min(1).max(255).regex(/^[^\\/:*?"<>|#%]+$/, "new_name can't contain \\ / : * ? \" < > | # %").optional(),
			conflict_behavior: z.enum(["rename", "replace", "fail"]).optional(),
		}),
		handler: (env, args) => copyItemImpl(env, args),
	},
});
