import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

// Mock the Graph client layer so we can assert exactly what the attachment tools
// request and where the bytes end up. Upload-session chunks go straight to
// `fetch` (pre-authenticated URL), so that's stubbed where needed.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphGetNextLink: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
	graphPostAccepted: vi.fn(),
}));

import { graphGet, graphPost, graphPutBinary, graphRequestRaw } from "../src/graph.js";
import {
	type AttachmentSummary,
	chooseAttachment,
	listEmailAttachmentsImpl,
	safeFileName,
	saveEmailAttachmentImpl,
} from "../src/tools/attachments.js";

const env = {} as never;
const MSG = "AAMkAD-test_id=";
const userSees = (re: RegExp) => expect.objectContaining({ userMessage: expect.stringMatching(re) });

const snapshotPdf = {
	"@odata.type": "#microsoft.graph.fileAttachment",
	id: "att-1",
	name: "rise-onboarding-harbourside-electrical.pdf",
	contentType: "application/pdf",
	size: 15115,
	isInline: false,
};
const logo = {
	"@odata.type": "#microsoft.graph.fileAttachment",
	id: "att-2",
	name: "image001.png",
	contentType: "image/png",
	size: 4200,
	isInline: true,
};

function routeAttachments(value: unknown[]) {
	vi.mocked(graphGet).mockResolvedValue({ value });
}

beforeEach(() => {
	vi.clearAllMocks();
});
afterEach(() => {
	vi.unstubAllGlobals();
});

describe("list_email_attachments", () => {
	it("lists name, type, size, inline flag and kind, without contents", async () => {
		routeAttachments([
			snapshotPdf,
			logo,
			{ "@odata.type": "#microsoft.graph.itemAttachment", id: "att-3", name: "Fwd: booking", size: 9000 },
			{ "@odata.type": "#microsoft.graph.referenceAttachment", id: "att-4", name: "Plan.docx", size: 0 },
		]);

		const out = await listEmailAttachmentsImpl(env, { message_id: MSG });

		expect(vi.mocked(graphGet).mock.calls[0]?.[1]).toBe(`/me/messages/${encodeURIComponent(MSG)}/attachments`);
		expect(vi.mocked(graphGet).mock.calls[0]?.[2]).toEqual({ $select: "id,name,contentType,size,isInline" });
		expect(out.attachments.map((a) => [a.name, a.kind, a.is_inline])).toEqual([
			["rise-onboarding-harbourside-electrical.pdf", "file", false],
			["image001.png", "file", true],
			["Fwd: booking", "item", false],
			["Plan.docx", "reference", false],
		]);
		expect(out.attachments[0]).not.toHaveProperty("contentBytes");
	});
});

describe("chooseAttachment", () => {
	const list: AttachmentSummary[] = [
		{ id: "a", name: "Snapshot.pdf", content_type: "application/pdf", size: 10, is_inline: false, kind: "file" },
		{ id: "b", name: "image001.png", content_type: "image/png", size: 10, is_inline: true, kind: "file" },
	];

	it("picks by id, by name case-insensitively, or the only non-inline file", () => {
		expect(chooseAttachment(list, { attachment_id: "b" }).id).toBe("b");
		expect(chooseAttachment(list, { attachment_name: "snapshot.PDF" }).id).toBe("a");
		expect(chooseAttachment(list, {}).id).toBe("a"); // the inline logo is ignored
	});

	it("never guesses between several files", () => {
		const two = [...list, { ...list[0]!, id: "c", name: "Financials.pdf" }];
		expect(() => chooseAttachment(two, {})).toThrow(ToolErrorLike(/2 attachments/));
		const dupes = [...list, { ...list[0]!, id: "d" }];
		expect(() => chooseAttachment(dupes, { attachment_name: "Snapshot.pdf" })).toThrow(ToolErrorLike(/More than one/));
		expect(() => chooseAttachment(list, { attachment_name: "missing.pdf" })).toThrow(ToolErrorLike(/No attachment called/));
		expect(() => chooseAttachment([list[1]!], {})).toThrow(ToolErrorLike(/no file attachments/));
	});
});

// ToolError.message is the internal detail; the user-facing text is userMessage.
function ToolErrorLike(re: RegExp) {
	return expect.objectContaining({ userMessage: expect.stringMatching(re) });
}

describe("safeFileName", () => {
	it("keeps ordinary names and cleans characters OneDrive or our paths reject", () => {
		expect(safeFileName("rise-onboarding-harbourside-electrical.pdf")).toBe("rise-onboarding-harbourside-electrical.pdf");
		expect(safeFileName("Q3: P&L #2.pdf")).toBe("Q3- P&L -2.pdf");
		expect(safeFileName("../../etc/passwd")).toBe("-..-etc-passwd");
		expect(safeFileName("  ...  ")).toBe("attachment");
	});
});

describe("save_email_attachment", () => {
	it("copies the chosen attachment into the destination folder, renaming on conflict by default", async () => {
		routeAttachments([snapshotPdf, logo]);
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(new Uint8Array(15115)));
		vi.mocked(graphPutBinary).mockResolvedValue({ id: "f1", name: snapshotPdf.name, size: 15115, file: { mimeType: "application/pdf" } });

		const out = (await saveEmailAttachmentImpl(env, {
			message_id: MSG,
			destination_folder: "Rise Advisory/Clients/Harbourside Electrical/02 Business Snapshot/",
		})) as { saved_to: string };

		expect(vi.mocked(graphRequestRaw).mock.calls[0]?.[1]).toBe(
			`/me/messages/${encodeURIComponent(MSG)}/attachments/att-1/$value`,
		);
		const [, path, body, mime] = vi.mocked(graphPutBinary).mock.calls[0]!;
		expect(path).toBe(
			"/me/drive/root:/Rise%20Advisory/Clients/Harbourside%20Electrical/02%20Business%20Snapshot/rise-onboarding-harbourside-electrical.pdf:/content?%40microsoft.graph.conflictBehavior=rename",
		);
		expect(mime).toBe("application/pdf");
		expect((body as Uint8Array).byteLength).toBe(15115);
		expect(out.saved_to).toBe("Rise Advisory/Clients/Harbourside Electrical/02 Business Snapshot/rise-onboarding-harbourside-electrical.pdf");
	});

	it("honours file_name and conflict_behavior", async () => {
		routeAttachments([snapshotPdf]);
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(new Uint8Array(10)));
		vi.mocked(graphPutBinary).mockResolvedValue({ id: "f" });
		await saveEmailAttachmentImpl(env, {
			message_id: MSG, destination_folder: "a", attachment_name: snapshotPdf.name,
			file_name: "2026-09-25 Harbourside Electrical - Business Snapshot.pdf", conflict_behavior: "replace",
		});
		expect(vi.mocked(graphPutBinary).mock.calls[0]?.[1]).toBe(
			"/me/drive/root:/a/2026-09-25%20Harbourside%20Electrical%20-%20Business%20Snapshot.pdf:/content?%40microsoft.graph.conflictBehavior=replace",
		);
	});

	it("uses an upload session for attachments over 4 MB", async () => {
		const size = 5 * 1024 * 1024;
		routeAttachments([{ ...snapshotPdf, size }]);
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(new Uint8Array(size)));
		vi.mocked(graphPost).mockResolvedValue({ uploadUrl: "https://tenant-my.sharepoint.com/upload/s1" });
		vi.stubGlobal("fetch", vi.fn(async (_u: string, init: RequestInit) => {
			const range = new Headers(init.headers).get("Content-Range")!;
			return range.endsWith(`${size - 1}/${size}`) ? Response.json({ id: "big" }, { status: 201 }) : new Response(null, { status: 202 });
		}));
		await saveEmailAttachmentImpl(env, { message_id: MSG, destination_folder: "a" });
		expect(graphPutBinary).not.toHaveBeenCalled();
		expect(vi.mocked(graphPost).mock.calls[0]?.[2]).toEqual({ item: { "@microsoft.graph.conflictBehavior": "rename" } });
	});

	it("refuses attached emails, cloud links, dangerous types and oversized files, before downloading", async () => {
		routeAttachments([{ "@odata.type": "#microsoft.graph.itemAttachment", id: "i", name: "Fwd: booking", size: 10 }]);
		await expect(saveEmailAttachmentImpl(env, { message_id: MSG, destination_folder: "a", attachment_id: "i" })).rejects.toMatchObject(userSees(/not a file/));
		routeAttachments([{ "@odata.type": "#microsoft.graph.referenceAttachment", id: "r", name: "Plan.docx", size: 0 }]);
		await expect(saveEmailAttachmentImpl(env, { message_id: MSG, destination_folder: "a", attachment_id: "r" })).rejects.toMatchObject(userSees(/link to a cloud file/));
		routeAttachments([{ ...snapshotPdf, name: "invoice.exe" }]);
		await expect(saveEmailAttachmentImpl(env, { message_id: MSG, destination_folder: "a" })).rejects.toMatchObject(userSees(/isn't allowed/));
		routeAttachments([{ ...snapshotPdf, size: 30 * 1024 * 1024 }]);
		await expect(saveEmailAttachmentImpl(env, { message_id: MSG, destination_folder: "a" })).rejects.toMatchObject(userSees(/limit is 25 MB/));
		expect(graphRequestRaw).not.toHaveBeenCalled();
	});
});
