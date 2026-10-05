import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { ToolError } from "@bashco/mcp-toolkit";

// Mock the Graph client layer so we can assert exactly what each document tool
// sends. Upload-session chunks and copy-monitor polls go straight to `fetch`
// (their URLs are pre-authenticated), so that's stubbed per test.
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

import {
	graphGet,
	graphPost,
	graphPostAccepted,
	graphPutBinary,
	graphRequestRaw,
} from "../src/graph.js";
import {
	UPLOAD_CHUNK_BYTES,
	bytesToBase64,
	convertToPdfImpl,
	copyItemImpl,
	createFolderImpl,
	downloadOneDriveFileImpl,
	extensionOf,
	formatBytes,
	pdfPathFor,
	uploadLargeFileImpl,
} from "../src/tools/files.js";

const env = {} as never;
const UPLOAD_URL = "https://tenant-my.sharepoint.com/upload/session-123";

function notFound(): ToolError {
	return new ToolError({ userMessage: "Outlook error 404: itemNotFound", internalMessage: "Graph 404", status: 404, upstreamName: "Graph" });
}

// ToolError.message is the internal detail; assert on what the user is shown.
const userSees = (re: RegExp) => expect.objectContaining({ userMessage: expect.stringMatching(re) });

function bytesOf(n: number): Uint8Array {
	const b = new Uint8Array(n);
	for (let i = 0; i < n; i++) b[i] = i % 251;
	return b;
}

beforeEach(() => {
	vi.clearAllMocks();
});
afterEach(() => {
	vi.unstubAllGlobals();
});

describe("path helpers", () => {
	it("swaps the extension for .pdf and keeps the folder", () => {
		expect(pdfPathFor("Rise Advisory/Clients/Acme/03 Roadmap/2026-09-22 Acme - 1-Page Roadmap.docx")).toBe(
			"Rise Advisory/Clients/Acme/03 Roadmap/2026-09-22 Acme - 1-Page Roadmap.pdf",
		);
		expect(pdfPathFor("report")).toBe("report.pdf");
		expect(pdfPathFor("v1.2 notes.md")).toBe("v1.2 notes.pdf");
	});

	it("reads the extension case-insensitively and ignores dotfiles", () => {
		expect(extensionOf("a/b/Plan.DOCX")).toBe("docx");
		expect(extensionOf("a/.hidden")).toBe("");
		expect(extensionOf("noext")).toBe("");
	});

	it("formats sizes readably at every scale", () => {
		expect(formatBytes(50_000)).toBe("49 KB");
		expect(formatBytes(87_929)).toBe("86 KB");
		expect(formatBytes(200)).toBe("1 KB");
		expect(formatBytes(5 * 1024 * 1024)).toBe("5.0 MB");
	});

	it("round-trips base64", () => {
		const b = bytesOf(100_000);
		const back = Uint8Array.from(atob(bytesToBase64(b)), (c) => c.charCodeAt(0));
		expect(back).toEqual(b);
	});
});

describe("convert_to_pdf", () => {
	it("asks Graph for format=pdf and saves the PDF next to the source", async () => {
		const pdf = bytesOf(2048);
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(pdf));
		vi.mocked(graphPutBinary).mockResolvedValue({ id: "pdf1", name: "Plan.pdf", size: 2048, file: { mimeType: "application/pdf" } });

		const out = (await convertToPdfImpl(env, { item_path: "Clients/Acme/Plan.docx" })) as { pdf: { name: string } };

		expect(vi.mocked(graphRequestRaw).mock.calls[0]?.[1]).toBe("/me/drive/root:/Clients/Acme/Plan.docx:/content?format=pdf");
		const [, path, body, mime] = vi.mocked(graphPutBinary).mock.calls[0]!;
		expect(path).toBe("/me/drive/root:/Clients/Acme/Plan.pdf:/content?%40microsoft.graph.conflictBehavior=replace");
		expect(mime).toBe("application/pdf");
		expect((body as Uint8Array).byteLength).toBe(2048);
		expect(out.pdf.name).toBe("Plan.pdf");
	});

	it("honours output_path and conflict_behavior", async () => {
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(bytesOf(10)));
		vi.mocked(graphPutBinary).mockResolvedValue({ id: "x" });
		await convertToPdfImpl(env, { item_path: "a/Plan.docx", output_path: "Sent to Client/Plan.pdf", conflict_behavior: "rename" });
		expect(vi.mocked(graphPutBinary).mock.calls[0]?.[1]).toBe(
			"/me/drive/root:/Sent%20to%20Client/Plan.pdf:/content?%40microsoft.graph.conflictBehavior=rename",
		);
	});

	it("rejects formats Graph can't render, and non-pdf output paths", async () => {
		await expect(convertToPdfImpl(env, { item_path: "a/already.pdf" })).rejects.toThrow(/Can't convert/);
		await expect(convertToPdfImpl(env, { item_path: "a/archive.zip" })).rejects.toThrow(/Can't convert/);
		await expect(convertToPdfImpl(env, { item_path: "a/Plan.docx", output_path: "a/Plan.docx" })).rejects.toThrow(/must end in \.pdf/);
		expect(graphRequestRaw).not.toHaveBeenCalled();
	});

	it("uses an upload session when the PDF is over 4 MB", async () => {
		const big = bytesOf(5 * 1024 * 1024);
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(big));
		vi.mocked(graphPost).mockResolvedValue({ uploadUrl: UPLOAD_URL });
		const ranges: string[] = [];
		vi.stubGlobal("fetch", vi.fn(async (_url: string, init: RequestInit) => {
			const range = new Headers(init.headers).get("Content-Range")!;
			ranges.push(range);
			const last = range.endsWith(`${big.byteLength - 1}/${big.byteLength}`);
			return last ? Response.json({ id: "big", name: "Plan.pdf" }, { status: 201 }) : new Response(null, { status: 202 });
		}));

		await convertToPdfImpl(env, { item_path: "a/Plan.pptx" });

		expect(graphPutBinary).not.toHaveBeenCalled();
		expect(vi.mocked(graphPost).mock.calls[0]?.[1]).toBe("/me/drive/root:/a/Plan.pdf:/createUploadSession");
		expect(ranges).toEqual([
			`bytes 0-${UPLOAD_CHUNK_BYTES - 1}/${big.byteLength}`,
			`bytes ${UPLOAD_CHUNK_BYTES}-${big.byteLength - 1}/${big.byteLength}`,
		]);
	});
});

describe("upload_large_file", () => {
	it("sends 320 KiB-aligned chunks with correct Content-Range headers", async () => {
		const size = 7 * 1024 * 1024;
		vi.mocked(graphPost).mockResolvedValue({ uploadUrl: UPLOAD_URL });
		const ranges: string[] = [];
		const fetchMock = vi.fn(async (_url: string, init: RequestInit) => {
			const range = new Headers(init.headers).get("Content-Range")!;
			ranges.push(range);
			expect(new Headers(init.headers).get("Authorization")).toBeNull();
			return range.endsWith(`${size - 1}/${size}`)
				? Response.json({ id: "f", name: "deck.pptx", size }, { status: 201 })
				: new Response(null, { status: 202 });
		});
		vi.stubGlobal("fetch", fetchMock);

		const out = (await uploadLargeFileImpl(env, {
			item_path: "Clients/Acme/deck.pptx",
			content_base64: bytesToBase64(bytesOf(size)),
			content_type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
		})) as { file: { name: string } };

		expect(UPLOAD_CHUNK_BYTES % (320 * 1024)).toBe(0);
		expect(ranges).toHaveLength(Math.ceil(size / UPLOAD_CHUNK_BYTES));
		expect(ranges[0]).toBe(`bytes 0-${UPLOAD_CHUNK_BYTES - 1}/${size}`);
		expect(ranges.at(-1)).toBe(`bytes ${2 * UPLOAD_CHUNK_BYTES}-${size - 1}/${size}`);
		expect(vi.mocked(graphPost).mock.calls[0]?.[2]).toEqual({ item: { "@microsoft.graph.conflictBehavior": "replace" } });
		expect(out.file.name).toBe("deck.pptx");
	});

	it("cancels the session when a chunk fails", async () => {
		vi.mocked(graphPost).mockResolvedValue({ uploadUrl: UPLOAD_URL });
		const methods: string[] = [];
		vi.stubGlobal("fetch", vi.fn(async (_url: string, init: RequestInit) => {
			methods.push(init.method ?? "GET");
			return init.method === "DELETE" ? new Response(null, { status: 204 }) : new Response("boom", { status: 500 });
		}));
		await expect(
			uploadLargeFileImpl(env, { item_path: "a/b.pdf", content_base64: bytesToBase64(bytesOf(10)), content_type: "application/pdf" }),
		).rejects.toMatchObject(userSees(/error 500/));
		expect(methods).toEqual(["PUT", "DELETE"]);
	});

	it("applies the same MIME and extension rules as upload_onedrive_file", async () => {
		await expect(
			uploadLargeFileImpl(env, { item_path: "a/tool.exe", content_base64: "AAAA", content_type: "application/pdf" }),
		).rejects.toThrow(/extension not allowed/);
		await expect(
			uploadLargeFileImpl(env, { item_path: "a/x.bin", content_base64: "AAAA", content_type: "application/x-msdownload" }),
		).rejects.toThrow(/MIME type not allowed/);
		expect(graphPost).not.toHaveBeenCalled();
	});

	it("refuses a non-https upload URL", async () => {
		vi.mocked(graphPost).mockResolvedValue({ uploadUrl: "http://example.com/upload" });
		vi.stubGlobal("fetch", vi.fn());
		await expect(
			uploadLargeFileImpl(env, { item_path: "a/b.pdf", content_base64: "AAAA", content_type: "application/pdf" }),
		).rejects.toMatchObject(userSees(/insecure/));
		expect(fetch).not.toHaveBeenCalled();
	});
});

describe("download_onedrive_file", () => {
	it("returns base64 and decoded text for text formats", async () => {
		const md = "# Rise\nKia ora";
		vi.mocked(graphGet).mockResolvedValue({ name: "notes.md", size: md.length, file: { mimeType: "text/markdown" } });
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(md));

		const out = (await downloadOneDriveFileImpl(env, { item_path: "Rise/notes.md" })) as Record<string, unknown>;

		expect(vi.mocked(graphRequestRaw).mock.calls[0]?.[1]).toBe("/me/drive/root:/Rise/notes.md:/content");
		expect(out.text).toBe(md);
		expect(atob(out.content_base64 as string)).toBe(md);
	});

	it("omits text for binary formats", async () => {
		vi.mocked(graphGet).mockResolvedValue({ name: "t.docx", size: 4, file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
		vi.mocked(graphRequestRaw).mockResolvedValue(new Response(bytesOf(4)));
		const out = (await downloadOneDriveFileImpl(env, { item_path: "t.docx" })) as Record<string, unknown>;
		expect(out.text).toBeUndefined();
		expect(out.size).toBe(4);
	});

	it("refuses folders and files over the limit before downloading", async () => {
		vi.mocked(graphGet).mockResolvedValueOnce({ name: "Clients", folder: { childCount: 3 } });
		await expect(downloadOneDriveFileImpl(env, { item_path: "Clients" })).rejects.toThrow(/is a folder/);
		vi.mocked(graphGet).mockResolvedValueOnce({ name: "big.pdf", size: 6 * 1024 * 1024, file: { mimeType: "application/pdf" } });
		await expect(downloadOneDriveFileImpl(env, { item_path: "big.pdf" })).rejects.toThrow(/is 6\.0 MB, over the 5\.0 MB limit/);
		expect(graphRequestRaw).not.toHaveBeenCalled();
	});
});

describe("create_folder", () => {
	it("creates only the missing levels, under the right parents", async () => {
		const existing = new Set(["Rise Advisory", "Rise Advisory/Clients"]);
		vi.mocked(graphGet).mockImplementation(async (_e: unknown, path: string) => {
			const p = decodeURIComponent(path.replace("/me/drive/root:/", ""));
			if (existing.has(p)) return { id: p, name: p.split("/").pop(), folder: { childCount: 0 } };
			throw notFound();
		});
		vi.mocked(graphPost).mockImplementation(async (_e: unknown, _p: string, body?: unknown) => ({
			id: "new",
			name: (body as { name: string }).name,
			folder: { childCount: 0 },
		}));

		const out = (await createFolderImpl(env, { folder_path: "Rise Advisory/Clients/Acme Plumbing/03 Roadmap/" })) as Record<string, unknown>;

		expect(out.created).toEqual(["Rise Advisory/Clients/Acme Plumbing", "Rise Advisory/Clients/Acme Plumbing/03 Roadmap"]);
		const posts = vi.mocked(graphPost).mock.calls.map((c) => [c[1], (c[2] as { name: string }).name]);
		expect(posts).toEqual([
			["/me/drive/root:/Rise%20Advisory/Clients:/children", "Acme Plumbing"],
			["/me/drive/root:/Rise%20Advisory/Clients/Acme%20Plumbing:/children", "03 Roadmap"],
		]);
		expect(vi.mocked(graphPost).mock.calls[0]?.[2]).toMatchObject({ folder: {}, "@microsoft.graph.conflictBehavior": "fail" });
	});

	it("creates a top-level folder under the drive root", async () => {
		vi.mocked(graphGet).mockRejectedValue(notFound());
		vi.mocked(graphPost).mockResolvedValue({ id: "r", name: "Rise Advisory", folder: {} });
		await createFolderImpl(env, { folder_path: "Rise Advisory" });
		expect(vi.mocked(graphPost).mock.calls[0]?.[1]).toBe("/me/drive/root/children");
	});

	it("reports already_existed and makes no changes when the path exists", async () => {
		vi.mocked(graphGet).mockResolvedValue({ id: "x", name: "x", folder: {} });
		const out = (await createFolderImpl(env, { folder_path: "a/b" })) as Record<string, unknown>;
		expect(out.already_existed).toBe(true);
		expect(graphPost).not.toHaveBeenCalled();
	});

	it("stops when a file is in the way, and passes other errors through", async () => {
		vi.mocked(graphGet).mockResolvedValueOnce({ id: "f", name: "a", file: { mimeType: "text/plain" } });
		await expect(createFolderImpl(env, { folder_path: "a/b" })).rejects.toThrow(/is a file/);
		vi.mocked(graphGet).mockRejectedValueOnce(new ToolError({ userMessage: "Outlook error 403", internalMessage: "", status: 403 }));
		await expect(createFolderImpl(env, { folder_path: "a" })).rejects.toMatchObject({ status: 403 });
		expect(graphPost).not.toHaveBeenCalled();
	});
});

describe("copy_item", () => {
	const MONITOR = "https://tenant-my.sharepoint.com/_api/v2.0/monitor/abc";

	it("copies into the destination folder by id and returns the finished item", async () => {
		vi.mocked(graphGet).mockImplementation(async (_e: unknown, path: string) => {
			if (path === "/me/drive/root:/Clients/Acme/Sent%20to%20Client") {
				return { id: "dest-id", folder: {}, parentReference: { driveId: "drive-1" } };
			}
			if (path === "/me/drive/items/new-id") return { id: "new-id", name: "Plan.pdf", size: 10, file: { mimeType: "application/pdf" } };
			throw notFound();
		});
		vi.mocked(graphPostAccepted).mockResolvedValue({ status: 202, location: MONITOR });
		const polls = [{ status: "inProgress", percentageComplete: 50 }, { status: "completed", resourceId: "new-id" }];
		vi.stubGlobal("fetch", vi.fn(async () => Response.json(polls.shift())));

		const out = (await copyItemImpl(
			env,
			{ source_path: "Clients/Acme/03 Roadmap/Plan.pdf", destination_folder: "Clients/Acme/Sent to Client" },
			{ pollIntervalMs: 0 },
		)) as { status: string; file: { name: string } };

		const [, path, body] = vi.mocked(graphPostAccepted).mock.calls[0]!;
		expect(path).toBe("/me/drive/root:/Clients/Acme/03%20Roadmap/Plan.pdf:/copy?%40microsoft.graph.conflictBehavior=rename");
		expect(body).toEqual({ parentReference: { driveId: "drive-1", id: "dest-id" } });
		expect(out.status).toBe("completed");
		expect(out.file.name).toBe("Plan.pdf");
	});

	it("passes new_name and reports in_progress when the copy is still running", async () => {
		vi.mocked(graphGet).mockResolvedValue({ id: "d", folder: {}, parentReference: { driveId: "drv" } });
		vi.mocked(graphPostAccepted).mockResolvedValue({ status: 202, location: MONITOR });
		vi.stubGlobal("fetch", vi.fn(async () => Response.json({ status: "inProgress" })));

		const out = (await copyItemImpl(
			env,
			{ source_path: "a.pdf", destination_folder: "b", new_name: "a (sent).pdf", conflict_behavior: "fail" },
			{ pollIntervalMs: 0, maxPolls: 2 },
		)) as { status: string };

		expect(vi.mocked(graphPostAccepted).mock.calls[0]?.[2]).toMatchObject({ name: "a (sent).pdf" });
		expect(vi.mocked(graphPostAccepted).mock.calls[0]?.[1]).toContain("conflictBehavior=fail");
		expect(out.status).toBe("in_progress");
	});

	it("refuses a destination that isn't a folder, and surfaces a failed copy", async () => {
		vi.mocked(graphGet).mockResolvedValueOnce({ id: "f", file: { mimeType: "application/pdf" } });
		await expect(copyItemImpl(env, { source_path: "a.pdf", destination_folder: "b.pdf" })).rejects.toThrow(/not a folder/);
		expect(graphPostAccepted).not.toHaveBeenCalled();

		vi.mocked(graphGet).mockResolvedValue({ id: "d", folder: {}, parentReference: { driveId: "drv" } });
		vi.mocked(graphPostAccepted).mockResolvedValue({ status: 202, location: MONITOR });
		vi.stubGlobal("fetch", vi.fn(async () => Response.json({ status: "failed" })));
		await expect(copyItemImpl(env, { source_path: "a.pdf", destination_folder: "b" }, { pollIntervalMs: 0 })).rejects.toMatchObject(userSees(/couldn't copy/));
	});
});
