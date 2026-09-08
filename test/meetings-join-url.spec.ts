import { describe, it, expect, vi, beforeEach } from "vitest";
import { ToolError } from "@bashco/mcp-toolkit";

// Mock the Graph client layer so we can assert exactly which body field the
// meeting tools ask Graph for. The outage these tests guard was invisible to
// every unit test in the repo: the join-URL fallback read `bodyPreview`, which
// Graph truncates at 255 characters, so any meeting whose body opened with the
// organiser's own note lost its URL mid-string and was dropped as if it were
// never a Teams meeting. Meetings with an agenda written on them are exactly
// the ones with recordings worth finding.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphGetNextLink: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet, graphGetNextLink } from "../src/graph.js";
import {
	extractJoinUrlFromBody,
	listRecentMeetingRecordingsImpl,
	meetingsTools,
} from "../src/tools/meetings.js";

const env = {} as never;

const SHORT_JOIN_URL = "https://teams.microsoft.com/meet/410801608411240?p=YoKg7JV7tghojEyMBP";

// A real Teams invite body: the organiser's note first, then the Teams block.
// Comfortably past the 255-character preview ceiling.
// Sized so the 255-character cut below lands *inside* the join URL, which is
// what Graph actually returned for the meeting that went missing.
const NOTE = "Hey Ashish, I have booked us at The Crate in Albany in the Latte room. "
	.repeat(4)
	.slice(0, 198);
const FULL_BODY =
	`<html><body><p>${NOTE}</p>` +
	`<div>________________________________</div>` +
	`<div>Microsoft Teams meeting</div>` +
	`<div>Join: <a href="${SHORT_JOIN_URL}">Click here to join the meeting</a></div>` +
	`</body></html>`;

// What Graph actually hands back in `bodyPreview` for that event — 255 chars,
// severed mid-URL. This exact string is why the tool returned nothing.
const TRUNCATED_PREVIEW =
	`${NOTE}\r\n________\r\nMicrosoft Teams meeting\r\nJoin: ${SHORT_JOIN_URL}`.slice(0, 255);

beforeEach(() => {
	vi.clearAllMocks();
	vi.mocked(graphGetNextLink).mockResolvedValue({ value: [] });
});

describe("extractJoinUrlFromBody", () => {
	it("finds a join URL sitting past the 255-character preview ceiling", () => {
		expect(FULL_BODY.length).toBeGreaterThan(255);
		expect(extractJoinUrlFromBody(FULL_BODY)).toBe(SHORT_JOIN_URL);
	});

	// The bug, stated as a test: this input is what the old code was fed.
	it("returns null for a preview severed mid-URL rather than a broken match", () => {
		expect(TRUNCATED_PREVIEW).toMatch(/https:\/\/teams\.$/);
		expect(extractJoinUrlFromBody(TRUNCATED_PREVIEW)).toBeNull();
	});

	it("decodes entities so a long-form context URL survives as one URL", () => {
		const longForm =
			"https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%22Tid%22%3a%22t%22%7d";
		const html = `<a href="${longForm.replace(/&/g, "&amp;")}&amp;anon=true">Join</a>`;

		expect(extractJoinUrlFromBody(html)).toBe(`${longForm}&anon=true`);
	});

	it("stops at the href delimiter instead of swallowing the rest of the tag", () => {
		expect(extractJoinUrlFromBody(`<a href="${SHORT_JOIN_URL}" title="Join">x</a>`)).toBe(
			SHORT_JOIN_URL,
		);
	});

	it("drops trailing prose punctuation", () => {
		expect(extractJoinUrlFromBody(`Join at ${SHORT_JOIN_URL}.`)).toBe(SHORT_JOIN_URL);
	});

	it("ignores non-meeting Teams links", () => {
		expect(
			extractJoinUrlFromBody('<a href="https://teams.microsoft.com/meetingOptions">Options</a>'),
		).toBeNull();
	});
});

describe("resolving a meeting from a calendar event", () => {
	function routeSingleEvent(event: unknown) {
		vi.mocked(graphGet).mockImplementation(
			async (_e: unknown, path: string): Promise<unknown> => {
				if (path.startsWith("/me/events/")) return event;
				if (path === "/me/onlineMeetings") return { value: [{ id: "om-1" }] };
				if (path === "/me/onlineMeetings/om-1") return { id: "om-1", subject: "Discovery" };
				return { value: [] };
			},
		);
	}

	function eventSelect(): string {
		const call = vi
			.mocked(graphGet)
			.mock.calls.find(c => (c[1] as string).startsWith("/me/events/"));
		return (call?.[2] as Record<string, string>)?.$select ?? "";
	}

	// bodyPreview cannot be trusted for a single event, and a single event is
	// one payload — there is no reason to ask for the truncated form.
	it("asks Graph for the full body, not the preview", async () => {
		routeSingleEvent({ isOnlineMeeting: true, body: { content: FULL_BODY } });

		await meetingsTools.dispatch(env, "find_online_meeting", { calendar_event_id: "e1" });

		expect(eventSelect()).toContain("body");
		expect(eventSelect()).not.toContain("bodyPreview");
	});

	it("resolves an event whose join URL is only in the long body", async () => {
		routeSingleEvent({ isOnlineMeeting: true, body: { content: FULL_BODY } });

		const result = (await meetingsTools.dispatch(env, "find_online_meeting", {
			calendar_event_id: "e1",
		})) as { id: string };

		expect(result.id).toBe("om-1");

		const filter = vi
			.mocked(graphGet)
			.mock.calls.find(c => c[1] === "/me/onlineMeetings")?.[2] as Record<string, string>;
		expect(filter.$filter).toContain(SHORT_JOIN_URL);
	});

	it("still prefers the structured field when Graph populates it", async () => {
		routeSingleEvent({
			isOnlineMeeting: true,
			onlineMeetingUrl: SHORT_JOIN_URL,
			body: { content: "<p>unrelated</p>" },
		});

		await meetingsTools.dispatch(env, "find_online_meeting", { calendar_event_id: "e1" });

		const filter = vi
			.mocked(graphGet)
			.mock.calls.find(c => c[1] === "/me/onlineMeetings")?.[2] as Record<string, string>;
		expect(filter.$filter).toContain(SHORT_JOIN_URL);
	});

	it("explains itself when no join URL can be found anywhere", async () => {
		routeSingleEvent({ isOnlineMeeting: true, body: { content: "<p>no link here</p>" } });

		await expect(
			meetingsTools.dispatch(env, "find_online_meeting", { calendar_event_id: "e1" }),
		).rejects.toThrow(/no joinUrl could be found/);
	});
});

describe("list_recent_meeting_recordings — body refetch", () => {
	// The listing keeps selecting the cheap preview; only candidates the preview
	// fails on cost a second request.
	function routeDiscovery(events: unknown[], opts: { bodies?: Record<string, string> } = {}) {
		vi.mocked(graphGet).mockImplementation(
			async (_e: unknown, path: string): Promise<unknown> => {
				if (path === "/me/events") return { value: events };
				if (path.startsWith("/me/events/")) {
					const id = path.slice("/me/events/".length);
					const content = opts.bodies?.[id];
					return content ? { body: { content } } : {};
				}
				if (path === "/me/onlineMeetings") return { value: [{ id: "om-1" }] };
				if (path.endsWith("/recordings")) {
					return { value: [{ id: "rec-1", createdDateTime: "2026-09-04T03:41:00Z" }] };
				}
				if (path.endsWith("/transcripts")) return { value: [] };
				return { value: [] };
			},
		);
	}

	// A factory, not a shared const: these objects are handed straight to the
	// impl, and one test reusing another's mutated fixture is exactly the kind
	// of false pass worth designing out.
	function truncatedEvent(id = "ev-discovery") {
		return {
			id,
			subject: "Discovery & Basics of Investing",
			start: { dateTime: "2026-09-04T01:00:00.0000000" },
			end: { dateTime: "2026-09-04T03:00:00.0000000" },
			isOnlineMeeting: true,
			onlineMeetingUrl: "",
			bodyPreview: TRUNCATED_PREVIEW,
			organizer: { emailAddress: { name: "Bashar Basheer" } },
		};
	}

	function bodyRefetches(): string[] {
		return vi
			.mocked(graphGet)
			.mock.calls.filter(c => (c[1] as string).startsWith("/me/events/"))
			.map(c => c[1] as string);
	}

	it("keeps the listing on bodyPreview so 200 events stay one cheap page", async () => {
		routeDiscovery([]);

		await listRecentMeetingRecordingsImpl(env, {});

		const query = vi.mocked(graphGet).mock.calls.find(c => c[1] === "/me/events")?.[2] as Record<
			string,
			string
		>;
		expect(query.$select).toContain("bodyPreview");
		expect(query.$select).not.toMatch(/,body(,|$)/);
	});

	// The end-to-end regression: this meeting was recorded, and the tool
	// reported an empty calendar.
	it("recovers a meeting whose preview truncated the join URL", async () => {
		routeDiscovery([truncatedEvent()], { bodies: { "ev-discovery": FULL_BODY } });

		const result = (await listRecentMeetingRecordingsImpl(env, { within_days: 30 })) as {
			count: number;
			recordings: Array<{ subject: string; recording_count: number }>;
		};

		expect(bodyRefetches()).toEqual(["/me/events/ev-discovery"]);
		expect(result.count).toBe(1);
		expect(result.recordings[0]?.subject).toBe("Discovery & Basics of Investing");
		expect(result.recordings[0]?.recording_count).toBe(1);
	});

	it("does not refetch bodies for events the preview already resolved", async () => {
		routeDiscovery([
			{
				...truncatedEvent("ev-short"),
				bodyPreview: `Microsoft Teams meeting\r\nJoin: ${SHORT_JOIN_URL}`,
			},
		]);

		await listRecentMeetingRecordingsImpl(env, {});

		expect(bodyRefetches()).toEqual([]);
	});

	it("reports why candidates were dropped instead of failing silently", async () => {
		routeDiscovery([truncatedEvent("ev-hopeless")]);

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			count: number;
			skipped: Record<string, number>;
		};

		expect(result.count).toBe(0);
		expect(result.skipped.no_join_url).toBe(1);
	});

	it("separates a quiet calendar from a broken one", async () => {
		vi.mocked(graphGet).mockImplementation(
			async (_e: unknown, path: string): Promise<unknown> => {
				if (path === "/me/events") {
					return {
						value: [
							{
								...truncatedEvent("ev-nocontent"),
								bodyPreview: `Microsoft Teams meeting\r\nJoin: ${SHORT_JOIN_URL}`,
							},
						],
					};
				}
				if (path === "/me/onlineMeetings") return { value: [{ id: "om-1" }] };
				return { value: [] };
			},
		);

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			skipped: Record<string, number>;
		};

		// Nobody hit record — not a resolution failure.
		expect(result.skipped.no_content).toBe(1);
		expect(result.skipped.no_join_url).toBe(0);
	});

	it("caps how many bodies one call will refetch", async () => {
		const many = Array.from({ length: 40 }, (_, i) => truncatedEvent(`ev-${i}`));
		routeDiscovery(many);

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			body_fetch_limit_reached?: number;
		};

		expect(bodyRefetches()).toHaveLength(25);
		expect(result.body_fetch_limit_reached).toBe(15);
	});

	// Meetings organised in another tenant resolve fine and then 403 on their
	// artifacts. That is Graph's cross-tenant boundary, not a fault, and it
	// recurs on every single call — so it must not read as an error.
	it("counts a cross-tenant 403 as forbidden, not as an error", async () => {
		vi.mocked(graphGet).mockImplementation(
			async (_e: unknown, path: string): Promise<unknown> => {
				if (path === "/me/events") {
					return {
						value: [
							{
								...truncatedEvent("ev-external"),
								bodyPreview: `Microsoft Teams meeting\r\nJoin: ${SHORT_JOIN_URL}`,
							},
						],
					};
				}
				if (path === "/me/onlineMeetings") return { value: [{ id: "om-external" }] };
				if (path.endsWith("/recordings") || path.endsWith("/transcripts")) {
					throw new ToolError({
						userMessage: "Outlook error 403: Forbidden",
						internalMessage: "Graph 403",
						status: 403,
						upstreamName: "Graph",
					});
				}
				return { value: [] };
			},
		);

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			skipped: Record<string, number>;
		};

		expect(result.skipped.forbidden).toBe(1);
		expect(result.skipped.error).toBe(0);
	});

	it("still reports an unexpected failure as an error", async () => {
		vi.mocked(graphGet).mockImplementation(
			async (_e: unknown, path: string): Promise<unknown> => {
				if (path === "/me/events") {
					return {
						value: [
							{
								...truncatedEvent("ev-broken"),
								bodyPreview: `Microsoft Teams meeting\r\nJoin: ${SHORT_JOIN_URL}`,
							},
						],
					};
				}
				if (path === "/me/onlineMeetings") return { value: [{ id: "om-1" }] };
				if (path.endsWith("/recordings") || path.endsWith("/transcripts")) {
					throw new Error("connection reset");
				}
				return { value: [] };
			},
		);

		const result = (await listRecentMeetingRecordingsImpl(env, {})) as {
			skipped: Record<string, number>;
		};

		expect(result.skipped.error).toBe(1);
		expect(result.skipped.forbidden).toBe(0);
	});
});
