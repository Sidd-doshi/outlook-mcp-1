import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the Graph client layer so we can assert exactly what complete_task puts
// on the wire. Microsoft To Do rejects any OData query option on these
// endpoints with `400 invalidRequest` / `RequestBroker--ParseUri`, so "no query
// string" is a correctness property here, not a style preference.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphGetNextLink: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet, graphGetNextLink, graphPatch, graphPost } from "../src/graph.js";
import {
	completeTaskImpl,
	createTaskListImpl,
	listTasksImpl,
	matchTasks,
} from "../src/tools/tasks.js";

const env = {} as never;

const DEFAULT_LIST = "list-default";
const FLAGGED_LIST = "list-flagged";

const lists = [
	{ id: DEFAULT_LIST, displayName: "Tasks", wellknownListName: "defaultList" },
	{ id: FLAGGED_LIST, displayName: "Flagged Emails", wellknownListName: "none" },
];

function task(id: string, title: string, status = "notStarted") {
	return { id, title, status, dueDateTime: { dateTime: "2026-09-10T00:00:00" } };
}

// Routes list enumeration and each list's tasks.
function routeTasks(byList: Record<string, unknown[]>) {
	vi.mocked(graphGet).mockImplementation(async (_e: unknown, path: string): Promise<unknown> => {
		if (path === "/me/todo/lists") return { value: lists };
		const match = path.match(/^\/me\/todo\/lists\/([^/]+)\/tasks$/);
		if (match) return { value: byList[match[1] as string] ?? [] };
		return { value: [] };
	});
}

function taskRequestPaths(): string[] {
	return vi
		.mocked(graphGet)
		.mock.calls.filter(c => (c[1] as string).endsWith("/tasks"))
		.map(c => c[1] as string);
}

beforeEach(() => {
	vi.clearAllMocks();
	vi.mocked(graphGetNextLink).mockResolvedValue({ value: [] });
	vi.mocked(graphPatch).mockResolvedValue({ id: "t1", status: "completed" });
	vi.mocked(graphPost).mockResolvedValue({ id: "list-new", displayName: "Renovation" });
});

describe("matchTasks", () => {
	const list = { id: DEFAULT_LIST, name: "Tasks" };
	const rows = [
		task("t1", "Send the Rise invoice"),
		task("t2", "Chase the Rise invoice payment"),
		task("t3", "Book flights"),
		task("t4", "Send the old invoice", "completed"),
	];

	it("matches a title as a case-insensitive substring", () => {
		expect(matchTasks(rows, list, { title: "book FLIGHTS" }).map(t => t.id)).toEqual(["t3"]);
	});

	it("returns every candidate rather than guessing between them", () => {
		expect(matchTasks(rows, list, { title: "invoice" }).map(t => t.id)).toEqual(["t1", "t2"]);
	});

	// A finished task with a similar name must not make a live one ambiguous.
	it("ignores completed tasks when matching by title", () => {
		expect(matchTasks(rows, list, { title: "old invoice" })).toEqual([]);
	});

	it("matches an id regardless of status, so an already-done task is reportable", () => {
		expect(matchTasks(rows, list, { task_id: "t4" }).map(t => t.status)).toEqual(["completed"]);
	});

	it("carries the list along, since a task id alone is not addressable", () => {
		const [hit] = matchTasks(rows, list, { task_id: "t1" });
		expect(hit?.list_id).toBe(DEFAULT_LIST);
		expect(hit?.list_name).toBe("Tasks");
	});
});

describe("complete_task", () => {
	it("PATCHes the task to completed with no query string", async () => {
		routeTasks({ [DEFAULT_LIST]: [task("t1", "Send the Rise invoice")] });

		const result = (await completeTaskImpl(env, { title: "rise invoice" })) as {
			success: boolean;
		};

		expect(result.success).toBe(true);
		expect(graphPatch).toHaveBeenCalledTimes(1);
		const [, path, body] = vi.mocked(graphPatch).mock.calls[0] as [unknown, string, unknown];
		expect(path).toBe(`/me/todo/lists/${DEFAULT_LIST}/tasks/t1`);
		expect(path).not.toContain("?");
		expect(body).toEqual({ status: "completed" });
	});

	// Deliberately the opposite default to create_task: completing has to find
	// something that already exists, and it shouldn't miss because the task was
	// filed under a non-default list.
	it("searches every list when no list_id is given", async () => {
		routeTasks({
			[DEFAULT_LIST]: [task("t1", "Book flights")],
			[FLAGGED_LIST]: [task("t9", "Reply to Bruce")],
		});

		await completeTaskImpl(env, { title: "reply to bruce" });

		expect(taskRequestPaths()).toHaveLength(2);
		const [, path] = vi.mocked(graphPatch).mock.calls[0] as [unknown, string, unknown];
		expect(path).toBe(`/me/todo/lists/${FLAGGED_LIST}/tasks/t9`);
	});

	it("stays inside the given list when list_id is supplied", async () => {
		routeTasks({
			[DEFAULT_LIST]: [task("t1", "Book flights")],
			[FLAGGED_LIST]: [task("t9", "Book the venue")],
		});

		await completeTaskImpl(env, { title: "book", list_id: DEFAULT_LIST });

		expect(taskRequestPaths()).toEqual([`/me/todo/lists/${DEFAULT_LIST}/tasks`]);
	});

	// Ticking off the wrong task is silent and annoying to undo, so an
	// ambiguous title must complete nothing at all.
	it("completes nothing and returns the candidates when a title is ambiguous", async () => {
		routeTasks({
			[DEFAULT_LIST]: [
				task("t1", "Send the Rise invoice"),
				task("t2", "Chase the Rise invoice payment"),
			],
		});

		const result = (await completeTaskImpl(env, { title: "invoice" })) as {
			success: boolean;
			reason: string;
			matches: Array<{ id: string }>;
		};

		expect(graphPatch).not.toHaveBeenCalled();
		expect(result.success).toBe(false);
		expect(result.reason).toBe("ambiguous");
		expect(result.matches.map(m => m.id)).toEqual(["t1", "t2"]);
	});

	it("fails loudly when nothing matches", async () => {
		routeTasks({ [DEFAULT_LIST]: [task("t1", "Book flights")] });

		await expect(completeTaskImpl(env, { title: "nonexistent" })).rejects.toThrow(
			/No open task matching/,
		);
		expect(graphPatch).not.toHaveBeenCalled();
	});

	it("reports an already-completed task without a redundant write", async () => {
		routeTasks({ [DEFAULT_LIST]: [task("t1", "Send the Rise invoice", "completed")] });

		const result = (await completeTaskImpl(env, { task_id: "t1" })) as {
			success: boolean;
			already_completed: boolean;
		};

		expect(result.success).toBe(true);
		expect(result.already_completed).toBe(true);
		expect(graphPatch).not.toHaveBeenCalled();
	});

	it("accepts an exact task_id", async () => {
		routeTasks({ [FLAGGED_LIST]: [task("t9", "Reply to Bruce")] });

		await completeTaskImpl(env, { task_id: "t9" });

		const [, path] = vi.mocked(graphPatch).mock.calls[0] as [unknown, string, unknown];
		expect(path).toBe(`/me/todo/lists/${FLAGGED_LIST}/tasks/t9`);
	});
});

describe("list_tasks", () => {
	// Without this, a task id read from list_tasks can't be turned into a
	// PATCH path, because every To Do write is scoped to its list.
	it("stamps each task with the list it came from", async () => {
		routeTasks({ [DEFAULT_LIST]: [task("t1", "Book flights")] });

		const result = (await listTasksImpl(env, {})) as Array<{ list_id: string }>;

		expect(result[0]?.list_id).toBe(DEFAULT_LIST);
	});
});

describe("create_task_list", () => {
	it("creates a list and returns it", async () => {
		routeTasks({});

		const result = (await createTaskListImpl(env, { name: "Renovation" })) as {
			created: boolean;
			list: { id: string; name: string };
		};

		const [, path, body] = vi.mocked(graphPost).mock.calls[0] as [unknown, string, unknown];
		expect(path).toBe("/me/todo/lists");
		expect(path).not.toContain("?");
		expect(body).toEqual({ displayName: "Renovation" });
		expect(result.created).toBe(true);
		expect(result.list.id).toBe("list-new");
	});

	// Two lists both called "Work" is worse than not creating one: tasks scatter
	// across lists that look identical in the To Do UI.
	it("returns the existing list instead of duplicating a name", async () => {
		routeTasks({});

		const result = (await createTaskListImpl(env, { name: "  flagged emails  " })) as {
			created: boolean;
			list: { id: string };
		};

		expect(graphPost).not.toHaveBeenCalled();
		expect(result.created).toBe(false);
		expect(result.list.id).toBe(FLAGGED_LIST);
	});

	it("rejects a name that is only whitespace", async () => {
		routeTasks({});

		await expect(createTaskListImpl(env, { name: "   " })).rejects.toThrow(/cannot be empty/);
		expect(graphPost).not.toHaveBeenCalled();
	});
});
