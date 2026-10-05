import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import {
	graphDelete,
	graphGet,
	graphGetNextLink,
	graphPatch,
	graphPost,
} from "../graph.js";
import { sanitizeTaskList, sanitizeTaskLists } from "../sanitize.js";
import type { Env } from "../types.js";

const taskListIdSchema = z.string().min(1).max(512);

// ── Why these calls carry no query string ─────────────────────────────────────
// Microsoft To Do is the one Graph workload here that refuses OData query
// options on its collection endpoints. `/me/todo/lists?$select=...` answers
// `400 invalidRequest` with an `innerError.code` of `RequestBroker--ParseUri`
// and a message of just "Invalid request" — it names neither the option nor a
// property, so it reads like a malformed URL rather than an unsupported feature.
//
// It is neither. Every property these tools used to select is real and spelled
// as the todoTaskList and todoTask resources document it, and the rejection
// survives sending `$select` literally instead of percent-encoded as `%24select`
// (the encoding Graph's Outlook backends accept for mail, calendar, contacts and
// files). The To Do backend just doesn't implement the option; Microsoft's own
// reference hedges with "supports some of the OData query parameters" and
// declines to say which.
//
// That took out both read paths from the day they were written. create_task
// escaped it by accident: resolving the default list is a bare GET and the
// create itself is a POST, so neither carries a query string — which is exactly
// the shape proven to work, and the shape everything below now uses.
//
// The sanitisers already project down to the fields we return, so dropping
// $select costs response size and nothing else. Status filtering moves into JS.

// Graph pages To Do collections with `@odata.nextLink`. Without a server-side
// `$filter` the completed tasks now take up page slots, so an open task can sit
// behind a page boundary — follow the link rather than truncating at page one.
const MAX_PAGES = 10;

interface GraphPage {
	value?: unknown[];
	"@odata.nextLink"?: string;
}

async function collectPages(env: Env, path: string): Promise<unknown[]> {
	let page = (await graphGet(env, path)) as GraphPage;
	const items: unknown[] = [...(page.value ?? [])];

	for (let fetched = 1; fetched < MAX_PAGES; fetched += 1) {
		const next = page["@odata.nextLink"];
		if (!next) break;
		page = (await graphGetNextLink(env, next)) as GraphPage;
		items.push(...(page.value ?? []));
	}

	return items;
}

export async function listTaskListsImpl(env: Env): Promise<unknown> {
	return sanitizeTaskLists(await collectPages(env, "/me/todo/lists"));
}

async function resolveTaskListId(env: Env, listId: string | undefined): Promise<string> {
	if (listId) return listId;
	const lists = (await collectPages(env, "/me/todo/lists")) as Array<{
		id: string;
		wellknownListName?: string;
	}>;
	const def = lists.find((l) => l.wellknownListName === "defaultList");
	const fallback = lists[0];
	if (def) return def.id;
	if (fallback) return fallback.id;
	throw new Error("No To Do lists found on this account.");
}

// Graph's taskStatus enum also carries `waitingOnOthers` and `deferred`, which
// the tool's own enum doesn't offer. Both are open work, so the default view
// keeps them by excluding only `completed` rather than listing what to include.
export function selectTasksByStatus(tasks: unknown[], status: string | undefined): unknown[] {
	return (tasks as Array<{ status?: string }>).filter((task) =>
		status ? task.status === status : task.status !== "completed",
	);
}

export async function listTasksImpl(
	env: Env,
	args: {
		list_id?: string;
		status?: "notStarted" | "inProgress" | "completed";
	},
): Promise<unknown> {
	const listId = await resolveTaskListId(env, args.list_id);
	const tasks = await collectPages(env, `/me/todo/lists/${listId}/tasks`);
	return sanitizeTaskList(selectTasksByStatus(tasks, args.status), listId);
}

export async function createTaskImpl(
	env: Env,
	args: {
		title: string;
		list_id?: string;
		due_date?: string;
		body?: string;
		importance?: "low" | "normal" | "high";
	},
): Promise<unknown> {
	const listId = await resolveTaskListId(env, args.list_id);

	const task: Record<string, unknown> = {
		title: args.title,
		importance: args.importance ?? "normal",
	};

	if (args.due_date) task.dueDateTime = { dateTime: args.due_date, timeZone: "UTC" };
	if (args.body) task.body = { content: args.body, contentType: "text" };

	const data = await graphPost(env, `/me/todo/lists/${listId}/tasks`, task);
	return { success: true, message: "Task created.", task: data };
}

// ── Completing a task ─────────────────────────────────────────────────────────
// `PATCH /me/todo/lists/{listId}/tasks/{taskId}` with `{ status: "completed" }`.
// Like everything else in this file the request carries no query string, which
// is what keeps To Do from answering `RequestBroker--ParseUri` (see the note at
// the top).
//
// Callers name a task one of two ways: `task_id` (exact, from list_tasks) or
// `title` (a case-insensitive substring). Title is the one people actually
// reach for — "mark off the invoice one" — so it has to be safe: a substring
// that hits two tasks completes neither and hands both back, because ticking
// off the wrong item is silent and annoying to undo.

interface RawTaskRow {
	id?: string;
	title?: string;
	status?: string;
	dueDateTime?: { dateTime?: string };
}

export interface TaskHit {
	id: string;
	title: string | undefined;
	status: string | undefined;
	due: string | null;
	list_id: string;
	list_name: string;
}

interface SearchList {
	id: string;
	name: string;
}

// Which lists to search. An explicit `list_id` narrows to that one list; with
// none, every list is searched rather than just the default.
//
// That is deliberately the opposite default to create_task. Creating needs one
// destination and the default list is the right guess; completing has to find
// something that already exists, and "mark off X" shouldn't miss because X was
// filed under Flagged Emails. Accounts here have a couple of lists, so the fan
// out is a handful of requests.
async function resolveSearchLists(
	env: Env,
	listId: string | undefined,
): Promise<SearchList[]> {
	const lists = (await collectPages(env, "/me/todo/lists")) as Array<{
		id?: string;
		displayName?: string;
	}>;

	if (listId) {
		const known = lists.find((l) => l.id === listId);
		// Unknown ids are still searched — the caller may hold a valid id for a
		// list this enumeration didn't return. Graph decides, not us.
		return [{ id: listId, name: known?.displayName ?? listId }];
	}

	return lists
		.filter((l): l is { id: string; displayName?: string } => Boolean(l.id))
		.map((l) => ({ id: l.id, name: l.displayName ?? l.id }));
}

function toHit(task: RawTaskRow, list: SearchList): TaskHit {
	return {
		id: task.id as string,
		title: task.title,
		status: task.status,
		due: task.dueDateTime?.dateTime ?? null,
		list_id: list.id,
		list_name: list.name,
	};
}

// Exported for tests: the matching rule on its own, no Graph involved.
//
// `task_id` matches any status — completing something already completed is
// worth reporting rather than hiding. `title` matches open tasks only, so a
// months-old finished task with a similar name can't make a live one ambiguous.
export function matchTasks(
	tasks: RawTaskRow[],
	list: SearchList,
	query: { task_id?: string; title?: string },
): TaskHit[] {
	if (query.task_id) {
		return tasks
			.filter((t) => t.id === query.task_id)
			.map((t) => toHit(t, list));
	}

	const needle = (query.title ?? "").trim().toLowerCase();
	if (!needle) return [];

	return tasks
		.filter(
			(t) =>
				t.status !== "completed" &&
				(t.title ?? "").toLowerCase().includes(needle),
		)
		.map((t) => toHit(t, list));
}

type TaskQuery = { task_id?: string; title?: string; list_id?: string };

// Narrow a query down to exactly one task, or explain why it couldn't.
// Shared by complete_task and delete_task so the two can't drift: whatever
// safety the matching gives one of them, it gives both.
//
// `verb` only shapes the message ("Nothing was deleted"), never the rule.
async function resolveOneTask(
	env: Env,
	args: TaskQuery,
	verb: string,
): Promise<{ hit: TaskHit } | { ambiguous: TaskHit[]; message: string }> {
	const lists = await resolveSearchLists(env, args.list_id);
	if (lists.length === 0) {
		throw new Error("No To Do lists found on this account.");
	}

	const perList = await Promise.all(
		lists.map(async (list) => {
			const tasks = (await collectPages(
				env,
				`/me/todo/lists/${list.id}/tasks`,
			)) as RawTaskRow[];
			return matchTasks(tasks, list, args);
		}),
	);
	const matches = perList.flat();

	const searched = args.list_id
		? `list ${lists[0]?.name}`
		: `${lists.length} list${lists.length === 1 ? "" : "s"}`;

	if (matches.length === 0) {
		throw ToolError.validation(
			args.task_id
				? `No task with id ${args.task_id} in ${searched}.`
				: `No open task matching "${args.title}" in ${searched}. Titles are matched as a case-insensitive substring; use list_tasks to see what's open.`,
		);
	}

	// Ambiguous: touch nothing, hand back the candidates so the next call can
	// pick one by id. Not an error — the caller asked a reasonable question and
	// this is the answer.
	if (matches.length > 1) {
		return {
			ambiguous: matches,
			message: `"${args.title}" matches ${matches.length} open tasks. Nothing was ${verb} — call again with task_id (and list_id) to choose one.`,
		};
	}

	return { hit: matches[0] as TaskHit };
}

export async function completeTaskImpl(env: Env, args: TaskQuery): Promise<unknown> {
	const resolved = await resolveOneTask(env, args, "completed");
	if ("ambiguous" in resolved) {
		return {
			success: false,
			reason: "ambiguous",
			message: resolved.message,
			matches: resolved.ambiguous,
		};
	}

	const hit = resolved.hit;

	if (hit.status === "completed") {
		return {
			success: true,
			already_completed: true,
			message: `"${hit.title}" was already completed.`,
			task: hit,
		};
	}

	const updated = await graphPatch(
		env,
		`/me/todo/lists/${hit.list_id}/tasks/${hit.id}`,
		{ status: "completed" },
	);

	return {
		success: true,
		message: `Completed "${hit.title}" in ${hit.list_name}.`,
		task: { ...hit, status: "completed" },
		raw: updated,
	};
}

// ── Creating a list ───────────────────────────────────────────────────────────
// `POST /me/todo/lists` with `{ displayName }`. No query string, same as
// everything else here.
//
// Graph will happily create a second list called "Work" alongside the first,
// which is almost never what someone means by "make me a Work list" — and a
// duplicate is worse than a no-op, because tasks then scatter across two lists
// that look identical in the UI. So an existing list with the same name
// (case-insensitive) is returned as-is instead, flagged with `created: false`.
export async function createTaskListImpl(
	env: Env,
	args: { name: string },
): Promise<unknown> {
	const name = args.name.trim();
	if (!name) throw ToolError.validation("List name cannot be empty.");

	const existing = (await collectPages(env, "/me/todo/lists")) as Array<{
		id?: string;
		displayName?: string;
	}>;
	const clash = existing.find(
		(l) => (l.displayName ?? "").trim().toLowerCase() === name.toLowerCase(),
	);
	if (clash) {
		return {
			success: true,
			created: false,
			message: `A list called "${clash.displayName}" already exists — returning it rather than creating a duplicate.`,
			list: sanitizeTaskLists([clash])[0],
		};
	}

	const data = await graphPost(env, "/me/todo/lists", { displayName: name });

	return {
		success: true,
		created: true,
		message: `Created task list "${name}".`,
		list: sanitizeTaskLists([data])[0],
	};
}

// ── Deleting ──────────────────────────────────────────────────────────────────
// The only irreversible operations here. Both reuse the matching rules above:
// an ambiguous title deletes nothing and returns the candidates, which matters
// far more for delete than it does for complete.

export async function deleteTaskImpl(env: Env, args: TaskQuery): Promise<unknown> {
	const resolved = await resolveOneTask(env, args, "deleted");
	if ("ambiguous" in resolved) {
		return {
			success: false,
			reason: "ambiguous",
			message: resolved.message,
			matches: resolved.ambiguous,
		};
	}

	const hit = resolved.hit;
	await graphDelete(env, `/me/todo/lists/${hit.list_id}/tasks/${hit.id}`);

	return {
		success: true,
		message: `Deleted "${hit.title}" from ${hit.list_name}.`,
		task: hit,
	};
}

interface RawListRow {
	id?: string;
	displayName?: string;
	wellknownListName?: string;
}

// Matching for lists mirrors matchTasks: exact on id, case-insensitive
// substring on name. Exported for tests.
export function matchLists(lists: RawListRow[], query: { list_id?: string; name?: string }) {
	if (query.list_id) return lists.filter((l) => l.id === query.list_id);

	const needle = (query.name ?? "").trim().toLowerCase();
	if (!needle) return [];
	return lists.filter((l) => (l.displayName ?? "").toLowerCase().includes(needle));
}

export async function deleteTaskListImpl(
	env: Env,
	args: { list_id?: string; name?: string; force?: boolean },
): Promise<unknown> {
	const lists = (await collectPages(env, "/me/todo/lists")) as RawListRow[];
	const matches = matchLists(lists, args);

	if (matches.length === 0) {
		throw ToolError.validation(
			args.list_id
				? `No task list with id ${args.list_id}.`
				: `No task list matching "${args.name}". Names are matched as a case-insensitive substring; use list_task_lists to see them.`,
		);
	}

	if (matches.length > 1) {
		return {
			success: false,
			reason: "ambiguous",
			message: `"${args.name}" matches ${matches.length} lists. Nothing was deleted — call again with list_id to choose one.`,
			matches: sanitizeTaskLists(matches),
		};
	}

	const list = matches[0] as RawListRow;

	// Graph refuses this anyway; saying so plainly beats surfacing its error.
	if (list.wellknownListName === "defaultList") {
		throw ToolError.validation(
			`"${list.displayName}" is the default To Do list and cannot be deleted.`,
		);
	}

	// Deleting a list takes every task in it, and Graph gives no warning and no
	// undo. Counting first turns a silent bulk delete into one the caller had to
	// mean — `force` is the "yes, and its contents too".
	const tasks = (await collectPages(
		env,
		`/me/todo/lists/${list.id}/tasks`,
	)) as RawTaskRow[];
	const open = tasks.filter((t) => t.status !== "completed").length;

	if (tasks.length > 0 && !args.force) {
		return {
			success: false,
			reason: "not_empty",
			message: `"${list.displayName}" holds ${tasks.length} task${tasks.length === 1 ? "" : "s"} (${open} still open). Deleting the list deletes them too, permanently. Call again with force: true to go ahead.`,
			task_count: tasks.length,
			open_task_count: open,
			list: sanitizeTaskLists([list])[0],
		};
	}

	await graphDelete(env, `/me/todo/lists/${list.id}`);

	return {
		success: true,
		message: `Deleted list "${list.displayName}"${tasks.length > 0 ? ` and its ${tasks.length} task${tasks.length === 1 ? "" : "s"}` : ""}.`,
		deleted_task_count: tasks.length,
		list: sanitizeTaskLists([list])[0],
	};
}

export const tasksTools = defineTools<Env>({
	list_task_lists: {
		description:
			"List all Microsoft To Do task lists. Use this to find list IDs before listing or creating tasks.",
		schema: z.object({}),
		handler: (env) => listTaskListsImpl(env),
	},

	list_tasks: {
		description:
			"List tasks in a To Do task list. Defaults to the default task list and excludes completed tasks.",
		schema: z.object({
			list_id: taskListIdSchema.optional(),
			status: z.enum(["notStarted", "inProgress", "completed"]).optional(),
		}),
		handler: (env, args) => listTasksImpl(env, args),
	},

	create_task_list: {
		description:
			"Create a new Microsoft To Do list. If a list with that name already exists it is returned untouched rather than duplicated — check the `created` field to tell the two apart.",
		schema: z.object({
			name: z.string().min(1).max(256),
		}),
		handler: (env, args) => createTaskListImpl(env, args),
	},

	create_task: {
		description: "Create a new task in a Microsoft To Do list.",
		schema: z.object({
			title: z.string().min(1).max(256),
			list_id: taskListIdSchema.optional(),
			due_date: z.string().min(1).max(64).optional(),
			body: z.string().max(64 * 1024).optional(),
			importance: z.enum(["low", "normal", "high"]).optional(),
		}),
		handler: (env, args) => createTaskImpl(env, args),
	},

	delete_task: {
		description:
			"Permanently delete a Microsoft To Do task. Identify it with `title` (case-insensitive substring of an open task) or `task_id`. `list_id` is optional: without it every list is searched. An ambiguous title deletes nothing and returns the candidates. This cannot be undone — prefer complete_task for work that is finished rather than mistaken.",
		schema: z
			.object({
				task_id: taskListIdSchema.optional(),
				title: z.string().min(1).max(256).optional(),
				list_id: taskListIdSchema.optional(),
			})
			.refine(
				(v) => [v.task_id, v.title].filter(Boolean).length === 1,
				"Provide exactly one of: task_id, title.",
			),
		handler: (env, args) => deleteTaskImpl(env, args),
	},

	delete_task_list: {
		description:
			"Permanently delete a Microsoft To Do list. Identify it with `name` (case-insensitive substring) or `list_id`. Deleting a list also deletes every task in it, so a non-empty list is refused unless you pass force: true — the refusal reports how many tasks would go. The default list cannot be deleted. An ambiguous name deletes nothing and returns the candidates.",
		schema: z
			.object({
				list_id: taskListIdSchema.optional(),
				name: z.string().min(1).max(256).optional(),
				force: z.boolean().optional(),
			})
			.refine(
				(v) => [v.list_id, v.name].filter(Boolean).length === 1,
				"Provide exactly one of: list_id, name.",
			),
		handler: (env, args) => deleteTaskListImpl(env, args),
	},

	complete_task: {
		description:
			"Mark a Microsoft To Do task as done. Identify it with either `title` (case-insensitive substring of an open task's title — what to use for 'tick off the invoice one') or `task_id` from list_tasks. `list_id` is optional: without it every To Do list is searched, not just the default one. If a title matches more than one open task nothing is completed and the candidates are returned, so call again with the task_id you want.",
		schema: z
			.object({
				task_id: taskListIdSchema.optional(),
				title: z.string().min(1).max(256).optional(),
				list_id: taskListIdSchema.optional(),
			})
			.refine(
				(v) => [v.task_id, v.title].filter(Boolean).length === 1,
				"Provide exactly one of: task_id, title.",
			),
		handler: (env, args) => completeTaskImpl(env, args),
	},
});
