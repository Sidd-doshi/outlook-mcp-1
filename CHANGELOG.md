# Changelog

Notable changes to the Outlook MCP server. Versions match `package.json` and the
`SERVER_VERSION` reported in the MCP handshake — `test/version.spec.ts` keeps the
two from drifting.

Entries before 2.5.0 were reconstructed from commit history after the fact, so
they summarise what shipped rather than what was written down at the time.

Releases are cut with `npm run release <version>`, which turns the **Unreleased**
section below into a dated heading, bumps `package.json` and `SERVER_VERSION`
together, commits, and tags. Add to Unreleased as you go.

## 2.8.0 — 2026-09-28

### Added

Reading attachments on received email, so an agent can pick up files that
arrive by email (for example a form's PDF export) without them passing
through the conversation.

- `list_email_attachments` — each attachment's id, name, content type, size,
  inline flag (signature logos and the like), and kind: file, item (an
  attached email or event) or reference (a cloud link). No contents.
- `save_email_attachment` — copy one file attachment straight into a OneDrive
  folder. Chooses by `attachment_name` (exact, case-insensitive) or
  `attachment_id`, or automatically when the email has exactly one
  non-inline file; with several candidates it lists them and saves nothing.
  File names from the sender are cleaned of characters OneDrive or the path
  rules reject, executables are refused, files up to 25 MB are accepted
  (upload session above 4 MB), and `conflict_behavior` defaults to `rename`.

### Changed

- `files.ts` exports its path schema, blocked-extension rule and upload
  helper so other tool modules share them.

## 2.7.0 — 2026-09-27

### Added

Five OneDrive tools that together cover a document workflow with no desktop
Office install: produce a .docx anywhere, upload it, convert it to PDF with
Microsoft's renderer, file it, and attach it to a draft.

- `convert_to_pdf` — render a docx/pptx/xlsx (and rtf, odt, html, md and more)
  to PDF via Graph `…/content?format=pdf` and save it to OneDrive, beside the
  source by default. Fonts and layout match Word, which a PDF produced on a
  Linux machine without the document's fonts would not.
- `download_onedrive_file` — read a file's contents as base64, with a decoded
  `text` field for text formats. 5 MB default cap, 10 MB maximum.
- `create_folder` — create a folder path, including missing parents. Safe to
  repeat; a file in the way is reported rather than renamed around.
- `copy_item` — copy a file or folder into another folder, optionally renamed.
  Polls Graph's async monitor so the common case returns the finished item.
  Defaults to `conflict_behavior: "rename"` so nothing is overwritten by
  accident.
- `upload_large_file` — chunked upload through a Graph upload session, up to
  25 MB. Chunks are 320 KiB-aligned, and a failed chunk cancels the session.

### Changed

- `upload_onedrive_file`'s over-4 MB error now points at `upload_large_file`.
- The raw Graph helpers share one error builder, and a new `graphPostAccepted`
  handles `202 Accepted` operations that answer with a monitor URL.

### Fixed

- `download_onedrive_file` size errors read in KB below 1 MB. A 50 KB
  `max_bytes` used to report "over the 0.0 MB limit" (found in the live test).

## 2.6.0 — 2026-09-08

### Added

- `delete_task` — permanently delete a To Do task by `title` or `task_id`.
- `delete_task_list` — permanently delete a list by `name` or `list_id`.

Both share the matching rules used by `complete_task`: a title is a
case-insensitive substring of an open task, and one that matches more than one
task acts on **none** of them and returns the candidates instead. Deleting the
wrong item cannot be undone, so an ambiguous phrase never guesses.

`delete_task_list` carries two guards of its own, because deleting a list takes
every task inside it and Graph gives no warning:

- a non-empty list is refused unless `force: true` is passed, and the refusal
  reports how many tasks would go and how many are still open;
- the default list is refused outright.

## 2.5.0 — 2026-09-08

### Fixed

- **Teams recordings and transcripts were invisible for any meeting with an
  agenda written on it.** `list_recent_meeting_recordings` would report an empty
  calendar while recordings sat in OneDrive.

  When Graph leaves an event's `onlineMeetingUrl` and `onlineMeeting.joinUrl`
  empty — common for invites created through the Outlook Teams add-in — the join
  URL is recovered from the message body. That fallback read `bodyPreview`, which
  Graph truncates at 255 characters. Any invite whose body opened with the
  organiser's own note pushed the Teams block past the ceiling and the preview
  ended mid-URL, at `https://teams.m`. No match, and the event was dropped as if
  it had never been a Teams meeting.

  Single-event lookups now select the full `body`. Discovery keeps the cheap
  `bodyPreview` in its listing and refetches full bodies only for the candidates
  preview failed on, capped at 25 per call. Extraction decodes HTML entities
  first so long-form `?context=…&…` links survive intact.

### Added

- `complete_task` — mark a To Do task done by `title` or `task_id`. Without a
  `list_id` it searches every list, not just the default: completing has to find
  something that already exists, and shouldn't miss because the task was filed
  elsewhere. An ambiguous title completes nothing and returns the candidates.
- `create_task_list` — create a To Do list. An existing list of the same name is
  returned rather than duplicated; check the `created` field to tell them apart.
  Two lists sharing a name scatter tasks across entries that look identical in
  the To Do UI.
- `list_recent_meeting_recordings` now returns a `skipped` breakdown of the
  meetings it examined but did not return, so an empty result explains itself:
  `no_content` and `forbidden` are ordinary and permanent, while `no_join_url`,
  `unresolved_meeting` and `error` mean something is wrong. Previously every
  per-event failure was swallowed silently.

### Changed

- `list_tasks` now stamps each task with its `list_id`. A To Do task id is not
  addressable on its own — every write is scoped to
  `/me/todo/lists/{listId}/tasks/{id}` — so a caller reading a task previously
  had no way to act on it.
- The package description claimed 37 tools against an actual 41.

## 2.4.0 — 2026-06-15

### Added

- Threaded draft `reply` / `reply-all` / `forward` tools that preserve the
  conversation for review before sending.
- `get_onedrive_file_info` and a OneDrive streaming download route.
- `update_contact`.
- Server-side signature injection, with `include_signature` on draft and
  calendar event tools, preserving quoted originals below the reply.

### Fixed

- **To Do collections rejected every OData query option.** `/me/todo/lists?$select=…`
  answers `400 invalidRequest` with an inner code of `RequestBroker--ParseUri`,
  naming neither the option nor a property. The properties were real and the
  spelling correct; the backend simply doesn't implement the option. Both To Do
  read paths had been broken since they were written. The query options are gone
  and status filtering moved into the Worker.
- `list_recent_meeting_recordings` no longer filters on `isOnlineMeeting`, which
  Graph rejects as unfilterable, failing the whole request.
- Meeting discovery follows Graph's `@odata.nextLink` verbatim instead of
  rebuilding the query, which silently truncated results after page one.
- Conversation lookups are paged and sorted in the Worker rather than in Graph.
- Contacts tools no longer send an invented `phones` property.
- The MCP handshake reported 2.0.0 while `package.json` had reached 2.4.0. It had
  drifted for four releases; nothing branched on the value, so nothing broke and
  nothing complained.
- Loopback OAuth callbacks now accept `localhost` as well as `127.0.0.1`.

### Changed

- Moved to `@bashco/mcp-toolkit` ^1.1.0.

## 2.3.0 — 2026-06-09

### Added

- `upload_onedrive_file` for OneDrive simple upload, enabling host-then-share
  workflows such as `.vcf` contact cards.
- `127.0.0.1` loopback redirects, for native MCP clients.

### Fixed

- Self-sent replies are redirected to the original recipients.

## 2.2.3 — 2026-05-19

### Added

- Teams meeting recordings and transcripts: `find_online_meeting`,
  `list_meeting_recordings`, `list_meeting_transcripts`,
  `get_transcript_content`, and `list_recent_meeting_recordings` for discovery.
  Each per-meeting tool accepts `meeting_id`, `calendar_event_id`, or `join_url`.

## 2.0.0 — 2026-05-18

First public release: email, calendar, contacts, tasks and files over proxied
Microsoft Graph OAuth, deployed on Cloudflare Workers.
