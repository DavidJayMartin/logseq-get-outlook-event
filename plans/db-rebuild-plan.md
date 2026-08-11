# Plan: Rebuild "Get Outlook Events" Logseq Plugin for the Database Version

## Summary

This repo contains a Logseq plugin (`main.js`) plus a local Python Flask API
(`outlook_events_api.py`) that reads events from a local Outlook calendar and
inserts them into Logseq journal pages. It was written for the **file/markdown**
version of Logseq and no longer works on the **database (DB)** version.

- The **Python API is independent of Logseq and requires no changes.**
- All breakage is in the **Logseq plugin (`main.js`)**.

**Root cause:** In DB graphs, the `key:: value` markdown property syntax that the
old code embedded in block text (e.g. `event-time:: 9:00 AM`) no longer creates
properties. DB properties are now **typed, first-class database entities** that
must be defined and set through new plugin APIs. The old page/journal detection
(`page.journalDay`) and the text-template output model also need to change.

## Goals / Decisions (confirmed with repo owner)

1. **Use native typed DB properties** via `upsertProperty` / `upsertBlockProperty`.
2. **Attendees are node/page references** (property type `node`, cardinality `many`).
3. **Simplify to a fixed layout** — remove the custom output template and
   `---CHILD---` block engine.
4. **Runtime feature-detection** — do not hard-require a version; detect whether the
   new APIs exist and show a friendly "please update Logseq" message if not.

### Open sub-decisions (recommended defaults)
- `event-time`: keep as a **text** property (DB's `date` type stores a whole day and
  drops clock time). Optionally add `event-start`/`event-end` datetime properties if
  the target build supports a datetime type.
- Event **title as a node**: keep the old `bracketEvents` setting as an option,
  default off.

## Key API Facts (DB version `@logseq/libs`)

Confirmed from the official plugin API docs. Relevant `logseq.Editor` methods:

- `upsertProperty(key, schema?, opts?)` — define a typed property.
  `schema.type` ∈ `default | number | node | date | checkbox | url`;
  `schema.cardinality` ∈ `one | many`.
- `upsertBlockProperty(block, key, value)` — set a property value on a block.
- `getBlockProperties(block)` / `getBlockProperty(block, key)` — read properties.
- `addBlockTag(blockId, tagId)`, `createTag(name)`, `getTag(nameOrIdent)`,
  `getTagsByName(name)`, `getTagObjects(name)` — tag/node helpers.
- `createJournalPage(date)`, `createPage(name, properties?, opts?)`.
- `insertBlock(srcBlock, content, { sibling, before, properties, customUUID })`.
- `getCurrentBlock()`, `getCurrentPage()`, `getPage(srcPage)`.

## Reference Links

- Plugin API reference (all interfaces): https://logseq.github.io/plugins/
- `IEditorProxy` (all Editor methods above): https://logseq.github.io/plugins/interfaces/IEditorProxy.html
- `PageEntity` (journal/page fields): https://logseq.github.io/plugins/interfaces/PageEntity.html
- `BlockEntity`: https://logseq.github.io/plugins/interfaces/BlockEntity.html
- Community plugin docs (guides/best practices): https://plugins-doc.logseq.com/
- Official plugin samples (patterns for insertBlock, slash commands, settings): https://github.com/logseq/logseq-plugin-samples
- Libs source (`LSPlugin.ts`, exact method signatures): https://github.com/logseq/logseq/blob/master/libs/src/LSPlugin.ts
- DB-version discussions (property model, migration context): https://github.com/logseq/logseq/discussions/categories/db-version

> Implementer note: verify each method exists on the owner's installed build via
> `typeof logseq.Editor.<method> === 'function'` before relying on it. Method
> availability tracks the libs version bundled with the Logseq build.

## Implementation Steps

### 1. Startup: feature-detection + property registration (`main()`)
- On load, check `typeof logseq.Editor.upsertBlockProperty === 'function'` and
  `typeof logseq.Editor.upsertProperty === 'function'`. If missing, call
  `logseq.UI.showMsg('Update Logseq to a DB build to use Get Outlook Events', 'error')`
  and abort registration.
- Register typed properties once (idempotent) via `upsertProperty`:
  - `event-time` → `{ type: 'default' }` (formatted start time text)
  - `event-duration` → `{ type: 'default' }` (HH:MM:SS)
  - `event-location` → `{ type: 'default' }`
  - `attendees` → `{ type: 'node', cardinality: 'many' }`
  - `recurring` → `{ type: 'checkbox' }` (optional)

### 2. Journal-date detection (replace `page.journalDay` reliance in `getEvents`)
- Read `const page = await logseq.Editor.getCurrentPage()`.
- Feature-detect the journal date: prefer `page.journalDay` if present, else check
  `page['journal?']` / `page.journal` and derive `YYYY-MM-DD`.
- Keep helpers `journalDayToDate()` and `formatDateForAPI()`.
- If not a journal page, `insertBlock` a friendly message and return.

### 3. Rewrite event insertion (`getEvents`) — depends on 1–2
- Fetch via existing `fetchEventsFromAPI(dateString)` (unchanged; hits
  `http://localhost:5000/events/<date>`).
- Sort events by start time (existing logic).
- For each event:
  - Insert a block whose **content is only** the subject + recurring emoji +
    meeting links (reuse `formatEventSubject` + `createMeetingLinks`).
  - Then set properties with `upsertBlockProperty(block.uuid, 'event-time', formatTime(event.start))`,
    `'event-duration'`, `'event-location'`, and `'recurring'` (boolean).
  - Respect `includeEmptyFields` (skip empty values when false).

### 4. Attendees as nodes — depends on 3
- In `formatAttendees`, keep exclusion (`excludeUserName`) filtering, but return an
  array of names (not a bracketed string).
- For each attendee, ensure a page exists (`getTag`/`getPage`, else `createPage`),
  then set the `attendees` node property via `upsertBlockProperty(block.uuid,
  'attendees', <names-or-ids>)`. Verify at implementation whether the node property
  accepts names (auto-create) or requires entity IDs.

### 5. Trim settings & remove template engine — parallel with 4
- Delete `formatEventContent` and the `outputFormat` / `---CHILD---` logic.
- Keep settings schema entries: `excludeUserName`, `apiUrl`, `meetingBaseUrls`,
  `timeFormat`, `descriptionMaxLength`, `includeEmptyFields`, and `bracketEvents`
  (repurposed to "make event title a node").
- Keep and reuse: `formatTime`, `calculateDuration`, `formatEventSubject`,
  `createMeetingLinks`, `fetchEventsFromAPI`.

### 6. Verify plugin manifest (`package.json`)
- Confirm whether the DB build requires a compatibility flag in the `logseq` manifest
  block. Add it **only** if verification shows it is required. (Check current samples
  and the installed build's plugin loader behavior.)

### 7. Update `readme.md` — parallel with 3–5
- Document DB-property behavior, the simplified settings, and that fields now appear
  as real typed properties (not `key:: value` text).

## Files

| File | Change |
|------|--------|
| `main.js` | Rewrite `getEvents` and `formatAttendees`; remove `formatEventContent` + template settings; add property registration + feature-detection in `main()`. Keep `formatTime`, `calculateDuration`, `formatEventSubject`, `createMeetingLinks`, `fetchEventsFromAPI`, `journalDayToDate`, `formatDateForAPI`. |
| `package.json` | Possibly add DB-compat manifest flag (pending verification). |
| `readme.md` | Update settings/behavior sections. |
| `outlook_events_api.py` | **No change** (Outlook COM API, independent of Logseq). |
| `index.html` | **No change** (loads `@logseq/libs` + `main.js`). |

## Verification

1. Load the unpacked plugin in a DB graph; confirm no startup error and that the new
   properties appear in the graph's property list.
2. On a journal page with known Outlook events, run `/Get Events`; confirm blocks
   show `event-time`, `event-duration`, `event-location` as **typed properties**, not
   literal text.
3. Confirm each attendee renders as a clickable page link and the excluded name is
   omitted.
4. Confirm recurring events show the emoji and meeting links are clickable.
5. Confirm a non-journal page shows the friendly "journal pages only" message.
6. Sanity-check the API is reachable: `http://localhost:5000/health` returns
   `{"status":"healthy",...}` and `http://localhost:5000/events/<known-date>` returns events.

## Out of Scope

- Changes to the Python API service, the auto-startup scripts
  (`start_api_service.bat`, `launch_api.vbs`, `stop_api_service.bat`), and `index.html`.