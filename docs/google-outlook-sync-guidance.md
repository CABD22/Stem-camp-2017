# Google ↔ Outlook Calendar Sync Flow Design

This document captures a maintainable Power Automate approach for Preston’s scenario:

- Events are created on Google Calendar (from an iPhone) and need a mirrored copy on the DBA Outlook calendar that includes Zoom details.
- Zero duplicates are acceptable; a single mapping between the Google and Outlook event IDs must be maintained.
- The flow should be clear, predictable, and easy to extend toward future bi-directional sync.

## Core pattern

1. **Single source of truth for mappings**
   - Keep an Excel table with columns: `GoogleId`, `OutlookId`, `LastSyncedUtc`, and optional `Etag`/`Sequence` for change detection.
   - Always read the mapping once per trigger run and pass it forward instead of re-querying in each branch.

2. **Deterministic branching by action type**
   - Switch on the Google trigger’s `actionType` (`added`, `updated`, `deleted`).
   - Each branch reads from the pre-fetched mapping and decides whether to create, update, or delete.

3. **Idempotent operations**
   - Use the stored `OutlookId` for updates/deletes; if absent, fall back to create + add mapping.
   - Set a `ClientState`/`Custom` field to a stable value (e.g., `"synced-from-google"`) to avoid treating your own writes as new inbound changes when you later add Outlook→Google sync.

## Recommended flow outline (Google → Outlook)

1. **Trigger**: Google Calendar “When an event is added, updated or deleted” on the target calendar. Configure short recurrence (e.g., 1 minute) and enable split-on.
2. **Compose: GoogleId** — `@triggerOutputs()?['body/id']`
3. **Compose: ChangeType** — `@triggerOutputs()?['body/actionType']`
4. **List rows in Excel table** — pull the mapping table.
5. **Filter array: MappingByGoogleId** — `item()?['GoogleId']` equals `outputs('Compose_GoogleId')`.
6. **Compose: HasMapping** — `@greater(length(body('MappingByGoogleId')), 0)`.
7. **Compose: OutlookId** — `@first(body('MappingByGoogleId'))?['OutlookId']`.
8. **Switch on ChangeType**:
   - **added**
     - If **HasMapping** is false → `Create event (V4)` in Outlook using Google payload (subject, body, location, start/end with timezone, attendees, reminders, online meeting data if present).
     - Store new `OutlookId` to Excel mapping with `GoogleId` and timestamps.
   - **updated**
     - If **HasMapping** is true → `Update event (V4)` using `OutlookId`; write back any Zoom/conference URL from Google into the Outlook body/location.
     - Else → treat as **added** (create then add mapping) to cover late-arriving updates.
   - **deleted**
     - If **HasMapping** is true → `Delete event (V4)` using `OutlookId`; remove the Excel row to prevent reprocessing and keep the table clean.
     - Else → no-op (nothing to delete).

## Duplicate prevention

- **Mapping gate**: All operations route through the mapping lookup. Creation only occurs when the mapping is missing, preventing duplicate Outlook events.
- **Stable identifiers**: Never regenerate `GoogleId`/`OutlookId`; read from triggers and stored mapping only.
- **Etag/Sequence check (optional)**: Store the Google `etag` or `sequence` in Excel and compare on update to avoid overwriting newer data.

## Prepare for future Outlook → Google sync

- Mirror the same mapping table but filter by `OutlookId` when processing Outlook triggers.
- Tag outbound writes (e.g., custom text in the body or extended properties) so the opposite flow can ignore its own updates.
- Consider a “debounce” Compose that skips processing when the triggering platform matches the last writer tag.

## Operational tips

- **Time zones**: Always pass the Google event’s time zone to Outlook to avoid offsets.
- **Attendees and reminders**: Map attendee emails and reminders consistently; include Zoom join URL in the description or location.
- **Error handling**: Add a scope with retry + failure notification so missed syncs can be replayed manually.
- **Maintenance**: Periodically purge stale mappings where either ID no longer exists to keep the table small and fast.

## Expression fixes for the latest outline you shared

Your newest step list is close. Use these small changes to keep the logic deterministic and avoid duplicate writes:

- **Mapping gate**
  - `Filter_MappingByGoogleId` should use `from: @outputs('List_rows_present_in_a_table')?['body/value']` and `where: @equals(item()?['GoogleId'], outputs('Compose_GoogleId'))`. This keeps the lookup consistent with the Compose step.
  - `Compose_HasMapping` can simply be `@greater(length(body('Filter_MappingByGoogleId')), 0)`.
  - Add a **Compose: OutlookId** right after HasMapping: `@first(body('Filter_MappingByGoogleId'))?['OutlookId']`. Reuse this in all branches so you don’t recompute it.

- **Switch: added**
  - Condition should check “no mapping yet”: `@not(outputs('Compose_HasMapping'))`.
  - After **Create event (V4)**, store `GoogleId`, `OutlookId`, and optionally timestamps/etag in Excel. This prevents the same Google event from creating multiple Outlook events.

- **Switch: updated**
  - Condition should be `@equals(outputs('Compose_HasMapping'), true)`.
  - If true → `Update event (V4)` using `OutlookId`. If false → reuse the **added** path (create + add mapping) so late-arriving updates don’t get skipped.

- **Switch: deleted**
  - Condition should also be `@equals(outputs('Compose_HasMapping'), true)`.
  - If true → `Delete event (V4)` using `OutlookId` and then delete the Excel row to keep the mapping table clean. If false → no-op (nothing to delete).

Following these expressions keeps the flow idempotent: create happens only when the mapping is absent, updates and deletes use the stored OutlookId, and the Excel table remains the single source of truth.

This structure keeps the logic predictable today and sets you up for bi-directional sync without duplicates later.
