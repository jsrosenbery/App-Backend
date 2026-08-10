# Low Enrollment Tracking API

Temporary review branch for the Phase 1 Low Enrollment Tracking backend.

## Storage

Low Enrollment Tracking workspaces are stored as JSON files under:

```text
DATA_DIR/low-enrollment-tracking/<termCode>.json
```

In Render, `DATA_DIR` should resolve to the mounted persistent disk, normally `/var/data/cos-app`. The per-term JSON file is authoritative.

Writes are atomic: the backend writes a temporary file in the same directory, flushes it, then renames it over the target file. Writes are serialized per term so snapshot saves and row comment saves do not overwrite each other.

## Authorization

Write endpoints require a valid bearer enrollment session with `development` or `admin` role level. Passwords are exchanged for bearer tokens through the existing role authentication endpoint.

Read endpoints follow the existing backend pattern and are not more restrictive than comparable report reads.

## Endpoints

### `GET /api/low-enrollment-tracking`

Lists saved term summaries sorted by `updatedAt` descending.

### `GET /api/low-enrollment-tracking/:termCode`

Returns the complete workspace for a six-digit Banner term code.

### `POST /api/low-enrollment-tracking/:termCode`

Creates or replaces an initial workspace.

```json
{
  "workspace": {},
  "replaceExisting": false
}
```

Existing workspaces are not overwritten unless `replaceExisting` is `true`. Replacement preserves the original `createdAt` and updates `updatedAt`.

### `POST /api/low-enrollment-tracking/:termCode/snapshots`

Adds or replaces a dated enrollment snapshot.

```json
{
  "snapshot": {},
  "uploadHistory": {},
  "rows": [],
  "replaceExisting": false
}
```

Snapshot row IDs must exactly match the saved workspace row IDs. Justification and VP comments from the saved workspace are preserved; only the row-specific PATCH endpoint may edit those fields.

### `PATCH /api/low-enrollment-tracking/:termCode/rows/:rowId`

Updates one or both editable fields:

```json
{
  "justification": "Dual Enrollment",
  "vpComments": "Reviewed with dean."
}
```

Unsupported fields are rejected. Justification must be blank or exactly match one of the saved workspace reasons.

### `POST /api/low-enrollment-tracking/:termCode/manual-import`

Atomically applies Justification and VP comment changes parsed from a Timber tracker export. Every update must contain a known stable row ID and both manual fields. Blank values intentionally clear saved values. Any unknown row or non-dropdown justification rejects the entire request; enrollment fields and snapshots are never accepted by this endpoint.

### `POST /api/low-enrollment-tracking/:termCode/rows/:rowId/exclusion`

Excludes or restores one row for the selected term. Exclusion requires a reason, retains all enrollment history, and is recorded in `exclusionHistory`. The row-level exclusion metadata is preserved by enrollment snapshot saves. A restored row becomes active again without recreating it.

## Error Shape

Errors use:

```json
{
  "error": "Human-readable message",
  "code": "MACHINE_READABLE_CODE"
}
```

Common codes include `INVALID_TERM`, `UNAUTHORIZED`, `FORBIDDEN`, `WORKSPACE_NOT_FOUND`, `WORKSPACE_EXISTS`, `INVALID_WORKSPACE`, `SNAPSHOT_EXISTS`, `INVALID_SNAPSHOT`, `ROW_NOT_FOUND`, `INVALID_JUSTIFICATION`, `INVALID_ROW_UPDATE`, and `STORAGE_FAILURE`.

## Test Command

```bash
npm test
```

The tests use temporary `DATA_DIR` paths and do not write to production persistent storage.
