# Problem 9: Track changes — toggle and accept

**Tier:** 3
**Concepts:** `Document.changeTrackingMode`, `Body.getRevisions()`, revision accept/reject
**Prerequisites:** Problems 1–8
**Time budget:** ~50 minutes

## Statement

Add three buttons:

1. `Enable tracking` — sets `context.document.changeTrackingMode` to `'TrackAll'`.
2. `Disable tracking` — sets it to `'Off'`.
3. `Accept all revisions` — accepts every pending revision in the document body.

The task pane must also display the current tracking mode and the count of pending revisions in a `<span id="tracking-status">` — refreshed after every button click.

## Constraints

- Use the string values `'TrackAll'` and `'Off'` (or the corresponding enum values if you prefer).
- Revision accept uses `body.getRevisions()` returning a `RevisionCollection`; accept each via `revision.accept()`.
- The status refresh happens in the same `Word.run` as the button action (one sync round-trip per click, ideally).

## Automated verification

The test file at `tests/09-track-changes.test.ts` asserts:

- `Enable tracking` sets `changeTrackingMode` to `'TrackAll'`.
- `Disable tracking` sets it to `'Off'`.
- `Accept all revisions` iterates the revision collection and calls `.accept()` on each.
- After each handler, the task pane's `tracking-status` reflects the new mode and revision count.

## Manual verification

- [ ] Click `Enable tracking`. Task pane shows `Mode: TrackAll, 0 revisions`.
- [ ] In the document, type a new sentence and delete an existing word. Word marks these as tracked changes.
- [ ] Click `Accept all revisions`. All revision marks disappear; the document contains the accepted edits.
- [ ] Task pane updates: `Mode: TrackAll, 0 revisions`.
- [ ] Click `Disable tracking`. Task pane shows `Mode: Off, 0 revisions`. Further edits are no longer marked.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Document.changeTrackingMode property" and "Word.Revision object".
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Trying to accept revisions inside a track-changes cycle (revisions on revisions on revisions). Keep it simple: enable, edit, accept, disable.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
Toggle: `context.document.changeTrackingMode = 'TrackAll'; await context.sync();`. Accept: `const revs = context.document.body.getRevisions(); revs.load('items'); await context.sync(); revs.items.forEach(r => r.accept()); await context.sync();`.
</details>

## Reference solution

`git diff solutions/9 -- app/`
