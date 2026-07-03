# Problem 11: Persist user settings across sessions

**Tier:** 4
**Concepts:** `Office.context.document.settings.set/get/saveAsync`
**Prerequisites:** Problems 1–10
**Time budget:** ~40 minutes

## Statement

Add a text input labelled `User's preferred date format` and a button labelled `Save`. When the user types a format string (e.g., `YYYY-MM-DD`) and clicks Save, the value is persisted via `Office.context.document.settings`.

On task pane load (`Office.onReady`), read the saved setting and prepopulate the input with it. If no setting exists yet, display an empty input.

Additionally: display the current saved value in a `<span id="saved-format-display">` that refreshes after each Save.

## Constraints

- Use `Office.context.document.settings.set(key, value)` and `.get(key)`.
- Must call `.saveAsync()` after `.set()` — the set alone doesn't persist.
- The key is exactly `preferredDateFormat`.
- Reads happen in `Office.onReady`; writes on Save button click.

## Automated verification

The test file at `tests/11-settings-storage.test.ts` asserts:

- On load, `settings.get('preferredDateFormat')` is called.
- On save, `settings.set('preferredDateFormat', <value>)` is followed by `settings.saveAsync`.
- The saved value is written to `saved-format-display` after `.saveAsync`'s callback resolves.
- If `.get` returns undefined on load, the input is left empty (no `undefined` shown as literal text).

## Manual verification

- [ ] Type `YYYY-MM-DD` in the input, click `Save`.
- [ ] `saved-format-display` updates to `YYYY-MM-DD`.
- [ ] Close and reopen the task pane (or the whole document). Confirm the input is prepopulated with `YYYY-MM-DD`.
- [ ] Change to `DD/MM/YYYY`, Save. Reopen. Input shows `DD/MM/YYYY`.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Office.Settings.saveAsync method". Note the async signature — it's not the same as the newer `Word.run`-style promise-based API.
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Forgetting `.saveAsync()`. Calling `.set()` alone modifies the in-memory settings object but doesn't persist to the document. Reopening loses your changes.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
On save: `Office.context.document.settings.set('preferredDateFormat', value); Office.context.document.settings.saveAsync((result) => { /* update DOM */ });`. On load: `const saved = Office.context.document.settings.get('preferredDateFormat'); inputEl.value = saved || '';`.
</details>

## Reference solution

`git diff solutions/11 -- app/`
