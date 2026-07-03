# Problem 4: Operate on the user's current selection

**Tier:** 2
**Concepts:** `context.document.getSelection()`, `Range.font`, `Range.insertText`
**Prerequisites:** Problems 1–3
**Time budget:** ~45 minutes

## Statement

Add a button labelled `Bold and shout`. When clicked, it reads the user's current selection, replaces it with the uppercased version of the same text, and applies bold formatting to the replacement. If the user has nothing selected (empty selection), the task pane displays `Nothing selected` and no change is made to the document.

## Constraints

- Must use `context.document.getSelection()`.
- Must handle the empty-selection case gracefully (display message, no error).
- The replacement text preserves the original selection's position (before + after ranges are unchanged).

## Automated verification

The test file at `tests/04-selection.test.ts` asserts:

- With a non-empty selection mock, `Range.insertText` is called with the uppercased original and location `'Replace'`.
- With a non-empty selection mock, `range.font.bold = true` is set before sync.
- With an empty selection mock, `Range.insertText` is NOT called and the task pane's DOM includes `Nothing selected`.

## Manual verification

- [ ] Type `hello there` in the document, select `hello`, click `Bold and shout`.
- [ ] Selection is replaced with `HELLO` in bold; `there` is unchanged.
- [ ] Select nothing (empty caret), click button — task pane shows `Nothing selected`, document unchanged.
- [ ] Ctrl+Z returns to original state.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Range object" and "Word.Range font property". Also the "isEmpty" property on selection.
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Applying `range.font.bold = true` before the replacement. In some Word versions the formatting doesn't survive the replacement text. Order: replace first, then style.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
`const range = context.document.getSelection(); range.load('text'); await context.sync(); if (!range.text) { /* empty */ return; } range.insertText(range.text.toUpperCase(), 'Replace'); range.font.bold = true; await context.sync();`.
</details>

## Reference solution

`git diff solutions/4 -- app/`
