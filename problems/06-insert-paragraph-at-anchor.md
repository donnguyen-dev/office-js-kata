# Problem 6: Insert a paragraph at a specific anchor

**Tier:** 2
**Concepts:** `Paragraph.insertParagraph()`, `InsertLocation` enum, anchor lookup
**Prerequisites:** Problems 1–5
**Time budget:** ~45 minutes

## Statement

Add a button labelled `Insert callout`. When clicked, it finds the first paragraph whose text starts with `Heading:` (case-sensitive) and inserts a new paragraph directly *after* it containing the text `[callout inserted by add-in]`.

If no heading paragraph is found, the task pane displays `No heading found` and no change is made.

## Constraints

- Use `body.paragraphs` collection + a filter on the first character-index, or `body.search('Heading:*', {...})` — either is acceptable.
- Insert using `paragraph.insertParagraph(text, 'After')` (or equivalent `InsertLocation.After`).
- The new paragraph must be a paragraph in its own right, not text appended to the existing one.

## Automated verification

The test file at `tests/06-insert-paragraph-at-anchor.test.ts` asserts:

- The heading paragraph is identified from the mocked paragraph collection.
- `insertParagraph` is called with `'[callout inserted by add-in]'` and location `'After'`.
- When no matching paragraph exists in the mock, `insertParagraph` is not called and the task pane shows `No heading found`.

## Manual verification

- [ ] In the document, type `Heading: Introduction` (as its own paragraph) followed by `Some body text.` on the next line.
- [ ] Click `Insert callout`.
- [ ] Document now reads: `Heading: Introduction`, then `[callout inserted by add-in]`, then `Some body text.` — three paragraphs.
- [ ] Delete the heading, click again — task pane shows `No heading found`; document unchanged.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Word.Paragraph.insertParagraph" and the `InsertLocation` enum (`Before`, `After`, `Start`, `End`, `Replace`).
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Using `insertText` instead of `insertParagraph`. The former inserts inline text; the latter inserts a new paragraph structure. The test differentiates.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
Load all paragraphs' text. Sync. Find the first with `.text.startsWith('Heading:')`. If found, call `paragraph.insertParagraph('[callout ...]', 'After')`; sync. If not, update the task pane message.
</details>

## Reference solution

`git diff solutions/6 -- app/`
