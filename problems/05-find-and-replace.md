# Problem 5: Find & replace

**Tier:** 2
**Concepts:** `Body.search()`, `RangeCollection` iteration, collection load semantics
**Prerequisites:** Problems 1–4
**Time budget:** ~45 minutes

## Statement

Add two text inputs (`find`, `replace`) and a button labelled `Replace all`. When clicked, it finds every case-insensitive occurrence of the `find` string in the document body and replaces each with the `replace` string.

Report the count of replacements in a task pane element with id `replace-count`, formatted as `Replaced N occurrence(s)`.

## Constraints

- Use `body.search(query, { matchCase: false })`.
- Iterate the returned `RangeCollection` and call `insertText` on each range.
- Collection load semantics: you must `.load()` the collection's `items` (or specific properties) before iterating, then sync.

## Automated verification

The test file at `tests/05-find-and-replace.test.ts` asserts:

- `body.search` is called with the query from the `find` input and `{ matchCase: false }`.
- The returned collection has `.load('items/text')` (or equivalent) called on it.
- `insertText` is called `N` times where `N` is the mock's returned range count.
- The task pane's `replace-count` element shows the exact string `Replaced N occurrence(s)`.

## Manual verification

- [ ] Type `the quick brown fox jumps over the lazy dog. THE END.` into the document.
- [ ] Find: `the`, Replace: `a`, click `Replace all`.
- [ ] Document reads `a quick brown fox jumps over a lazy dog. a END.`.
- [ ] Task pane shows `Replaced 3 occurrence(s)`.
- [ ] Ctrl+Z reverts all replacements (Word batches them).

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Body.search method" and "Search options" (matchCase, matchWildcards, matchWholeWord).
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Forgetting to `.load()` the collection before iterating. `results.items` will be an empty array if you skip this step.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
`const results = body.search(findText, { matchCase: false }); results.load('items'); await context.sync(); results.items.forEach(range => range.insertText(replaceText, 'Replace')); await context.sync();`.
</details>

## Reference solution

`git diff solutions/5 -- app/`
