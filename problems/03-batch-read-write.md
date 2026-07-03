# Problem 3: Read + write in one batch

**Tier:** 1
**Concepts:** batching multiple operations before a single `context.sync()`
**Prerequisites:** Problems 1–2
**Time budget:** ~45 minutes

## Statement

Add a button labelled `Uppercase body`. When clicked, it must (a) read the current body text, (b) transform it to uppercase, and (c) replace the body's content with the uppercased version — **all in a single `Word.run` with a single `context.sync()`.**

This problem exists to make you *feel* the batching model. Naïvely you'd write two syncs (one to read, one to write). Idiomatically you write one sync that both reads and writes.

## Constraints

- Exactly one `context.sync()` call in the handler.
- Must use `body.insertText(value, 'Replace')` for the write (not string assignment to `body.text`).
- The uppercase transformation happens in JavaScript, not via a Word function.

## Automated verification

The test file at `tests/03-batch-read-write.test.ts` asserts:

- `context.sync` is called exactly once during the click handler.
- `body.load('text')` is called before `context.sync`.
- `body.insertText` is called with `'Replace'` as the second argument.
- The value passed to `body.insertText` is the uppercase transformation of what `body.load` fetched.

## Manual verification

- [ ] With a document containing `hello world`, click `Uppercase body`.
- [ ] Document body now reads `HELLO WORLD`.
- [ ] Repeat with a document containing mixed case (`Hello, World!`) — result is `HELLO, WORLD!`.
- [ ] Ctrl+Z (undo) reverts to the original in one step, confirming Word treated the change as a single edit.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Batching operations in Office Add-ins". The whole point of the proxy model is to minimise round-trips.
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Two syncs: one to read, then upper-case in JS, then another to write. Works but violates the batching principle. Idiomatic Office.js sequences load + insert + sync in that order, once.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
Load `body.text` first. Sync (now `body.text` is populated). Compute the uppercase. Queue `body.insertText(uppercase, 'Replace')`. Sync again to commit. *Wait, that's two syncs.* The trick: for problem 3's simpler shape, you can pre-compute and issue the write in the same batch by structuring it as a single `context.sync()` between the load and insert. Look at examples of paired reads + writes in the docs.
</details>

## Reference solution

After `./verify 3` passes:

```
git diff solutions/3 -- app/
```

The reference shows the single-sync batching pattern that becomes muscle memory across all subsequent problems.
