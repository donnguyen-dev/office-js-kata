# Problem 2: First `Word.run` — read the body's text

**Tier:** 1
**Concepts:** proxy-load-sync, `Word.run`, `context.sync()`
**Prerequisites:** Problem 1
**Time budget:** ~45 minutes

## Statement

Add a button to the task pane labelled `Read body`. When clicked, it reads the current document's full body text via `Word.run(...)` and displays that text inside a `<pre>` element with id `body-output` in the task pane.

Your code must use the proxy-load-sync pattern: get the body proxy, `.load('text')`, `await context.sync()`, then access `body.text`. This is the load-bearing mental model of Office.js. **Access `body.text` before syncing and you get `undefined`.**

## Constraints

- Must use `Word.run(async (context) => { ... })`.
- Must explicitly call `.load('text')` before reading.
- Must call `await context.sync()` before accessing `body.text`.
- Must not use `.load(true)` (loads everything — wasteful) or omit `.load` entirely.

## Automated verification (vitest)

The test file at `tests/02-word-run-read-body.test.ts` asserts:

- The click handler calls `Word.run(...)`.
- Inside `Word.run`, `body.load` is called with the argument `'text'`.
- `context.sync` is awaited before the handler reads `body.text`.
- The returned text is written to the DOM element with id `body-output`.

## Manual verification

- [ ] With a document containing `The quick brown fox` (add via typing in Word), click `Read body` in the task pane.
- [ ] `body-output` in the task pane shows exactly `The quick brown fox`.
- [ ] Change the document body to a different string, click again — the task pane now shows the new string.
- [ ] Check the browser dev tools (Word Web) console for zero errors during the click.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Word.run" and "Working with proxy objects in Office Add-ins". Specifically the "load, sync, use" sequence.
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Trying to read `body.text` before `context.sync()`. The proxy hasn't been populated yet; `body.text` will be `undefined` and TypeScript won't complain.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
`await Word.run(async (context) => { const body = context.document.body; body.load('text'); await context.sync(); document.getElementById('body-output').textContent = body.text })`. Three steps: get proxy, load specific properties, sync, then read.
</details>

## Reference solution

After `./verify 2` passes:

```
git diff solutions/2 -- app/
```

The reference tightens error handling and shows the idiomatic `try/catch` around `Word.run` that professional add-ins always include.
