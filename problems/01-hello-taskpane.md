# Problem 1: Hello task pane + `Office.onReady`

**Tier:** 1
**Concepts:** manifest, sideloading, task pane lifecycle, `Office.onReady`
**Prerequisites:** none
**Time budget:** ~30 minutes

## Statement

Sideload your add-in into Word (Web or Desktop). Once loaded, the task pane must render a paragraph with the exact text `Ready` — but only after `Office.onReady` has fired. Before `Office.onReady` resolves, the paragraph must contain `Loading…`.

This problem is the smoke test for the whole loop. If your dev workflow (edit → hot-reload → observe in Word) doesn't work end-to-end here, fix it before proceeding.

## Constraints

- Must use `Office.onReady(...)` (not `Office.initialize`, which is the older API).
- The `Ready` text must only render after `Office.onReady`'s callback fires — not on document.body load.
- TypeScript compiles with no errors, no `any`.

## Automated verification (vitest)

The test file at `tests/01-hello-taskpane.test.ts` asserts:

- The taskpane's `renderReady()` function returns the string `'Ready'`.
- The taskpane's `renderLoading()` function returns the string `'Loading…'`.
- `Office.onReady` is called during module initialisation.

## Manual verification (Word Web / Word Desktop)

- [ ] Sideload `app/manifest.xml` into Word.
- [ ] Open the task pane from the ribbon.
- [ ] Confirm the task pane briefly shows `Loading…` (may be too fast to see reliably; check dev tools if needed).
- [ ] Confirm the task pane settles on `Ready` after Office.js has loaded.
- [ ] Reload the task pane (`Ctrl+F5` in Word Web) and confirm the sequence repeats.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Office.onReady". Note the newer promise-based signature vs. the older callback form.
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Rendering the ready state from module top level rather than inside the `Office.onReady` callback. This makes the "Ready" text show up before Office.js is actually ready to be used.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
On module load, render `Loading…` immediately. Call `Office.onReady(...)`; inside the callback, replace the DOM content with `Ready`. Two functions, one lifecycle hook.
</details>

## Reference solution

After `./verify 1` passes:

```
git diff solutions/1 -- app/
```
