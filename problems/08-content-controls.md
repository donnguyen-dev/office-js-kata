# Problem 8: Content controls — insert + retrieve by tag

**Tier:** 3
**Concepts:** `Range.insertContentControl()`, `ContentControl.tag`, `getByTag()`
**Prerequisites:** Problems 1–7
**Time budget:** ~55 minutes

## Statement

**This is the most important problem in the kata.** Content controls are how serious Word add-ins mark structured regions of a document. Contract review, template automation, redlining assistance — all built on content controls.

Add a button labelled `Mark selection as amount`. When clicked, it wraps the user's current selection in a rich-text content control with `tag = 'invoice-amount'` and `title = 'Invoice Amount'`. If the selection is empty, display `Select some text first` in the task pane.

Add a button labelled `Find amount`. When clicked, it locates the content control tagged `invoice-amount`, reads its text, and displays it in a task pane element with id `amount-output`. If no such control exists, display `No amount tagged`.

## Constraints

- Use `range.insertContentControl()` (returns the new control).
- Set `.tag` and `.title` on the control before sync.
- Retrieval uses `context.document.contentControls.getByTag('invoice-amount')` — returns a `ContentControlCollection`, always non-null even if empty.
- Handle the "collection is empty" case correctly.

## Automated verification

The test file at `tests/08-content-controls.test.ts` asserts:

- `range.insertContentControl` is called on non-empty selection.
- The resulting control has `tag = 'invoice-amount'` and `title = 'Invoice Amount'` assigned before sync.
- On the find handler, `contentControls.getByTag('invoice-amount')` is called.
- If the mock returns a non-empty collection, the first control's text is loaded and written to `amount-output`.
- If the mock returns an empty collection, `amount-output` shows `No amount tagged`.

## Manual verification

- [ ] Type `Total: $1,250.00` in the document. Select `$1,250.00`.
- [ ] Click `Mark selection as amount`. The selection is now surrounded by a content control marker (visible in Word Web as a box around the text with a small tab).
- [ ] Click `Find amount`. Task pane's `amount-output` shows `$1,250.00`.
- [ ] Change the text inside the content control to `$999.99` manually. Click `Find amount` again — task pane shows `$999.99`.
- [ ] Delete the content control (right-click → Remove Content Control), click `Find amount` — task pane shows `No amount tagged`.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Content controls in Word add-ins" and specifically "getByTag" vs "getByTitle".
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Assuming `getByTag` returns a single control. It returns a *collection* — always non-null, sometimes empty. Check `.items.length` after loading.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
Insert: `const cc = range.insertContentControl(); cc.tag = 'invoice-amount'; cc.title = 'Invoice Amount'; await context.sync();`. Retrieve: `const controls = context.document.contentControls.getByTag('invoice-amount'); controls.load('items/text'); await context.sync(); if (controls.items.length === 0) { /* none */ } else { display controls.items[0].text }`.
</details>

## Why this matters most in the kata

Every non-trivial Word add-in eventually reaches for content controls. They are the extensibility primitive that gives you *addressable regions* of a document — the way you mark "here is the invoice total" or "here is the party name" so your code can find it reliably even after the user reformats surrounding text. Contract-review add-ins live entirely on content controls. So do template generators. So does anything that needs to persist metadata alongside document content.

If you take one thing from this kata into production work, it's fluency with content controls.

## Reference solution

`git diff solutions/8 -- app/`
