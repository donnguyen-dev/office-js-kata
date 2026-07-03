# Problem 12: Dialog API + auth round-trip

**Tier:** 4
**Concepts:** `Office.context.ui.displayDialogAsync`, `messageChild`, `messageParent`, token handling
**Prerequisites:** Problems 1–11
**Time budget:** ~55 minutes

## Statement

Add a button labelled `Sign in`. When clicked, it opens a dialog hosted at `/dialog.html` (a static page in your `app/`) which:

1. Renders a fake sign-in form (two inputs and a submit).
2. On submit, generates a fake access token (any random string is fine — `Math.random().toString(36)`).
3. Posts the token back to the task pane via `Office.context.ui.messageParent(token)`.
4. Closes the dialog.

The task pane receives the token and stores it in memory in a variable `accessToken`. It also displays `Signed in ✓ (token: ${firstFour}...)` in a `<span id="auth-status">` showing only the first 4 characters of the token for verification.

## Constraints

- Dialog is opened via `Office.context.ui.displayDialogAsync(url, options)`.
- The dialog page uses `Office.onReady(() => { ... })` before calling `messageParent`.
- The task pane listens via `dialog.addEventHandler(Office.EventType.DialogMessageReceived, callback)`.
- The dialog closes itself after messaging (dialog.close on the taskpane side once message is received).
- No real network calls — this is a message-flow exercise only.

## Automated verification

The test file at `tests/12-dialog-api-auth.test.ts` asserts:

- `displayDialogAsync` is called with the `/dialog.html` URL.
- The taskpane-side handler subscribes to `DialogMessageReceived`.
- When the mock fires the event with a fake token, `accessToken` is set to that token.
- The `auth-status` element shows the correct prefix + ellipsis format.
- After the message is received, `dialog.close()` is called.

## Manual verification

- [ ] Click `Sign in`. A dialog opens showing the sign-in form.
- [ ] Enter any values and submit. The dialog closes.
- [ ] Task pane's `auth-status` updates to `Signed in ✓ (token: <4chars>...)`.
- [ ] Refresh the task pane — `auth-status` reverts to unauthenticated (token is only in memory, as expected).

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Office.Dialog object" and "Use the Office Dialog API in Office Add-ins". Also the DialogMessageReceived event.
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Forgetting to call `Office.onReady` inside the dialog page before `messageParent`. The message will fire before Office.js is ready inside the dialog frame, and the parent won't receive it.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
Taskpane: `Office.context.ui.displayDialogAsync(url, options, (result) => { const dialog = result.value; dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => { accessToken = arg.message; updateAuthStatus(); dialog.close(); }); });`. Dialog: `Office.onReady(() => { submitBtn.onclick = () => { const token = generateToken(); Office.context.ui.messageParent(token); }; });`.
</details>

## Reference solution

`git diff solutions/12 -- app/`

## Cumulative shape at problem 12

If you've done all 12, you have a working add-in that reads and writes document content, reacts to selection, uses content controls, exercises the OOXML escape hatch, persists per-document settings, and rounds-trips an auth flow via the Dialog API. Not a shippable product, but every primitive a shippable product needs.
