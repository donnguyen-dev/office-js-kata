# Problem template

Every `problems/N-{slug}.md` file follows this format.

---

```markdown
# Problem N: {Title}

**Tier:** {1–4}
**Concepts:** {comma-separated list — e.g. "proxy-load-sync, context.sync()"}
**Prerequisites:** Problems 1..N-1
**Time budget:** ~45 minutes

## Statement

{1–2 paragraphs describing what to build. Concrete and unambiguous about the externally observable outcome — what the user sees in Word after your code runs.}

## Constraints

- {Must use X — e.g. "Must use `context.sync()` explicitly; don't rely on implicit sync."}
- {Must not use Y — e.g. "Must not access `body.text` before calling `.load('text')`."}
- {Quality bar — e.g. "TypeScript must compile with no `any`."}

## Automated verification (vitest)

The test file is at `tests/N-{slug}.test.ts`. Vitest runs the mocked Office.js host and asserts:

- {Bullet 1 — a specific behaviour the mock can verify, e.g. "`body.insertText` was called with the correct arguments in the correct sequence."}
- {Bullet 2}
- {Bullet 3}

## Manual verification (Word Web / Word Desktop)

`./verify N` prints this checklist after the vitest suite passes. You confirm each step yourself in real Word.

- [ ] {Bullet 1 — an observable step in Word, e.g. "Task pane shows 'Ready' after loading."}
- [ ] {Bullet 2 — e.g. "Clicking the button inserts the expected text at the cursor position."}
- [ ] {Bullet 3 — e.g. "Undo returns the document to its prior state."}

Run with: `./verify N`

## Hints (optional — read only if stuck > 15 minutes)

<details>
<summary>Hint 1 — what to look up first</summary>
{One sentence pointing at a doc page or concept. NOT the answer.}
</details>

<details>
<summary>Hint 2 — common mistake</summary>
{One sentence about a common gotcha for this problem.}
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
{2–3 sentences of pseudocode. Still not the answer.}
</details>

## Reference solution

After your `./verify N` passes:

```
git diff solutions/N -- app/
```

Read the differences. For each non-cosmetic difference, explain to the agent (in plain English) why the reference chose that shape over yours. The agent must push back if your explanation is wrong or hand-wavy.
```

---

## Notes for the curriculum generator

- **Statement** must be testable — either via the mock (proxy call sequence, argument shape) OR via a specific observable Word behaviour on the manual checklist.
- **Automated bullets** must be things `@microsoft/office-addin-mock` can actually verify. If a behaviour needs real Word to show up, put it in the manual checklist, not the automated section.
- **Manual bullets** are trust-based. Keep them concrete and unambiguous — the learner should not have to interpret them.
- **Hints** are progressive disclosure. Hint 1 points at a doc page; hint 2 names a common mistake; hint 3 sketches shape. Never include a complete implementation.
- **Reference solution** lives in `solutions/N`. Must pass both the vitest suite and the manual checklist (learner-confirmed on first pass).
