# office-js-kata

> A problem-set approach to learning Office.js for Microsoft Word add-in development.

A small kata — 12 problems, not 20 — because Office.js is **API-deep, not concept-deep.** Once you internalise the proxy-load-sync pattern (problems 1-3), the rest is API surface you look up. The kata drills the load-bearing mental model on real Word behaviour, plus the Word-specific power features (content controls, OOXML insertion, custom XML parts) that separate serious Word add-in engineers from tourists.

**This kata is configured for Microsoft Word specifically.** The Office.js kata format is host-neutral — the same structure would work for Excel or Outlook — but the curriculum is Word.

---

## How this repo is meant to be used

Two workflows, like framework-kata:

- **Human learner:** see [§ For learners](#for-learners) below.
- **AI agent** scaffolding the curriculum: see [§ For the curriculum generator](#for-the-curriculum-generator) below.

---

## Repository structure

```
office-js-kata/
├── README.md              ← this file
├── CURRICULUM.md          ← the 12 problems, ordered
├── PROBLEM_TEMPLATE.md    ← format every problems/N-*.md follows
├── VERIFY_TEMPLATE.md     ← the vitest + manual-checklist verify contract
├── verify                 ← shell script: ./verify N — runs vitest + prints checklist
├── problems/              ← 12 problem statements
├── tests/                 ← per-problem vitest specs (once app/ is scaffolded)
└── app/                   ← the Office.js Word add-in project (NOT scaffolded yet — see setup below)
```

---

## Setup (do this once)

**Prerequisites:**

- Node 20+ (`node --version`)
- npm or pnpm
- A Microsoft 365 account with permission to sideload add-ins (personal M365 works; work accounts may require IT approval)
- Microsoft Word — either Word Web (fastest dev loop; open at [office.com](https://www.office.com/)) or Word Desktop

**Scaffold the add-in project (interactive; do this on your Mac, not via agent):**

```bash
cd office-js-kata
npm install -g yo generator-office
yo office
# Answer the prompts:
#   Project type:     Office Add-in Task Pane project
#   Script type:      TypeScript
#   Name of project:  app
#   Office client:    Word
```

That creates `office-js-kata/app/` with a working "Hello World" Word add-in, a manifest.xml, Vite/webpack build config, and sideload tooling.

**Add the kata's test dependencies:**

```bash
cd app
npm install --save-dev vitest office-addin-mock
```

**Sideload the add-in (first time only):**

- Word Web: from Word, `Insert → Add-ins → Upload My Add-in → Browse → select app/manifest.xml`.
- Word Desktop: `Insert → My Add-ins → Upload My Add-in → Browse → manifest.xml`.

You now have a task pane in Word showing your dev-time add-in. Every code change hot-reloads.

**Test the setup:**

```bash
cd office-js-kata
./verify 1
```

Runs the vitest suite for problem 1 and prints the manual checklist. If both pass, you're ready.

---

## For learners

### Daily loop

For problem N:

```bash
git checkout main
git checkout -b attempts/N
cat problems/N-*.md               # read the problem statement
# Write your code in app/ — as far as you can without external help
./verify N                        # runs vitest + prints manual checklist
# Follow the checklist in Word Web / Word Desktop
# Once both pass, diff against reference (once available):
git diff solutions/N -- app/
```

### The verify contract

Different from framework-kata. Office.js can't be fully automated because real Word behaviour (event timing, selection state, content control lifecycle) doesn't reliably reproduce in mocks.

Every `verify N` does two things:

1. **Runs the vitest suite** at `tests/N-{slug}.test.ts` against `office-addin-mock`. This catches the pure-logic mistakes.
2. **Prints a manual checklist** — 3-5 steps you perform in real Word (Web or Desktop) to confirm host behaviour. You confirm each step yourself; there's no automated gate.

The manual checklist is trust-based. Don't skip it. The whole point of Office.js is that mocks can't tell you whether your add-in *actually works in Word.*

### The rules (self-enforced)

1. **No peeking.** Don't look at `solutions/N` until your `./verify N` passes both halves.
2. **No agent writes app code during the daily loop.** Agent for narrow lookups only.
3. **Sideload real Word.** Don't be tempted to skip the manual checklist. It's the whole point.

### Pacing

12 problems, roughly 1 per session (~45 min each). ~10-14 evenings to complete. Aim to compress on second pass if you re-do the kata later.

---

## For the curriculum generator

*(This section is written for an AI agent tasked with expanding this scaffold into a complete kata. Skip if you're a human learner.)*

The scaffold in this initial commit contains:

- README.md (this file)
- CURRICULUM.md (12 problems listed)
- PROBLEM_TEMPLATE.md (format for problems/N-*.md)
- VERIFY_TEMPLATE.md (format for tests/N-*.test.ts + manual checklist)
- verify script (orchestration only; individual test files not yet written)
- 12 problem statements in problems/

Not yet done (agent-generation targets):

- The 12 vitest test files in `tests/N-{slug}.test.ts` following VERIFY_TEMPLATE.md.
- Solution branches (`solutions/01` … `solutions/12`) with reference implementations.
- Fully worked scaffold for the Word add-in project (deferred to human because `yo office` is interactive).

To generate the rest: read the problem files, follow VERIFY_TEMPLATE.md to write per-problem tests, verify each fails on main and passes on its solution branch. Same rigour gates as framework-kata.

---

## Status

Scaffold only. Problem statements complete; test files and reference solutions pending. See [CURRICULUM.md](./CURRICULUM.md) for the full problem list.
