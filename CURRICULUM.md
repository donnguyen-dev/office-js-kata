# Curriculum

12 problems, ordered. Word host specifically.

## Tier 1 — The proxy-load-sync mental model (problems 1–3)

This is the load-bearing conceptual lesson. If it doesn't stick in the first three problems, nothing else will.

| #  | Title                                          | Concept introduced                                    |
| -- | ---------------------------------------------- | ----------------------------------------------------- |
| 1  | Hello task pane + `Office.onReady`             | manifest, sideloading, task pane lifecycle            |
| 2  | First `Word.run` — read the body's text        | proxy objects, `.load()`, `context.sync()`            |
| 3  | Read + write in one batch                      | batching multiple ops before a single `context.sync()` |

## Tier 2 — Addressing content (problems 4–6)

Three ways to point at content: Body, Range, Selection. The three canonical operations on each.

| #  | Title                                          | Concept introduced                                    |
| -- | ---------------------------------------------- | ----------------------------------------------------- |
| 4  | Operate on the user's current selection        | `context.document.getSelection()`, `Range.font`, `Range.insertText` |
| 5  | Find & replace                                 | `Body.search()`, `RangeCollection` iteration          |
| 6  | Insert a paragraph at a specific anchor        | `Paragraph.insertParagraph()`, `InsertLocation` enum  |

## Tier 3 — Word-specific power features (problems 7–10)

Where a Word add-in earns its keep. Content controls are the single most important primitive; OOXML is the escape hatch.

| #  | Title                                          | Concept introduced                                    |
| -- | ---------------------------------------------- | ----------------------------------------------------- |
| 7  | Tables — insert, read, write specific cells    | `body.insertTable()`, `TableRow`, `TableCell`         |
| 8  | Content controls — insert + retrieve by tag    | `Range.insertContentControl()`, `ContentControl.tag`, `getByTag()` |
| 9  | Track changes — toggle and accept              | `Document.changeTrackingMode`, `Body.getRevisions()`  |
| 10 | Insert OOXML directly (the escape hatch)       | `Range.insertOoxml()`; when the typed API isn't expressive enough |

## Tier 4 — Integration concerns (problems 11–12)

Every real add-in needs these, and they're rarely tutorialised well.

| #  | Title                                          | Concept introduced                                    |
| -- | ---------------------------------------------- | ----------------------------------------------------- |
| 11 | Persist user settings across sessions          | `Office.context.document.settings.set/get/saveAsync`  |
| 12 | Dialog API + auth round-trip                   | `Office.context.ui.displayDialogAsync`, `messageChild`, `messageParent`, token handling |

## Optional extensions (not in v1)

- **Events** — `SelectionChanged`, `ContentControlEntered`. Useful but not load-bearing for v1; can be picked up on real work.
- **Custom XML Parts** — powerful for structured templates but domain-specific; skip until you actually need it.
- **Custom Functions** — a whole separate add-in flavour, mostly Excel; only relevant if you're doing Word-specific formula work.

## Why this order

- **Tier 1** cements the single most important mental model: *nothing is real until `context.sync()`*. Every subsequent problem assumes this fluency.
- **Tier 2** exercises the three ways to address content, each in a small realistic scenario.
- **Tier 3** touches the Word-specific power features. Content controls (problem 8) is the most important single problem — serious Word add-ins live and die by how well they use them.
- **Tier 4** covers the cross-cutting concerns every real add-in eventually needs.

## Cumulative shape

By problem 12 you'll have built, across `app/`, what is effectively a small productivity add-in — read/write document content, react to user selection, structure regions via content controls, persist per-document settings, and open a dialog for auth. Not a complete product; a foundation you could extend.
