# Problem 7: Tables — insert, read, write specific cells

**Tier:** 3
**Concepts:** `body.insertTable()`, `TableRow`, `TableCell`, cell indexing
**Prerequisites:** Problems 1–6
**Time budget:** ~50 minutes

## Statement

Add a button labelled `Insert 3×3 table`. When clicked, it inserts a 3-row × 3-column table at the end of the document body with the header row `Name | Role | Location` and the two data rows:

```
Alice | Engineer | Sydney
Bob   | Lawyer   | Melbourne
```

Additionally, add a button labelled `Read cell 2,1`. When clicked, it reads the value of the cell at row index 2 (zero-based; the `Bob` row), column index 1 (`Lawyer`), and displays it in the task pane in a `<span>` with id `cell-output`.

## Constraints

- Use `body.insertTable(rows, columns, 'End', values)` where `values` is a 2D array.
- Cell access uses `table.rows.items[i].cells.items[j]`.
- The read button must load the target cell's `body.text` (cells are `Body`-like containers, not flat strings).

## Automated verification

The test file at `tests/07-tables.test.ts` asserts:

- `body.insertTable` is called with `(3, 3, 'End', <3x3 array>)`.
- The array contains exactly the specified header and data rows.
- The read handler navigates `table.rows.items[2].cells.items[1]` and calls `.body.load('text')` before syncing.
- After sync, the task pane's `cell-output` shows `Lawyer`.

## Manual verification

- [ ] Click `Insert 3×3 table`. A table with 3 columns and 3 rows appears at the end of the document with the specified content.
- [ ] Click `Read cell 2,1`. Task pane's `cell-output` shows `Lawyer`.
- [ ] Manually edit cell (2,1) in Word to `Attorney`. Click `Read cell 2,1` again — task pane now shows `Attorney`.

## Hints

<details>
<summary>Hint 1 — what to look up first</summary>
Microsoft Learn: "Word.Body.insertTable" and "Work with tables using the Word JavaScript API".
</details>

<details>
<summary>Hint 2 — common mistake</summary>
Treating cell content as a simple string. `cell.body` is a `Body` object; you must `.load('text')` on it and sync before reading.
</details>

<details>
<summary>Hint 3 — sketch of approach</summary>
Insert: `body.insertTable(3, 3, 'End', [['Name','Role','Location'],['Alice','Engineer','Sydney'],['Bob','Lawyer','Melbourne']])`. Read: navigate to the target cell via `.rows.items[2].cells.items[1]`, load `.body.text`, sync, display.
</details>

## Reference solution

`git diff solutions/7 -- app/`
