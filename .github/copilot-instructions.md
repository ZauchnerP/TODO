<!-- Copilot / AI-agent instructions for the Excel VBA To-Do project -->
# Project Snapshot

This repository is a single VBA module (`Code.bas`) that generates two Excel sheets: a multi-row To-Do list and a time-blocked Day planner. The module is intended to be imported into a macro-enabled workbook (`.xlsm`) and run from the VBA editor.

# Important Files

- `Code.bas` — main VBA module. Key entry points: `Create_TODO_sheet`, `Create_Today_sheet`, `Main_Sort`.
- `README.md` — setup, run instructions and descriptions of columns/buttons.

# Big-picture architecture / data flow

- The workbook contains one or more sheets produced by the module. Each To-Do sheet has:
  - Header rows: row 1 (visual/title) and row 2 (column headers). Data starts at row 3.
  - Columns A..H with fixed meanings: `Category`, `Importance`, `Time`, `Emotional effort`, `Dependence`, `Task`, `When`, `Hide`.
- Button shapes are created on the sheet and wired to VBA procedures via `Shape.OnAction` (e.g., shape named `Sort_All` calls `Main_Sort`).
- Several helper procedures prepare data for sorting/formatting: `Replace_Empty_Dependence` fills `E` with `.`, `Insert_0_Hide` fills `H` with `0`.
- Sorting and coloring assume those fill-steps; changing the order or removing them will change behavior.

# Developer workflows (explicit)

- Setup (developer/test):
  1. Open or create a new Excel workbook and save as `.xlsm`.
  2. Developer tab → Visual Basic → Import `Code.bas` (or paste into a module).
  3. Place cursor inside `Create_TODO_sheet()` and press Run to create a To-Do sheet.
  4. To create the day planner run `Create_Today_sheet()`.

- Buttons and multiple workbooks: if multiple workbooks contain the module, `Shape.OnAction` must be qualified to the workbook, e.g. `shape.OnAction = "'MyWorkbook.xlsm'!Function_Name"`.

# Project-specific conventions & patterns

- Column/row conventions:
  - Data rows begin at row 3 for To-Do sheets. Many routines use `ws.Cells(ws.Rows.Count, "A").End(xlUp).Row` to find lastRow.
  - `Dependence` uses `.` as an explicit placeholder; many filters and sorts depend on it.
  - `Hide` uses string numeric values (`"0"`/`"1"`) and is used with AutoFilter logic.

- Button naming: shapes are created with explicit names (`Sort_All`, `Hide_Low`, `Make_Lines_TODO`, `Sort_Time`, `Hide_Dependence`, `Show_All`, `Hide_Hide`, `Set0`, `Clean_Today`). Use these names if you need to find or delete shapes programmatically.

- Coloring and formatting:
  - Category-specific colors live in `Color_Category` (look for `Topic1/Topic2/Topic3` placeholders).
  - Conditional formatting for "today" dates is created by `Today_Red` on column G.
  - Many RGB color values are hardcoded in the file — change them there.

- Sorting behavior:
  - `Main_Sort` calls `Replace_Empty_Dependence`, `Insert_0_Hide`, then `Sort_TODO` and several coloring routines. Preserve that order unless intentionally changing behavior.

# Integration points & external dependencies

- No external packages or services. The only dependency is Excel (developer macro-enabled environment). The README states Microsoft Office LTSC Professional Plus 2021, but code is plain VBA and should work in many recent Excel versions.

# Editing tips & safe-change checklist for AI agents

- Prefer minimal, localized edits. Many routines assume columns A..H and header rows — do not change column indices globally without updating all helpers.
- If adding or renaming a button shape, also update or create the corresponding `OnAction` binding.
- When changing how `Dependence` or `Hide` are represented, update all filters/sorts/formatters that rely on `.` or `0`.
- Always recommend the user run changes in a copy of the workbook (`.xlsm`) and keep a backup before running macros.

# Examples agents may use when proposing code edits

- To add a notebook test instruction in the README: reference `Create_TODO_sheet()` as the canonical entrypoint.
- To change category names and color mapping: edit `Color_Category` and replace `Topic1/Topic2/Topic3` with actual category strings.
- To qualify `OnAction` for a workbook named `MyTodo.xlsm`: set `shp.OnAction = "'MyTodo.xlsm'!Main_Sort"` right after shape creation.

# Questions for maintainers (include these in PRs)

- Are multi-workbook scenarios expected in regular use? If yes, prefer saving `OnAction` with workbook-qualified names by default.
- Do you want category names (Topic1/2/3) codified into a small configuration section at top of `Code.bas` for easier edits?

--
Please review this file for missing project-specific notes or preferences I should include. I can iterate on wording or add examples of small patches.
