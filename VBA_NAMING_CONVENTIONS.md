# VBA Naming & Convention Standards

This document defines naming and style standards for modules, procedures, variables, constants, and userforms in this repository.

## 1) Goals
- Make code easier to read and maintain.
- Reduce ambiguity in shared modules.
- Keep behavior-preserving refactors low-risk by using predictable naming.

## 2) File and module naming

### Standard modules (`src/standard_modules`)
- Use `mod` prefix + PascalCase noun phrase.
- Format: `mod<DomainOrCapability>.bas`
- Examples:
  - `modManning.bas`
  - `modGuardsAndTables.bas`
  - `modRegressionSmoke.bas`

### UserForms (`src/userforms`)
- Use `frm` prefix + PascalCase noun phrase.
- Format: `frm<Feature>.frm` (paired `.frx` asset).
- Examples:
  - `frmSelect.frm`
  - `frmMessages.frm`

## 3) Procedure naming

### Public entry points (macros called by users/buttons)
- Use verb-first PascalCase that describes the action.
- Format: `<Verb><Object><OptionalQualifier>`
- Examples:
  - `RefreshAllQueries`
  - `BuildRosterLegend`
  - `RunPublicEntryPointRegression`

### Private helper procedures
- Use verb-first PascalCase.
- Prefix with `Try` when returning success/failure pattern.
- Prefix with `Ensure` when creating or validating prerequisite state.
- Examples:
  - `TryGetTable`
  - `EnsureScratchSheet`

### Boolean functions
- Prefix with `Is`, `Has`, `Can`, or `Should`.
- Examples:
  - `IsOffShift`
  - `HasTable`

## 4) Variable naming

### General rule
- Use meaningful camelCase names for local variables and parameters.
- Avoid single-letter names except short loop indices (`i`, `j`).

### Approved short forms
- `ws` = Worksheet
- `wb`/`wbk` = Workbook
- `lo`/`tbl` = ListObject
- `rng` = Range

### Legacy prefixes
- Existing `v*` names are tolerated in unchanged legacy code.
- For new code and touched code, prefer descriptive names over generic `v*` prefixes.
  - Prefer `isGoAll` over `vGoAll`
  - Prefer `managerName` over `strManager`

## 5) Constants and enums

### Constants
- Use `UPPER_SNAKE_CASE` for module-level constants.
- Keep units or scope clear in name.
- Examples:
  - `DEFAULT_SHEET_NAME`
  - `MAX_CONSECUTIVE_SHIFTS`

### User-defined types / enums
- Use PascalCase with `T` prefix for custom types.
- Examples:
  - `TAppGuardState`
  - `TSheetGuardState`

## 6) Object naming in worksheets/tables
- Table names: `tbl<Domain><Purpose>` (PascalCase after prefix)
  - Example: `tblLookahead`
- Named ranges: `rng<Domain><Purpose>`
- Avoid spaces and punctuation in names.

## 7) Error handling conventions
- Use `Const PROC_NAME As String = "<ProcedureName>"` in non-trivial procedures.
- Use one local error handler label: `ErrHandler` or `FailCase`.
- Route errors to centralized logging where available (`LogError`).
- Prefer guard clauses and early exits over deep nesting.

## 8) Formatting conventions
- Require `Option Explicit` in every module.
- Indentation: 4 spaces.
- Keep line length reasonable; wrap long argument lists across lines.
- Group sections with short headers in long procedures:
  - validation
  - setup
  - main logic
  - cleanup

## 9) Public API stability rules
- Do not rename public procedures/functions without a migration step.
- If a rename is required:
  1. Keep old wrapper temporarily.
  2. Mark it with a deprecation comment.
  3. Update all call sites before final removal.

## 10) Migration guidance for this repository
- Apply standards incrementally during touched-file refactors.
- Do not perform repository-wide rename-only commits without functional changes.
- Prioritize high-churn modules first (`modManning`, `modFormatting`, `modAwayRoster`).

## 11) Pull request checklist (naming/convention scope)
- [ ] New modules/forms follow `mod*` / `frm*` prefixes.
- [ ] New public procedures are verb-first PascalCase.
- [ ] New boolean names use `Is/Has/Can/Should`.
- [ ] New constants use `UPPER_SNAKE_CASE`.
- [ ] `Option Explicit` present in touched VBA modules.
- [ ] No unnecessary public API renames.
