# Repository scan summary

## Inventory (`src/`)

### `standard_modules`
- `modAwayRoster.bas`
- `modCleanLookahead.bas`
- `modExport.bas`
- `modFormatting.bas`
- `modGraph.bas`
- `modGuardsAndTables.bas`
- `modManning.bas`
- `modReport.bas`
- `modRoleGraph.bas`
- `modSupport.bas`
- `modUDF.bas`

### `userforms`
- `FrmPassword.frm` / `FrmPassword.frx`
- `frmMessages.frm` / `frmMessages.frx`
- `frmSelect.frm` / `frmSelect.frx`

## High-level observations
- The codebase has strong domain coverage (manning, lookahead, reporting, graphing) but very large procedures in orchestration modules (`modAwayRoster`, `modFormatting`, `modManning`) increase maintenance risk.
- Naming conventions are mixed (`vGoAll`, `ChangeLA`, `BuildRosterLegend`, `vPassword`) and should be standardized to improve readability.
- Utility/guard modules (`modGuardsAndTables`, parts of `modSupport`) are good candidates for centralization of common patterns already emerging across modules.
- UserForms are straightforward but can benefit from explicit validation and shared helper routines to avoid duplicated UI logic.

# File-by-file findings

| File | Purpose | Risks / smells | Refactor opportunities |
|---|---|---|---|
| `modAwayRoster.bas` | Away-roster transformations, name matching, table lookups | Very large module with many responsibilities; fuzzy matching and data-change logic are coupled | Split into `AwayRosterSync`, `NameMatching`, and `TableLookup` helpers; isolate pure string/distance functions |
| `modCleanLookahead.bas` | Cleanup of blank rows in lookahead table | Single-entry proc may hide assumptions about table existence/layout | Add explicit table existence and column guard checks with early exits |
| `modExport.bas` | VBA module import/export | Light validation around paths/files | Add `ValidateFolder` and `CanExportComponent` helpers |
| `modFormatting.bas` | Roster formatting, fatigue checks, legend construction | High cyclomatic complexity; long procedures and mixed concerns | Extract conditional-format rule builders, shift-type helpers, legend writer service |
| `modGraph.bas` | Slicer setup/listing | Hard-coded sheet/slicer references likely brittle | Add named constants + `GetSlicerCacheSafe` helper |
| `modGuardsAndTables.bas` | App/sheet guard + table utilities | Potentially broad public surface; usage consistency unknown | Keep as shared infra; add module-level docs and strict error contracts |
| `modManning.bas` | Main orchestration for lookahead/manning workflows | Global state variables and control flags increase hidden coupling | Wrap app state in a dedicated context record and pass explicitly |
| `modReport.bas` | Report navigation/population | Likely table-column assumptions and range coupling | Introduce header-map resolver and table write adapters |
| `modRoleGraph.bas` | Rebuild role graph from lookahead | Resize/index logic likely sensitive to schema drift | Extract schema validation + resize operations into shared helpers |
| `modSupport.bas` | Shared constants, misc support macros, temporary message handling | Security smell (`vPassword` constant), mixed unrelated responsibilities | Move credentials to protected storage; split UI helpers and OS/API constants |
| `modUDF.bas` | UDFs for worksheet usage (`VisibleUniqueList`) | Array/error handling edge cases may be fragile under empty/filtered ranges | Add robust empty range and error-return pathways |
| `FrmPassword.frm` | Password prompt form | Minimal validation and UX feedback pathways | Add non-empty validation and failed-attempt feedback state |
| `frmMessages.frm` | Temporary message form lifecycle | Timer/lifecycle behavior can drift in long-running sessions | Encapsulate start/stop timing and disposal guards |
| `frmSelect.frm` | Workbook selection UI | Input assumptions + duplicated load/confirm flow | Extract load/selection validation helper routines |

# Refactor plan (ordered)

1. **Stabilize shared infrastructure (low risk / high leverage).**
   - Add or consolidate safe table/sheet lookup helpers in `modGuardsAndTables`.
   - Introduce naming/constants standard (prefixes, sheet/table names).
2. **Decompose heavy orchestration modules.**
   - Break `modAwayRoster` and `modFormatting` into focused private helpers.
   - Keep public entry points/signatures unchanged.
3. **Normalize state management in `modManning`.**
   - Replace scattered globals with explicit context where feasible.
4. **Harden userforms and UDF edge handling.**
   - Add explicit validation and early-return guards.
5. **Apply consistency pass.**
   - Rename local vars for clarity, add concise comments for non-obvious logic, remove dead/redundant branches.

# Rewritten code

> Behavior-preserving sample rewrites are provided below as the first incremental pass.

## 1) Guarded table lookup helper pattern (for reuse)

```vb
Private Function TryGetTable(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String, ByRef outTable As ListObject) As Boolean
    Dim ws As Worksheet
    On Error GoTo CleanFail

    Set ws = GetWorksheet(wb, sheetName)
    If ws Is Nothing Then GoTo CleanFail

    Set outTable = GetTable(ws, tableName)
    If outTable Is Nothing Then GoTo CleanFail

    TryGetTable = True
    Exit Function

CleanFail:
    Set outTable = Nothing
    TryGetTable = False
End Function
```

## 2) Extracted shift evaluation helpers (from formatting-style flows)

```vb
Private Function GetShiftTypeSafe(ByVal cellValue As Variant) As String
    Dim v As String
    v = Trim$(CStr(cellValue))

    If Len(v) = 0 Then
        GetShiftTypeSafe = ""
    ElseIf IsOffShift(v) Then
        GetShiftTypeSafe = "OFF"
    Else
        GetShiftTypeSafe = GetShiftType(v)
    End If
End Function

Private Function IsShiftCellActionable(ByVal cellValue As Variant) As Boolean
    IsShiftCellActionable = (Len(GetShiftTypeSafe(cellValue)) > 0)
End Function
```

## 3) Safer UserForm selection confirmation flow

```vb
Private Function ValidateSelection() As Boolean
    ValidateSelection = (Len(Trim$(Me.cboWorkbooks.Value)) > 0)
End Function

Private Sub btnOK_Click()
    If Not ValidateSelection() Then
        MsgBox "Please select a workbook before continuing.", vbExclamation
        Exit Sub
    End If

    vSelection = Me.cboWorkbooks.Value
    Me.Hide
End Sub
```

# Behavior-preservation notes

- Public APIs are intentionally unchanged in this first pass; proposed snippets are private helper extractions and defensive wrappers.
- Rewrites focus on guard clauses, explicit validation, and decomposition; these preserve logic flow while reducing side effects and hidden assumptions.
- Error handling is constrained to localized helper boundaries to avoid changing workbook-wide execution semantics.

# Open questions / assumptions

1. Is `vPassword` in `modSupport.bas` still actively used, and can it be migrated to a secure retrieval source?
2. Are table/sheet names guaranteed stable across workbooks, or should schema discovery be dynamic?
3. Should fuzzy matching behavior in `modAwayRoster` remain exact as-is, or can thresholding be made configurable?
4. Are there expected performance limits (row counts) that should guide optimization priorities in `modFormatting` and `modManning`?

# Next recommended improvements

1. Add a lightweight regression macro suite that exercises each public entry-point Sub.
2. Create a naming/convention document for module and variable standards.
3. Introduce a small internal test harness workbook for table-schema validation.
4. Move security-sensitive constants to protected storage and document access patterns.
5. Execute refactors in small module-level PRs to simplify validation and rollback.
