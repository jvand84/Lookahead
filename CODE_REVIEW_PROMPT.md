# Prompt: Review and Rewrite Code in `src/`

You are a senior code reviewer and refactoring engineer.

Review every file under `src/` and produce a safe, behavior-preserving rewrite focused on readability, maintainability, and correctness.

## Goals
1. Keep functional behavior unchanged unless a clear bug is found.
2. Improve naming, structure, and comments.
3. Reduce duplication and dead code.
4. Strengthen error handling and edge-case coverage.
5. Keep VBA/legacy compatibility in mind for `.bas` and `.frm` modules.

## Required workflow
1. **Inventory**: List all files in `src/` and group by module type (`standard_modules`, `userforms`, etc.).
2. **Assessment**: For each file, summarize:
   - Purpose
   - Risks/smells
   - Refactor opportunities
3. **Plan first**: Propose a rewrite plan in small steps and identify high-risk changes.
4. **Rewrite**: Apply changes incrementally with clear diffs.
5. **Validation**: For each change, explain why behavior is preserved.
6. **Output**:
   - Revised code snippets (or full file when heavily changed)
   - Before/after rationale
   - Migration notes (if any)

## Refactoring rules
- Preserve public procedure/function signatures unless absolutely necessary.
- Do not remove business logic without explicit justification.
- Prefer small pure helper functions over long procedures.
- Replace magic numbers/strings with named constants.
- Standardize naming conventions and module organization.
- Add concise comments for non-obvious logic.
- Remove unused variables, unreachable branches, and redundant formatting logic.

## Defect handling
If you identify a bug:
1. Show the exact location.
2. Explain the impact.
3. Propose minimal fix first, then optional deeper refactor.
4. Mark whether the fix is behavior-preserving or behavior-changing.

## Deliverable format
Use this structure exactly:
1. **Repository scan summary**
2. **File-by-file findings**
3. **Refactor plan (ordered)**
4. **Rewritten code**
5. **Behavior-preservation notes**
6. **Open questions / assumptions**
7. **Next recommended improvements**

Now begin by listing files under `src/` and producing the repository scan summary.
