# GEMMA 4 TASK DELEGATION TEMPLATE
## ArcTool Project

Use this template when Claude delegates implementation work to Gemma 4.

---

## SYSTEM / ROLE CONTEXT TO INJECT

Read and follow:

```text
Memory/gemma4_worker_constitution.md
Memory/gemma4_error_learning_log.xml     // relevant prior entries only
CLAUDE.md                                // project technical context
.Dossier/                                // relevant bounded technical records
```

Gemma 4 is the code-generation worker/student. Claude is the architect/reviewer. Gemma must generate bounded code to spec, preserve contracts, and produce XML learning entries after corrections.

---

## TASK DELEGATION BLOCK

```text
TASK_ID: {{unique_task_id}}

OBJECTIVE:
{{one clear implementation objective}}

FEATURE_AREA:
{{Quick Dimension | Filter Manager | Coordinate | Excel to Revit | Packaging | Other}}

FILES_TO_MODIFY:
{{exact paths only}}

FILES_TO_READ_FOR_CONTEXT:
{{exact paths or excerpts supplied by Claude}}

CONTRACTS_IN_SCOPE:
{{public models/interfaces/signatures Gemma may change, or "none"}}

CONTRACTS_PROTECTED:
{{public models/interfaces/signatures Gemma must not change}}

REFERENCE_CONTEXT:
{{relevant snippets, Revit API references, existing method signatures, compiler errors}}

HARD_CONSTRAINTS:
- Modify only FILES_TO_MODIFY.
- Preserve existing public contracts unless listed in CONTRACTS_IN_SCOPE.
- Do not invent Revit API members.
- No TODOs/placeholders.
- Keep edits surgical.
- Match existing style.
- For Quick Dimension mixed L/T aggregation, do not generate or port code until Claude supplies an accepted research-and-self-critique model. The required critique covers L-L, T-T, L-T, T-L, mid-run T-joint, reversed-axis, and coincident-station cases.
- Quick Dimension always handles one selected straight wall and creates one reviewable chain only; never introduce bulk/automatic multi-wall dimension creation.

ACCEPTANCE_CRITERIA:
- {{criterion 1}}
- {{criterion 2}}
- {{criterion 3}}

OUTPUT_REQUIRED:
- ANALYSIS
- CODE
- INTEGRATION NOTES
- SELF-CHECK
- LEARNING ENTRY XML only if this is a correction turn
```

---

## CORRECTION TURN BLOCK

Use this when Gemma produced bad output and Claude is asking for a fix.

```text
CORRECTION_TURN: true
PREVIOUS_ERROR_TYPE: {{compile_error | api_misuse | contract_violation | architecture_boundary_violation | over_editing | missing_edge_case | incorrect_assumption | test_failure | revit_runtime_risk | style_or_consistency_issue}}

WHAT_WAS_WRONG:
{{specific defect in Gemma's prior answer}}

WHY_IT_WAS_WRONG:
{{contract/API/architecture/compiler reason}}

CORRECTIVE_INSTRUCTION:
{{specific instruction for the corrected attempt}}

DO_NOT_REPEAT:
- {{mistake 1}}
- {{mistake 2}}

REQUIRED_LEARNING_ENTRY_ID:
{{GEMMA-ERR-####}}

OUTPUT_REQUIRED:
1. Corrected code.
2. Short self-check.
3. Complete <learning_entry> XML block using the approved schema.
```

---

## RAG PRELOAD BLOCK

Before asking Gemma to generate code, Claude should add only the most relevant prior lessons from `Memory/gemma4_error_learning_log.xml`.

```text
RELEVANT_PRIOR_LESSONS:
- {{GEMMA-ERR-####}}: {{short learning_rule + anti_pattern}}
- {{GEMMA-ERR-####}}: {{short learning_rule + anti_pattern}}
```

Keep this short. Prefer the 3-7 most relevant lessons instead of dumping the full XML log.
