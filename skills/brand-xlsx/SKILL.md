---
name: brand-xlsx
description: >-
  Brand-aware Excel engine. Use to (1) EXTRACT a company's brand from a .xlsx template into a
  reusable "Brand Profile", (2) COMPREHEND the template with the model (optional), (3) VERIFY it,
  (4) GENERATE a new on-brand .xlsx from a GridDocument fill manifest. Trigger on "extract our
  brand", "use our workbook template", "generate a branded workbook from our profile", or when a
  ./brand-kit exists. For one-off spreadsheet edits with no saved brand profile, use the normal
  xlsx skill instead. NOT for .docx (brand-docx), .pptx (brand-pptx), or PDFs.
---

# brand-xlsx

Use this skill when the user wants reusable branded Excel/workbook generation
from a company `.xlsx` template and variable user-provided data.

This is an AI-agent skill for Codex and Claude Code. The user should describe
the workbook/model they want filled; the agent maps that request to named cells
and named regions, invokes the internal Python engine, verifies the output, and
returns the generated `.xlsx`.

## The four verbs

Every brand skill (`brand-docx`, `brand-pptx`, `brand-xlsx`) implements the same
contract: **extract / comprehend / verify / generate**.

| Verb | Input | Output |
|---|---|---|
| **extract** | a company `.xlsx` template | a reusable Brand Profile |
| **comprehend** *(optional, model-driven)* | a saved profile + a model-authored `comprehension.json` | the profile with a validated, cached `comprehension` block |
| **verify** | a saved Brand Profile | QA findings + a verdict |
| **generate** | data (a GridDocument) + a profile | a new on-brand `.xlsx` |

`comprehend` is **optional**: `generate` works on the deterministic profile alone.
See [reference/comprehension.md](reference/comprehension.md) for the full step.

## Hard Rules

- Treat `python scripts/brandkit/cli.py ...` as an internal engine command, not the user-facing workflow.
- Extract opens the source template read-only and saves `brand-kit/<name>/template/shell.xlsx` byte-for-byte.
- Generate opens the saved shell and resolves every named cell/region through `profile.json`.
- Do not put style names, colors, fonts, or brand identifiers in a GridDocument.
- If the user did not provide a template or enough data, ask for the missing input.
- Return the generated file path plus a QA summary.
- Consult `profile.json.artifact_catalog` before generation when the user asks to mimic a specific piece of the template.

## Agent Workflow

1. Determine the brand name and locate the user-provided `.xlsx` template.
2. If no matching `brand-kit/<name>` exists, **extract** one.
3. **Comprehend** the template (optional, model-driven) — see below. Skip when a
   current comprehension is already cached or no model is available.
4. Convert the user's tabular/model data into `GridDocument` JSON.
5. **Generate** the `.xlsx` with the internal engine.
6. Run **QA** and report any warnings honestly.

Before generation, inspect `profile.json.artifact_catalog` when the user asks
to mimic a specific workbook piece. It records OOXML parts, named ranges,
formulas, sheet dimensions, table names, merged cells, row/column sizing, cell
styles, and number formats.

## Internal Extract

```bash
python scripts/brandkit/cli.py extract --name <brand> --template <template.xlsx> --scope project
```

## Internal Comprehend (optional, model-driven)

Read [reference/comprehension.md](reference/comprehension.md) for the full
guidance, the four questions, and the anti-overfitting directive. In short:

```bash
python scripts/brandkit/cli.py comprehend-input --name <brand>   # prints {facts, excerpt} for the model
python scripts/brandkit/cli.py comprehend --name <brand> --input comprehension.json  # the ONLY writer
```

Skip this verb when `comprehension.status` is `present` **and** its
`source_shell_sha256` equals the live `provenance.shell.sha256`. Never re-run it
at generate time.

> **xlsx readiness.** The Excel extractor does not yet surface cover-anchor or
> derived-index inventories, so those questions have no ids to bind to and
> `comprehension.status` stays `absent` — `generate` runs the deterministic path.
> Do not force a cover or index shape onto an empty inventory; a ref into an empty
> inventory is fail-closed and will be rejected. Deeper xlsx comprehension (region
> geometry, demo classification keyed to named regions) lands on the xlsx
> fact-enrichment milestone.

## Internal Verify

```bash
python scripts/brandkit/cli.py verify --name <brand> --scope auto --qa auto
```

## Internal Generate

```bash
python scripts/brandkit/cli.py generate --name <brand> --input <grid-document.json> --output <output.xlsx> --scope auto --qa auto
```

See `reference/comprehension.md`.

## Current Guarantees and Limits

M2 fills named cells and named regions while preserving formulas and workbook
topology in the shell. Region fills that exceed the named range are refused
before saving. Cover/index comprehension is staged behind the xlsx
fact-enrichment milestone; until then comprehension stays `absent` and generation
uses the proven deterministic path.
