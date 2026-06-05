# Example templates

Three **100% synthetic, DocuSkills-branded** Office templates - one per format - that
showcase what the skills extract and respect. They contain **no proprietary or customer
content**: each is generated from scratch by the reproducible builder beside it, so the
binaries committed here are safe by construction.

| Template | Built by | What it stresses |
|---|---|---|
| [`templates/docuskills_template.docx`](templates/docuskills_template.docx) | [`builders/build_docuskills_docx.py`](builders/build_docuskills_docx.py) | multi-slot cover, three indexes (TOC + table + figure), real `numbering.xml` (2-level bullets + numbered list), a custom **DocuSkills Table** style, `SEQ` captions, a boxed **DocuSkills Callout**, a header logo, two sections, a footnote |
| [`templates/docuskills_template.pptx`](templates/docuskills_template.pptx) | [`builders/build_docuskills_pptx.py`](builders/build_docuskills_pptx.py) | multi-placeholder cover, the deck's real masters & layouts, sections, an agenda slide, a native table + chart, a picture, a demo slide |
| [`templates/docuskills_template.xlsx`](templates/docuskills_template.xlsx) | [`builders/build_docuskills_xlsx.py`](builders/build_docuskills_xlsx.py) | five sheets, named regions, cross-sheet formulas, number formats, named cell styles, a table, conditional formatting, frozen panes, a chart |

Brand palette (cohesive with the project hero): DocuSkills navy `#16213F`, blue `#2B7CD3`,
amber `#E0742B`, on light `#EAF1FF` / band `#DCE7FF`.

## Try them

```bash
# Extract a Brand Profile from a template, then generate an on-brand document of the SAME format
python scripts/brandkit/cli.py extract  --name demo --template examples/templates/docuskills_template.docx --scope project
python scripts/brandkit/cli.py verify   --name demo --scope auto --qa auto
python scripts/brandkit/cli.py generate --name demo --input your_content.json --output out.docx --scope auto --qa auto
```

(Use the matching `.pptx` / `.xlsx` template to drive `brand-pptx` / `brand-xlsx`. Each
skill stays in its own lane - a Word template makes Word documents, never a deck or a sheet.)

## Regenerate

```bash
python examples/builders/build_docuskills_docx.py
python examples/builders/build_docuskills_pptx.py
python examples/builders/build_docuskills_xlsx.py
```

The builders are deterministic: all timestamps are pinned, so rebuilding yields
byte-identical templates and the regenerate step leaves a clean git tree. They are adapted from
the synthetic complex fixtures under `tests/fixtures/builders/`, re-themed with the
DocuSkills brand.
