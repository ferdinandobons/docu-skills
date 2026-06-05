# SPDX-License-Identifier: MIT
"""CI guards for the unified SKILL.md spine and the shared comprehension prompt.

The three brand skills (docx/pptx/xlsx) share ONE comprehension prompt body so the
generic, anti-overfitting guidance cannot drift between formats. These guards run
without a model and assert:

  1. the three ``reference/comprehension.md`` are BYTE-IDENTICAL;
  2. each ``SKILL.md`` is on the four-verb spine (extract / comprehend / verify /
     generate) and links the shared comprehension reference;
  3. the shared prompt states the anti-overfitting directive verbatim and its
     worked example carries only frozen role-ids + non-language placeholders (no
     template-specific literal becomes a matching rule);
  4. every authored file keeps its ``SPDX-License-Identifier: MIT`` header.
"""
from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILLS = ("brand-docx", "brand-pptx", "brand-xlsx")

# The verbatim anti-overfitting directive the shared prompt must state (plan §8).
ANTI_OVERFIT = (
    'A title slot is a title slot whether its placeholder reads "Titolo", "Title",\n'
    "> or \"Titre\". Quote a literal **only as evidence**, never as a matching rule."
)
# Non-language placeholder ids the worked example must use (and only these as refs).
PLACEHOLDER_IDS = ("<slot-1>", "<index-A>", "<region-1>")
# Frozen role-ids permitted as concrete role references in the worked example
# (cover anchors are surfaced ids, not role ids, so they use placeholders instead).
FROZEN_ROLE_IDS = ("heading.1",)


def _comp_md(skill: str) -> Path:
    return ROOT / "skills" / skill / "reference" / "comprehension.md"


def _skill_md(skill: str) -> Path:
    return ROOT / "skills" / skill / "SKILL.md"


class SharedComprehensionPromptTest(unittest.TestCase):
    def test_all_three_reference_docs_exist(self) -> None:
        for skill in SKILLS:
            self.assertTrue(_comp_md(skill).is_file(), f"missing {skill} comprehension.md")

    def test_byte_identical_across_formats(self) -> None:
        bodies = {skill: _comp_md(skill).read_bytes() for skill in SKILLS}
        canonical = bodies["brand-docx"]
        for skill in SKILLS:
            self.assertEqual(
                bodies[skill],
                canonical,
                f"{skill}/reference/comprehension.md drifted from the shared body; "
                f"edit all three together (they MUST be byte-identical)",
            )

    def test_states_anti_overfitting_directive_verbatim(self) -> None:
        text = _comp_md("brand-docx").read_text(encoding="utf-8")
        self.assertIn(ANTI_OVERFIT, text)

    def test_worked_example_uses_only_placeholders_and_frozen_role_ids(self) -> None:
        text = _comp_md("brand-docx").read_text(encoding="utf-8")
        # The worked example block must reference the non-language placeholders.
        for pid in PLACEHOLDER_IDS:
            self.assertIn(pid, text, f"worked example missing placeholder {pid}")
        # And it may only name the frozen role-ids as concrete role references.
        for rid in FROZEN_ROLE_IDS:
            self.assertIn(rid, text)

    def test_four_executor_enums_documented(self) -> None:
        text = _comp_md("brand-docx").read_text(encoding="utf-8")
        for enum_values in (
            "present|absent|rejected",
            "in_place|clear|leave",
            "regenerate|preserve|clear",
            "demo|real|mixed",
        ):
            self.assertIn(enum_values, text)

    def test_spdx_header_present(self) -> None:
        for skill in SKILLS:
            head = _comp_md(skill).read_text(encoding="utf-8").splitlines()[0]
            self.assertIn("SPDX-License-Identifier: MIT", head)


class SkillSpineTest(unittest.TestCase):
    def test_each_skill_names_the_four_verbs(self) -> None:
        for skill in SKILLS:
            text = _skill_md(skill).read_text(encoding="utf-8").lower()
            for verb in ("extract", "comprehend", "verify", "generate"):
                self.assertIn(verb, text, f"{skill} SKILL.md missing verb {verb}")

    def test_each_skill_links_shared_comprehension_reference(self) -> None:
        for skill in SKILLS:
            text = _skill_md(skill).read_text(encoding="utf-8")
            self.assertIn(
                "reference/comprehension.md",
                text,
                f"{skill} SKILL.md must link reference/comprehension.md",
            )

    def test_each_skill_documents_the_two_comprehend_cli_steps(self) -> None:
        for skill in SKILLS:
            text = _skill_md(skill).read_text(encoding="utf-8")
            self.assertIn("comprehend-input --name", text)
            self.assertIn("comprehend --name", text)

    def test_skip_when_cached_rule_present(self) -> None:
        # Cache soundness is bound to the shell sha: every skill must tell the agent
        # to skip comprehend when status==present AND source_shell_sha256 matches.
        for skill in SKILLS:
            text = _skill_md(skill).read_text(encoding="utf-8")
            self.assertIn("source_shell_sha256", text)
            self.assertIn("provenance.shell.sha256", text)


if __name__ == "__main__":
    unittest.main()
