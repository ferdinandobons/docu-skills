# SPDX-License-Identifier: MIT
"""Fail-closed guard for the frozen QA check-id registry (CONVENTIONS.md section 9).

Check ids are a PERSISTED contract: they key ``generation_report.json`` digests,
the cross-run regression multiset, and the ``learn`` distillation. The gate
orchestrator and the test suites hardcode them (there is no runtime registry
lookup), so a typo in a ``Finding(check_id, ...)`` call would drift silently.

This module AST-scans ``qa/checks_deterministic.py`` for every ``Finding``
construction and asserts:

  1. every check id is a STATIC string literal (a dynamic id could not be
     audited or asserted against);
  2. every emitted id is a member of the frozen ``CHECK_REGISTRY`` in
     ``qa/model.py`` (the typo guard);
  3. every registered id is actually emitted by the module (the freshness
     guard - a retired check must be retired from the registry too, with its
     persisted-history implications weighed explicitly).
"""

from __future__ import annotations

import ast
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.qa.model import CHECK_REGISTRY

ROOT = Path(__file__).resolve().parents[1]
CHECKS_MODULE = ROOT / "scripts" / "brandkit" / "qa" / "checks_deterministic.py"


def _emitted_check_ids() -> tuple[set[str], list[int]]:
    """All static check ids stamped on a ``Finding`` in the deterministic module,
    plus the line numbers of any DYNAMIC (non-literal) check argument."""
    tree = ast.parse(CHECKS_MODULE.read_text(encoding="utf-8"))
    ids: set[str] = set()
    dynamic: list[int] = []

    class Visitor(ast.NodeVisitor):
        def visit_Call(self, node: ast.Call) -> None:
            name = getattr(node.func, "id", None) or getattr(node.func, "attr", None)
            if name == "Finding":
                check = None
                if node.args:
                    check = node.args[0]
                for kw in node.keywords:
                    if kw.arg == "check":
                        check = kw.value
                if isinstance(check, ast.Constant) and isinstance(check.value, str):
                    ids.add(check.value)
                else:
                    dynamic.append(node.lineno)
            self.generic_visit(node)

    Visitor().visit(tree)
    return ids, dynamic


class CheckRegistryTest(unittest.TestCase):
    def test_every_check_id_is_a_static_literal(self) -> None:
        _, dynamic = _emitted_check_ids()
        self.assertEqual(
            [],
            dynamic,
            f"Finding() calls at lines {dynamic} of checks_deterministic.py carry a "
            f"non-literal check id; check ids must be static strings so the registry "
            f"guard (and every hardcoded assertion) can audit them",
        )

    def test_every_emitted_id_is_registered(self) -> None:
        emitted, _ = _emitted_check_ids()
        unregistered = sorted(emitted - CHECK_REGISTRY)
        self.assertEqual(
            [],
            unregistered,
            f"checks_deterministic.py emits unregistered check id(s) {unregistered}; "
            f"a new check must be added to CHECK_REGISTRY in qa/model.py (and follow "
            f"the naming convention in CONVENTIONS.md section 9) - a typo in an "
            f"existing id must be fixed at the Finding call site",
        )

    def test_every_registered_id_is_emitted(self) -> None:
        emitted, _ = _emitted_check_ids()
        stale = sorted(CHECK_REGISTRY - emitted)
        self.assertEqual(
            [],
            stale,
            f"CHECK_REGISTRY lists check id(s) {stale} that checks_deterministic.py "
            f"never emits; retire the registry entry together with the check (check "
            f"ids key persisted generation history, so weigh the removal explicitly)",
        )


if __name__ == "__main__":
    unittest.main()
