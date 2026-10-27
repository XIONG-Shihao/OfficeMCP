"""The operation catalog: every action, declared as data, with the function that carries it out.

The catalog is the single source for the MCP schemas, what `office_actions` lists, the settings
page's catalog (`contracts/catalog.json`) and the agent guide (`agent/OfficeMCP.md`).
"""

from collections import Counter
from dataclasses import dataclass, field
from typing import Any, Callable

FORMAT_NAMES = {"docx": "DOCX (Word)", "xlsx": "XLSX (Excel)", "pptx": "PPTX (PowerPoint)", "pdf": "PDF"}


@dataclass
class Edit:
    """What one batch is allowed to remove, and what its actions report, as the actions run."""

    config: Any
    allowed: Counter = field(default_factory=Counter)  # fidelity keys an action declared it removes
    warnings: list = field(default_factory=list)
    recalculation_needed: bool = False

    def allow(self, key: str, count: int = 1) -> None:
        self.allowed[key] += count


@dataclass
class Action:
    format: str
    name: str
    family: str
    summary: str
    params: dict
    required: list
    run: Callable[[Any, dict, Edit], dict]
    tier: str = "primitive"  # or "recipe": several library calls behind one action
    creates: bool = False  # PDF compose actions apply only while creating a PDF

    def schema(self) -> dict:
        return {
            "type": "object",
            "title": self.name,
            "description": self.summary,
            "properties": {"action": {"const": self.name}, **self.params},
            "required": ["action", *self.required],
            "additionalProperties": False,
        }

    def entry(self) -> dict:
        return {
            "format": self.format,
            "name": self.name,
            "family": self.family,
            "tier": self.tier,
            "summary": self.summary,
            "input": self.schema(),
            **({"createOnly": True} if self.creates else {}),
        }


ACTIONS: dict[tuple[str, str], Action] = {}


def action(format: str, name: str, family: str, summary: str, params: dict, required=(), tier="primitive", creates=False):
    """Declare an action; the decorated function receives (document, arguments, edit) and returns a receipt."""

    def register(run):
        key = (format, name)
        assert key not in ACTIONS, key
        ACTIONS[key] = Action(format, name, family, summary, params, list(required), run, tier, creates)
        return run

    return register


def actions(format: str | None = None, family: str | None = None) -> list[Action]:
    from . import formats  # noqa: F401  (registers every format's actions)

    return [a for a in ACTIONS.values() if (format is None or a.format == format) and (family is None or a.family == family)]


def lookup(format: str, name: str) -> Action | None:
    from . import formats  # noqa: F401

    return ACTIONS.get((format, name))


# Shared parameter schemas, so the same idea is spelled the same way everywhere.
def text(description: str, max_length: int = 100_000) -> dict:
    return {"type": "string", "maxLength": max_length, "description": description}


def address(description: str) -> dict:
    return {"type": "string", "minLength": 1, "maxLength": 500, "description": description}


EXPECTED_TEXT = text("The text the target holds now; the action refuses if it differs (expected_mismatch).")
POSITION = {"type": "string", "enum": ["before", "after"], "default": "after", "description": "Where to insert relative to the target."}
CELL_VALUE = {
    "description": "A number, text, true/false, or null to clear. Text starting with '=' is a formula.",
    "type": ["number", "string", "boolean", "null"],
}
