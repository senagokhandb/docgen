"""DOCGEN helper library — tested building blocks the agent calls from its
own scripts (``import docgen.docx`` etc.).

Design rules shared by every module:
- In-place edits go through :class:`docgen.package.Package`: only the parts
  that were actually modified are re-serialized; every other ZIP entry is
  written back byte-identical (no directory entries are added).
- Targets are resolved explicitly. ``one()`` raises :class:`Ambiguous` when
  more than one candidate matches and :class:`NotFound` when none does —
  the agent must then report the candidates instead of guessing.
- Destructive operations raise :class:`NeedsConfirmation` unless the caller
  passes ``force=True`` (only when the user wrote ``--onay``).
"""


class DocgenError(Exception):
    """Base class for helper errors that must be reported to the user."""


class NotFound(DocgenError):
    pass


class Ambiguous(DocgenError):
    def __init__(self, what, candidates):
        self.candidates = list(candidates)
        lines = "\n".join(f"  - {c}" for c in self.candidates)
        super().__init__(f"{what}: {len(self.candidates)} candidates, be more specific:\n{lines}")


class NeedsConfirmation(DocgenError):
    """Raised for destructive/risky operations; retry with force=True only
    when the user explicitly confirmed (``--onay`` in the request)."""


class NotSupported(DocgenError):
    pass


def one(items, what="target"):
    """Return the single item or raise NotFound / Ambiguous."""
    items = list(items)
    if not items:
        raise NotFound(f"{what}: no match")
    if len(items) > 1:
        raise Ambiguous(what, items)
    return items[0]
