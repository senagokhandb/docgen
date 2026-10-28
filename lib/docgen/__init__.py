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
  passes ``force=True`` (only when the user wrote ``--onay``). In the agent
  sandbox the library checks that consent itself: see :func:`require_consent`.
"""
import functools

CONSENT_VAR = "DOCGEN_ONAY"


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


# ---- consent (--onay) --------------------------------------------------------
@functools.lru_cache(maxsize=1)
def _main_process_env():
    """Environment of the container's main process (PID 1); {} without /proc."""
    try:
        with open("/proc/1/environ", "rb") as fh:
            raw = fh.read()
    except OSError:
        return {}
    return dict(x.decode(errors="replace").split("=", 1) for x in raw.split(b"\0") if b"=" in x)


def consented():
    """Whether force=True may override a NeedsConfirmation guard.

    In the agent sandbox the only consent is the user's ``--onay``:
    entrypoint.sh reads it from the request before the agent starts and pins
    DOCGEN_ONAY=1|0 into the environment of the container's main process.
    The agent can set variables in its own processes but cannot change that
    one, so neither force=True nor DOCGEN_ONAY=1 written by the agent counts.
    Without DOCGEN_ONAY on the main process (tests, host scripts) the caller's
    force=True stands."""
    return _main_process_env().get(CONSENT_VAR, "1") == "1"


def require_consent(what):
    """Call where force=True is about to override a guard; raises
    NeedsConfirmation when the user's request had no --onay."""
    if not consented():
        raise NeedsConfirmation(f"{what}: force=True / --onay is accepted only when the user's request contains "
                                f"--onay, and this one does not. Write no file; explain the risk and ask the user.")
