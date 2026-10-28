#!/usr/bin/env python3
"""Image build step: replace Qwen Code's headless "questions" instruction.

Qwen Code (version pinned in the Dockerfile) tells every headless run: "This
is a non-interactive, single-turn run and no reply can be received ... Never
ask the user a question ... Make reasonable assumptions when safe and
complete the task". For docgen the premise is wrong: each run is one turn of
a conversation, and the final answer reaches the user, who replies in the next
message. Overriding the sentence from QWEN.md would leave two contradicting
rules in the system prompt, so the sentence itself is replaced. It is the
text Qwen also repeats as the "Interaction mode reminder".

The new text keeps Qwen's own policy of asking little ("Ask less", "make a
reasonable assumption and proceed"); it only allows the one question that is
needed when the run cannot continue safely without the user's decision.

The build fails unless the original text is found exactly once, so a Qwen
upgrade cannot drop the replacement silently.
"""
import subprocess
import sys
from pathlib import Path

OLD = ("This is a non-interactive, single-turn run and no reply can be received after your response. "
       "Never ask the user a question, even if the user explicitly requests one. "
       "Do not call '${ToolNames.ASK_USER_QUESTION}' or output a textual question. "
       "Make reasonable assumptions when safe and complete the task; if required information is "
       "unavailable, report the blocker as the final result.")
NEW = ("This run is one turn of a conversation: your final answer is shown to the user, who can reply in "
       "their next message. '${ToolNames.ASK_USER_QUESTION}' cannot be answered during the run, so do not "
       "call it. Ask only when you cannot proceed safely without the user's decision (for example a tool "
       "refused the change and says the user must decide, the target is ambiguous, or a destructive change "
       "has no consent): then make no change, explain why, and end your final answer with that one question "
       "or the options. For everything else, make reasonable assumptions, complete the task, and state the "
       "assumptions in your final answer.")


def main():
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else \
        Path(subprocess.run(["npm", "root", "-g"], capture_output=True, text=True, check=True).stdout.strip())
    pkg = root / "@qwen-code" / "qwen-code"
    hits = [(f, f.read_text(encoding="utf-8")) for f in sorted(pkg.rglob("*.js"))]
    hits = [(f, s) for f, s in hits if OLD in s]
    if len(hits) != 1 or hits[0][1].count(OLD) != 1:
        found = [(str(f), s.count(OLD)) for f, s in hits]
        sys.exit(f"patch_qwen_prompt: expected the headless 'questions' text exactly once in {pkg}, "
                 f"found {found or 'none'}; review the new Qwen Code prompt and update this script")
    f, s = hits[0]
    f.write_text(s.replace(OLD, NEW), encoding="utf-8")
    print(f"patch_qwen_prompt: replaced the headless 'questions' text in {f.relative_to(pkg)}")


if __name__ == "__main__":
    main()
