#!/usr/bin/env python3
"""
protect_source_data.py - PreToolUse hook enforcing CLAUDE.md Rule 9 and the
Masking control.

Two distinct protections, both scoped to the two read-only source trial
balances (data/trial_balance_2025.csv, data/trial_balance_2026.csv):

1. NO MODIFICATION (Rule 9). Blocks Edit/Write/MultiEdit/NotebookEdit calls
   targeting the source CSVs, so the model can't "fix" the source numbers
   even if a prompt - including one hidden in reviewed data - tries to talk
   it into doing so.

2. NO BULK RAW INGESTION (Masking). Blocks Read/Grep calls that would load
   the full source CSV into the model's context, and blocks Bash/PowerShell
   commands that mention a protected filename (cat, type, Get-Content,
   python -c "...read_csv(...)", etc.). This forces all number-crunching
   through skills/trial-balance-comparison/scripts/generate_visuals.py,
   which reads the data locally and only ever prints aggregate results
   (reconciliation OK/FAIL, flag counts) - never a row dump - so the raw
   31-account x 8-month ledger never has to enter the model's context to
   get the analysis done.

This is "no bulk raw ledger dump into the cloud model," not "zero dollar
figures ever discussed." Varia's job still requires citing specific flagged
account amounts back to the reviewer in chat (CLAUDE.md, Tone) and opening
the Excel workbook to verify formulas - those narrow, already-computed
figures are the deliverable, not the leak this hook is closing.

Known gaps, accepted for this simple demo:
- Repo-wide Grep/search calls that don't explicitly target a protected file
  by path are not blocked (ripgrep could still surface a matching line from
  the CSV in a broad search).
- Bash/PowerShell commands that read the file without ever naming it in the
  command string (e.g. reading a filename out of a variable set elsewhere)
  are not caught.
- File existence checks that name the file (e.g. `test -f data/trial_balance_2025.csv`)
  get blocked as a false positive - use a directory listing instead.

Exit code 2 blocks the tool call; stderr is surfaced to the model as the
reason.
"""

import json
import os
import sys

PROTECTED_FILENAMES = {"trial_balance_2025.csv", "trial_balance_2026.csv"}
MODIFYING_TOOLS = {"Edit", "Write", "MultiEdit", "NotebookEdit"}
READING_TOOLS = {"Read", "Grep"}
SHELL_TOOLS = {"Bash"}


def protected_path(path):
    norm = os.path.normpath(path).replace("\\", "/")
    parts = norm.split("/")
    if len(parts) < 2:
        return False
    filename, parent = parts[-1], parts[-2]
    return filename in PROTECTED_FILENAMES and parent == "data"


def deny(reason):
    sys.stderr.write(reason + "\n")
    return 2


def main():
    try:
        payload = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        return 0

    tool_name = payload.get("tool_name", "")
    tool_input = payload.get("tool_input", {})

    if tool_name in MODIFYING_TOOLS:
        for key in ("file_path", "notebook_path"):
            path = tool_input.get(key)
            if path and protected_path(path):
                return deny(
                    f"Blocked: '{path}' is a read-only source trial balance "
                    "(CLAUDE.md, Rule 9). Never modify files in data/ - write "
                    "results to outputs/ instead."
                )
        return 0

    if tool_name in READING_TOOLS:
        path = tool_input.get("file_path") or tool_input.get("path")
        if path and protected_path(path):
            return deny(
                f"Blocked: '{path}' may not be read directly into the model's "
                "context (CLAUDE.md, Canary/Masking control). Run "
                "skills/trial-balance-comparison/scripts/generate_visuals.py "
                "instead - it reads the source locally and reports aggregate "
                "results (reconciliation status, flag counts, flagged-account "
                "figures), never the full raw ledger."
            )
        return 0

    if tool_name in SHELL_TOOLS:
        command = tool_input.get("command", "")
        for filename in PROTECTED_FILENAMES:
            if filename in command:
                return deny(
                    f"Blocked: command references '{filename}' directly "
                    "(CLAUDE.md, Canary/Masking control). Do not cat/type/print "
                    "the source trial balances into the conversation - run "
                    "skills/trial-balance-comparison/scripts/generate_visuals.py "
                    "instead, or use a directory listing if you only need to "
                    "confirm the file exists."
                )
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
