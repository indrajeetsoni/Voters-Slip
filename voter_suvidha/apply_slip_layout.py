# -*- coding: utf-8 -*-
"""
apply_slip_layout.py — wire slip_layout.py into server.py.

server.py defines its own render_slip_html() and get_full_page_css().
This script adds one import line AFTER those definitions, so the versions
from slip_layout.py take over. Nothing else in server.py is touched.

Run once, from the voter_suvidha folder:

    python apply_slip_layout.py

It writes a backup (server.py.bak) first and refuses to run twice.
"""

import os
import re
import shutil
import sys

MARKER = "from slip_layout import render_slip_html, get_full_page_css"
LINE = MARKER + "  # new slip design (see slip_layout.py)\n"

# these appear immediately after get_full_page_css() in server.py; the first
# one found is used as the insertion point
ANCHORS = (
    "def find_browser(",
    "def save_base64_image_to_temp_file(",
    "class ThreadedHTTPServer",
    "class VoterSuvidhaHandler",
)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    server = os.path.join(here, "server.py")

    if not os.path.exists(server):
        sys.exit(f"server.py not found in {here}\n"
                 "Run this script from inside the voter_suvidha folder.")
    if not os.path.exists(os.path.join(here, "slip_layout.py")):
        sys.exit("slip_layout.py is missing — copy it into this folder first.")

    with open(server, encoding="utf-8") as fh:
        text = fh.read()

    if MARKER in text:
        print("Already patched — nothing to do.")
        return

    for name in ("render_slip_html", "get_full_page_css"):
        if f"def {name}(" not in text:
            sys.exit(f"Could not find def {name}() in server.py. "
                     "This server.py is not the expected version; tell me and "
                     "I will adjust the patch.")

    lines = text.splitlines(keepends=True)
    idx = None
    for i, line in enumerate(lines):
        if any(line.startswith(a) for a in ANCHORS):
            idx = i
            break

    if idx is None:
        sys.exit("Could not find a safe place to insert the import.\n"
                 "Add this line to server.py by hand, anywhere BELOW the\n"
                 "definition of get_full_page_css():\n\n    " + MARKER)

    backup = server + ".bak"
    if not os.path.exists(backup):
        shutil.copy2(server, backup)
        print(f"backup written: {os.path.basename(backup)}")

    lines.insert(idx, "\n" + LINE + "\n")
    with open(server, "w", encoding="utf-8") as fh:
        fh.writelines(lines)

    print(f"patched server.py — import added before line {idx + 1} "
          f"({lines[idx + 2].strip()[:40]}...)")
    print("Restart the server for the new design to take effect.")


if __name__ == "__main__":
    main()
