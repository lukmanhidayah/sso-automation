import os
from typing import Set


def _load_selected_ids(file_path: str) -> set[str]:
    ids: set[str] = set()
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if not s or s.startswith("#"):
                    continue
                ids.add(s)
    except FileNotFoundError:
        # If the file is missing, fall back to empty set
        pass
    return ids


# Default path: project_root/data/selected_no_peserta.txt
_DEFAULT_PATH = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "data", "selected_no_peserta.txt")
)

# Public export
selected_no_peserta: set[str] = _load_selected_ids(_DEFAULT_PATH)

