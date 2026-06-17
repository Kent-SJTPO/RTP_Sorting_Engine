from pathlib import Path
import shutil
import sys


SOURCE_REPO = Path(r"C:\Users\Kschellinger\SJTPO_Git\RTP_Sorting_Engine")
DESTINATION = Path(
    r"H:\Planning\RTPs\RTP 2050 (2025)\Financial Plan\RTP-Sorting-Engine"
)

EXCLUDE_DIRS = {
    ".git",
    ".venv",
    ".conda",
    "__pycache__",
    ".pytest_cache",
}

EXCLUDE_SUFFIXES = {
    ".pyc",
    ".pyo",
    ".log",
}


def ignore_items(directory, names):
    ignored = []

    for name in names:
        item = Path(directory) / name

        if item.is_dir() and name in EXCLUDE_DIRS:
            ignored.append(name)

        if item.is_file() and item.suffix.lower() in EXCLUDE_SUFFIXES:
            ignored.append(name)

    return ignored


def main():
    print("Mirror RTP Sorting Engine only")
    print(f"Source:      {SOURCE_REPO}")
    print(f"Destination: {DESTINATION}")

    if SOURCE_REPO.name != "RTP_Sorting_Engine":
        print("ERROR: Source repo is not RTP_Sorting_Engine.")
        return 1

    if not SOURCE_REPO.exists():
        print(f"ERROR: Source repo does not exist: {SOURCE_REPO}")
        return 1

    if not (SOURCE_REPO / ".git").exists():
        print(f"ERROR: Source does not appear to be a Git repo: {SOURCE_REPO}")
        return 1

    if not DESTINATION.parent.exists():
        print(f"ERROR: Destination parent does not exist: {DESTINATION.parent}")
        return 1

    if DESTINATION.exists():
        shutil.rmtree(DESTINATION)

    shutil.copytree(SOURCE_REPO, DESTINATION, ignore=ignore_items)

    print("Mirror completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())