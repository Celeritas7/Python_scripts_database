"""
PDF Password Remover
====================
Removes the password from encrypted PDFs *you have the password for*, then
saves the unlocked file back into the same folder under its original name.
The original (still-encrypted) file is moved into an "Old" subfolder.

Drop this script into a folder of PDFs and run it. It scans its own folder,
unlocks every PDF it can, and leaves you with unlocked files in place and the
originals safely backed up in "Old".

This does NOT crack or guess unknown passwords. It only decrypts PDFs whose
password you already know (or that use an empty owner password). A file whose
password can't be found is left exactly where it is — never moved.

How the password is supplied (checked in this order, per file):
    1. The PASSWORDS list below (edit it), tried in order.
    2. A "passwords.txt" file next to this script, one password per line.
    3. The empty password (covers owner-only encryption).
    4. An interactive prompt, only for files still locked after the above.

Requirements:
    pip install pypdf
"""

import getpass
import io
import shutil
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import PdfReadError
except ImportError:
    print("ERROR: pypdf not installed. Run: pip install pypdf")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Config — edit if you like, or leave empty and use passwords.txt / prompts.
# ---------------------------------------------------------------------------
PASSWORDS = [
    # "mypassword",
    # "another-if-files-differ",
]
OLD_FOLDER_NAME = "Old"        # where originals are backed up
PROMPT_FOR_UNSOLVED = True     # ask interactively when no known password works
# ---------------------------------------------------------------------------


def resolve_work_dir() -> Path:
    """Same logic as the image script: prefer cwd if launched elsewhere."""
    script_dir = Path(__file__).resolve().parent
    work_dir = Path.cwd()
    return work_dir if work_dir != script_dir else script_dir


def load_passwords(base_dir: Path) -> list:
    """PASSWORDS list + passwords.txt + empty password (deduped, order kept)."""
    candidates = list(PASSWORDS)
    pwfile = base_dir / "passwords.txt"
    if pwfile.exists():
        for line in pwfile.read_text(encoding="utf-8", errors="ignore").splitlines():
            if line.rstrip("\n"):
                candidates.append(line.rstrip("\n"))
    candidates.append("")  # owner-only encryption has an empty user password
    seen, ordered = set(), []
    for p in candidates:
        if p not in seen:
            seen.add(p)
            ordered.append(p)
    return ordered


def try_unlock(reader, passwords):
    """Attempt each password until one is accepted."""
    for pw in passwords:
        try:
            if reader.decrypt(pw):
                return True
        except (PdfReadError, NotImplementedError):
            continue
    return False


def safe_backup_path(old_folder: Path, filename: str) -> Path:
    """Destination inside Old/, avoiding clobbering an existing backup."""
    dest = old_folder / filename
    if not dest.exists():
        return dest
    stem, suffix = Path(filename).stem, Path(filename).suffix
    n = 1
    while (old_folder / f"{stem} ({n}){suffix}").exists():
        n += 1
    return old_folder / f"{stem} ({n}){suffix}"


def process_pdf(path: Path, old_folder: Path, passwords) -> str:
    """Unlock one PDF, then move its original into Old/. Returns a status line."""
    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        return f"ERROR  {path.name}  ({type(exc).__name__}: {exc})"

    if not reader.is_encrypted:
        return f"skip   {path.name}  (not encrypted, left in place)"

    ok = try_unlock(reader, passwords)
    if not ok and PROMPT_FOR_UNSOLVED and sys.stdin.isatty():
        pw = getpass.getpass(f"    password for '{path.name}' (blank to skip): ")
        if pw:
            ok = try_unlock(reader, [pw])

    if not ok:
        return f"FAILED {path.name}  (no matching password, left in place)"

    # Build the unlocked PDF in memory first, so the original file handle is
    # free before we move it.
    try:
        writer = PdfWriter()
        writer.append(reader)
        buffer = io.BytesIO()
        writer.write(buffer)
        del reader  # release the source file
    except Exception as exc:
        return f"ERROR  {path.name}  (decrypt/build failed: {exc})"

    # Step 1: move the original into Old/
    backup_path = safe_backup_path(old_folder, path.name)
    try:
        shutil.move(str(path), str(backup_path))
    except Exception as exc:
        return f"ERROR  {path.name}  (backup move failed: {exc})"

    # Step 2: write the unlocked copy back to the original location
    try:
        with open(str(path), "wb") as f:
            f.write(buffer.getvalue())
    except Exception as exc:
        # Roll back so nothing is lost if the write fails
        shutil.move(str(backup_path), str(path))
        return f"ERROR  {path.name}  (write failed, original restored: {exc})"

    return f"OK     {path.name}  (unlocked; original -> {OLD_FOLDER_NAME}/{backup_path.name})"


def main() -> None:
    base_dir = resolve_work_dir()
    old_folder = base_dir / OLD_FOLDER_NAME
    old_folder.mkdir(exist_ok=True)

    pdfs = sorted(
        p for p in base_dir.glob("*.pdf")
        if p.is_file() and p.parent == base_dir
    )
    if not pdfs:
        print(f"No PDF files found in: {base_dir}")
        return

    passwords = load_passwords(base_dir)
    print(f"Found {len(pdfs)} PDF(s) in: {base_dir}")
    print(f"Originals will be backed up to: {old_folder}\n")
    print("-" * 50)

    results = [process_pdf(p, old_folder, passwords) for p in pdfs]

    for r in results:
        print(f"  {r}")
    print("-" * 50)
    done = sum(r.startswith("OK") for r in results)
    print(f"Done: {done}/{len(pdfs)} unlocked.")


if __name__ == "__main__":
    main()
