"""Resize JPG/JPEG images in a folder and all its subfolders.

Each image's original is moved into an 'Old' folder beside that image.
'Old' folders (including their descendants) are never scanned.

Examples:
    python Image_shrink_R004_recursive.py
    python Image_shrink_R004_recursive.py --folder "C:\\Pictures" --width 800

Requires: pip install Pillow
"""

import argparse
import os
import shutil
import tempfile
from pathlib import Path

from PIL import Image, ImageOps


def process_images(root: Path, width: int) -> tuple[int, int, int]:
    resized = skipped = failed = 0

    for current_dir, folders, files in os.walk(root):
        # Do not visit original-image backup directories or any of their children.
        folders[:] = [f for f in folders if f.lower() != "old" and not (Path(current_dir) / f).is_symlink()]
        current = Path(current_dir)

        for filename in files:
            if Path(filename).suffix.lower() not in {".jpg", ".jpeg"}:
                continue

            source = current / filename
            backup_dir = current / "Old"
            backup = backup_dir / filename
            temp_path = None

            # Prevent backup replacement and repeated lossy resizing on later runs.
            if backup.exists():
                print(f"SKIPPED (original already backed up): {source}")
                skipped += 1
                continue

            try:
                # Prepare the result first, without touching the original.
                with Image.open(source) as original:
                    image = ImageOps.exif_transpose(original)
                    new_height = max(1, round(image.height * width / image.width))
                    resized_image = image.resize((width, new_height), Image.Resampling.LANCZOS)
                    if resized_image.mode not in ("RGB", "L"):
                        resized_image = resized_image.convert("RGB")

                    with tempfile.NamedTemporaryFile(
                        suffix=".jpg", prefix=".image_shrink_", dir=current, delete=False
                    ) as temp:
                        temp_path = Path(temp.name)
                    resized_image.save(temp_path, format="JPEG", quality=90)

                backup_dir.mkdir(exist_ok=True)
                # Keep the original; never silently replace an existing backup.
                if backup.exists():
                    print(f"SKIPPED (backup appeared): {source}")
                    skipped += 1
                    continue
                shutil.move(str(source), str(backup))
                try:
                    os.replace(temp_path, source)
                    temp_path = None
                except Exception:
                    # Try to restore original if placing resized output fails.
                    shutil.move(str(backup), str(source))
                    raise

                print(f"RESIZED: {source.relative_to(root)}  ->  {width}px wide")
                resized += 1
            except Exception as exc:
                print(f"FAILED: {source} ({exc})")
                failed += 1
            finally:
                if temp_path is not None:
                    temp_path.unlink(missing_ok=True)

    return resized, skipped, failed


def main() -> None:
    parser = argparse.ArgumentParser(description="Shrink JPG/JPEG images recursively; preserve originals in Old folders.")
    parser.add_argument("--folder", type=Path, help="Folder to scan (default: working folder, as in R003)")
    parser.add_argument("--width", type=int, default=800, help="Output width in pixels (default: 800)")
    args = parser.parse_args()
    if args.width < 1:
        parser.error("--width must be a positive number")

    # Match R003's behavior: the working folder when launched elsewhere;
    # otherwise the script's own folder.
    script_dir = Path(__file__).resolve().parent
    working_dir = Path.cwd().resolve()
    root = (args.folder or (working_dir if working_dir != script_dir else script_dir)).resolve()
    if not root.is_dir():
        parser.error(f"Folder not found: {root}")

    print(f"Scanning recursively: {root}")
    resized, skipped, failed = process_images(root, args.width)
    print(f"\nFinished: {resized} resized, {skipped} skipped, {failed} failed.")


if __name__ == "__main__":
    main()
