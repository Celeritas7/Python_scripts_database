import os
import shutil
from PIL import Image

# Folder where this Python file is located
script_dir = os.path.dirname(os.path.abspath(__file__))
work_dir = os.getcwd()

# If called from a different folder (e.g. via tools), use that folder
# If run directly from its own folder, use script's folder
if work_dir != script_dir:
    script_dir = work_dir

# Folder to store original files
old_folder = os.path.join(script_dir, "Old")

# Create "Old" folder if not present
os.makedirs(old_folder, exist_ok=True)

# Settings
SUPPORTED = (".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff", ".tif")

# Ask user for quality
print("JPEG Quality (lower = smaller file, more compression):")
print("  90 = barely noticeable")
print("  75 = good balance (default)")
print("  50 = noticeable on zoom")
print("  30 = aggressive, visible artifacts")
print("  10 = maximum compression")
raw = input("\nEnter quality 1-95 (default 75): ").strip()
QUALITY = int(raw) if raw else 75
print()

total_before = 0
total_after = 0
count = 0

for filename in os.listdir(script_dir):
    if not filename.lower().endswith(SUPPORTED):
        continue

    input_path = os.path.join(script_dir, filename)
    backup_path = os.path.join(old_folder, filename)

    # Skip files already inside "Old"
    if os.path.commonpath([input_path, old_folder]) == old_folder:
        continue

    original_size = os.path.getsize(input_path)

    # --- Step 1: Move original file to "Old" ---
    shutil.move(input_path, backup_path)

    # --- Step 2: Open original image from "Old" ---
    image = Image.open(backup_path)

    # Convert to RGB for JPEG compatibility
    if image.mode in ("RGBA", "P", "LA"):
        image = image.convert("RGB")

    # --- Step 3: Save as compressed JPG (same dimensions) ---
    output_name = os.path.splitext(filename)[0] + ".jpg"
    output_path = os.path.join(script_dir, output_name)
    image.save(output_path, "JPEG", quality=QUALITY, optimize=True)

    new_size = os.path.getsize(output_path)
    reduction = ((original_size - new_size) / original_size * 100)
    total_before += original_size
    total_after += new_size
    count += 1

    print(f"  {filename} → {output_name}  ({original_size//1024}KB → {new_size//1024}KB, {reduction:.0f}% smaller)")

# Summary
print()
print(f"✅ {count} images processed (quality: {QUALITY})")
print(f"   Total: {total_before//1024//1024}MB → {total_after//1024//1024}MB "
      f"({((total_before-total_after)/total_before*100):.0f}% smaller)")
print(f"   Originals safe in: {old_folder}")
