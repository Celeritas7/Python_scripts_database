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

# Resize settings
new_width = 800  # target width

# Loop through all JPG images in the script folder
for filename in os.listdir(script_dir):

    if not filename.lower().endswith(".jpg"):
        continue

    input_path = os.path.join(script_dir, filename)
    backup_path = os.path.join(old_folder, filename)

    # Skip files already inside "Old"
    if os.path.commonpath([input_path, old_folder]) == old_folder:
        continue

    # --- Step 1: Move original file to "Old" ---
    shutil.move(input_path, backup_path)

    # --- Step 2: Open original image from "Old" ---
    image = Image.open(backup_path)

    # Resize while keeping aspect ratio
    aspect_ratio = new_width / image.width
    new_height = int(image.height * aspect_ratio)
    resized_image = image.resize(
        (new_width, new_height),
        Image.Resampling.LANCZOS
    )

    # --- Step 3: Save resized image back to script folder ---
    resized_image.save(input_path)

    print(f"Processed & replaced: {filename}")

print("✅ All images updated successfully!")
