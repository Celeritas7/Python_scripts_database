import os
import shutil
import subprocess

# Folder where this Python file is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Folder to store original videos
old_folder = os.path.join(script_dir, "Old")
os.makedirs(old_folder, exist_ok=True)

# Compression settings
CRF_VALUE = "28"          # 26–28 = best balance
PRESET = "medium"         # slow = smaller, fast = quicker

for filename in os.listdir(script_dir):

    if not filename.lower().endswith(".mp4"):
        continue

    input_path = os.path.join(script_dir, filename)
    backup_path = os.path.join(old_folder, filename)

    # Skip files already inside Old
    if os.path.commonpath([input_path, old_folder]) == old_folder:
        continue

    # Move original video to Old
    shutil.move(input_path, backup_path)

    # FFmpeg compression command (NO resizing)
    cmd = [
        "ffmpeg",
        "-i", backup_path,
        "-c:v", "libx265",
        "-crf", CRF_VALUE,
        "-preset", PRESET,
        "-c:a", "copy",          # keep original audio
        input_path
    ]

    subprocess.run(cmd, check=True)

    print(f"Compressed & replaced: {filename}")

print("✅ All videos compressed successfully!")
