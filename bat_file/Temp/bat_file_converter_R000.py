import os
import re

def classify_line(line):
    """Return the type of path and generate the corresponding BAT command."""
    line = line.strip()

    # --- 1️⃣ Website link ---
    if re.match(r"^(https?://|www\.)", line, re.IGNORECASE):
        return f'start "" "{line}"'

    # --- 2️⃣ Folder path (no extension, exists or looks like a dir) ---
    if (os.path.isdir(line)) or (not os.path.splitext(line)[1] and "\\" in line):
        return f'start "" "{line}"'

    # --- 3️⃣ File with extension ---
    if re.search(r"\.[A-Za-z0-9]{2,5}$", line):
        return f'start "" "{line}"'

    # --- Default fallback ---
    return f'echo ⚠️ Unrecognized line: {line}'

# --- MAIN SCRIPT ---
folder = os.getcwd()

for filename in os.listdir(folder):
    if filename.lower().endswith(".txt"):
        txt_path = os.path.join(folder, filename)
        bat_name = os.path.splitext(filename)[0] + ".bat"
        bat_path = os.path.join(folder, bat_name)

        with open(txt_path, "r", encoding="utf-8") as txt_file:
            lines = [line.strip() for line in txt_file if line.strip()]

        with open(bat_path, "w", encoding="utf-8") as bat_file:
            bat_file.write("@echo off\n")
            bat_file.write(f"echo Launching items from {filename}...\n\n")

            for line in lines:
                command = classify_line(line)
                bat_file.write(command + "\n")

            bat_file.write("\necho Done!\npause\n")

        print(f"✅ Created: {bat_name}")
