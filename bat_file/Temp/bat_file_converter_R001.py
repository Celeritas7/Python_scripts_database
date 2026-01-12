import os
import re

def classify_line(line):
    """Return the correct BAT command for a given line."""
    line = line.strip()

    # --- 1️⃣ Website link ---
    if re.match(r"^(https?://|www\.)", line, re.IGNORECASE):
        return f'start "" "{line}"'

    # --- 2️⃣ Folder path ---
    if (os.path.isdir(line)) or (not os.path.splitext(line)[1] and "\\" in line):
        return f'start "" "{line}"'

    # --- 3️⃣ File with extension (including .md) ---
    if re.search(r"\.[A-Za-z0-9]{2,5}$", line):
        abs_path = os.path.abspath(line)
        return f'start "" "{abs_path}"'

    # --- Default fallback ---
    return f'echo ⚠️ Unrecognized line: {line}'


# --- MAIN SCRIPT ---
# Use the folder where this Python file is located
folder = os.path.dirname(os.path.abspath(__file__))
print(f"📂 Scanning current folder: {folder}")

# Process all .txt files in the same folder
for filename in os.listdir(folder):
    if filename.lower().endswith(".txt"):
        txt_path = os.path.join(folder, filename)
        bat_name = os.path.splitext(filename)[0] + ".bat"
        bat_path = os.path.join(folder, bat_name)

        # Read .txt lines
        with open(txt_path, "r", encoding="utf-8") as txt_file:
            lines = [line.strip() for line in txt_file if line.strip()]

        # Write .bat file
        with open(bat_path, "w", encoding="utf-8") as bat_file:
            bat_file.write("@echo off\n")
            bat_file.write(f"echo Launching items from {filename}...\n\n")

            for line in lines:
                command = classify_line(line)
                bat_file.write(command + "\n")

            bat_file.write("\necho Done!\npause\n")

        print(f"✅ Created: {bat_name}")

print("\nAll .bat files generated successfully!")
