import sys
import os
from pathlib import Path
import subprocess

def convert_notebook_to_word(input_path: Path, output_folder: Path):
    """Convert a single IPYNB file to Word (.docx) using nbconvert."""
    # Skip checkpoint files and non-notebooks
    if input_path.suffix.lower() != '.ipynb' or '.ipynb_checkpoints' in str(input_path):
        return False

    print(f"  Converting: {input_path.name}")
    
    try:
        # Using nbconvert to export to script-friendly DOCX via Pandoc
        subprocess.run([
            "jupyter", "nbconvert", 
            "--to", "docx", 
            "--output-dir", str(output_folder),
            str(input_path)
        ], check=True, capture_output=True)
        return True
    except Exception as e:
        print(f"  [ERROR] Failed {input_path.name}: {e}")
        # Hint for common error
        if "Pandoc" in str(e) or "pandoc" in str(e):
            print("          Tip: Please install Pandoc from https://pandoc.org/installing.html")
        return False

def main():
    print("=" * 50)
    print("      Auto-Folder IPYNB to WORD Converter")
    print("=" * 50)

    # Automatically set the path to the folder where this script is located
    current_folder = Path(__file__).parent.absolute()
    
    # Find all .ipynb files in the current folder
    files_to_convert = list(current_folder.glob("*.ipynb"))
    
    # Filter out hidden checkpoint files
    files_to_convert = [f for f in files_to_convert if ".ipynb_checkpoints" not in str(f)]

    if not files_to_convert:
        print(f"[INFO] No .ipynb files found in:\n{current_folder}")
        input("\nPress Enter to exit...")
        return

    # Setup Output Folder
    output_folder = current_folder / "Word_Outputs"
    output_folder.mkdir(exist_ok=True)
    
    print(f"Target Folder: {current_folder}")
    print(f"Found {len(files_to_convert)} notebook(s).")
    print(f"Saving to: {output_folder}\n")

    # Process Loop
    success_count = 0
    for file in files_to_convert:
        if convert_notebook_to_word(file, output_folder):
            print(f"  [SUCCESS] Created Word doc for {file.name}")
            success_count += 1
        else:
            success_count += 0

    print("\n" + "=" * 50)
    print(f"FINISHED: {success_count}/{len(files_to_convert)} files converted.")
    print(f"Check the 'Word_Outputs' folder.")
    print("=" * 50)
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()