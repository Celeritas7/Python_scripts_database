import sys
import os
from pathlib import Path
import subprocess

def convert_notebook_to_word(input_path: Path, output_folder: Path):
    if input_path.suffix.lower() != '.ipynb' or '.ipynb_checkpoints' in str(input_path):
        return False

    print(f"  Converting: {input_path.name}")
    
    try:
        # Added --debug to see exactly where it fails if it crashes
        result = subprocess.run([
            "jupyter", "nbconvert", 
            "--to", "docx", 
            "--output-dir", str(output_folder),
            str(input_path)
        ], capture_output=True, text=True, check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"  [ERROR] Failed {input_path.name}")
        # This will print the actual error from Jupyter/Pandoc
        if "pandoc" in e.stderr.lower():
            print("          >>> REASON: Pandoc not found. Please install from pandoc.org and RESTART.")
        else:
            print(f"          >>> DETAILS: {e.stderr.strip()}")
        return False

def main():
    print("=" * 50)
    print("      IPYNB to WORD Batch Converter")
    print("=" * 50)

    current_folder = Path(__file__).parent.absolute()
    files_to_convert = [f for f in current_folder.glob("*.ipynb") if ".ipynb_checkpoints" not in str(f)]

    if not files_to_convert:
        print(f"[INFO] No notebooks found in: {current_folder}")
        input("\nPress Enter to exit...")
        return

    output_folder = current_folder / "Word_Outputs"
    output_folder.mkdir(exist_ok=True)
    
    success_count = 0
    for file in files_to_convert:
        if convert_notebook_to_word(file, output_folder):
            print(f"  [SUCCESS] Created {file.stem}.docx")
            success_count += 1

    print("\n" + "=" * 50)
    print(f"FINISHED: {success_count}/{len(files_to_convert)} files converted.")
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()