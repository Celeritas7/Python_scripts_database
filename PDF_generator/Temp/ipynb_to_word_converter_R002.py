import sys
import os
from pathlib import Path
import subprocess

# This library helps bridge the gap when nbconvert doesn't see the docx exporter
try:
    import pypandoc
except ImportError:
    print("Error: pypandoc not found. Please run: pip install pypandoc")
    input("Press Enter to exit...")
    sys.exit(1)

def convert_notebook_to_word(input_path: Path, output_folder: Path):
    """Convert IPYNB to DOCX by using Pandoc directly."""
    if input_path.suffix.lower() != '.ipynb' or '.ipynb_checkpoints' in str(input_path):
        return False

    output_file = output_folder / f"{input_path.stem}.docx"
    print(f"  Converting: {input_path.name}")
    
    try:
        # Use pypandoc to convert directly from file to file
        # This avoids the 'Unknown exporter docx' error in nbconvert
        pypandoc.convert_file(str(input_path), 'docx', outputfile=str(output_file))
        return True
    except Exception as e:
        print(f"  [ERROR] Failed {input_path.name}: {e}")
        return False

def main():
    print("=" * 50)
    print("      IPYNB to WORD Batch Converter (Hybrid)")
    print("=" * 50)

    # Automatically target the folder where this script lives
    current_folder = Path(__file__).parent.absolute()
    files_to_convert = [f for f in current_folder.glob("*.ipynb") if ".ipynb_checkpoints" not in str(f)]

    if not files_to_convert:
        print(f"[INFO] No notebooks found in: {current_folder}")
        input("\nPress Enter to exit...")
        return

    output_folder = current_folder / "Word_Outputs"
    output_folder.mkdir(exist_ok=True)
    
    print(f"Targeting: {current_folder}")
    print(f"Saving to: {output_folder}\n")

    success_count = 0
    for file in files_to_convert:
        if convert_notebook_to_word(file, output_folder):
            print(f"  [SUCCESS] Created {file.stem}.docx")
            success_count += 1

    print("\n" + "=" * 50)
    print(f"FINISHED: {success_count}/{len(files_to_convert)} notebooks converted.")
    print("=" * 50)
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()