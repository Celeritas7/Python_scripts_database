import sys
import os
from pathlib import Path
import subprocess

def convert_notebook_to_pdf(input_path: Path, output_folder: Path):
    """Convert a single IPYNB file to PDF using nbconvert."""
    if input_path.suffix.lower() != '.ipynb':
        return False

    output_path = output_folder / input_path.with_suffix('.pdf').name
    print(f"  Converting: {input_path.name}")
    
    try:
        # Uses webpdf to preserve layouts and charts
        subprocess.run([
            "jupyter", "nbconvert", 
            "--to", "webpdf", 
            "--output-dir", str(output_folder),
            "--allow-chromium-download",
            str(input_path)
        ], check=True, capture_output=True)
        return True
    except Exception as e:
        print(f"  [ERROR] Failed {input_path.name}: {e}")
        return False

def main():
    print("=" * 50)
    print("      Jupyter Notebook to PDF Batch Converter")
    print("=" * 50)

    # 1. Get Path (Drag & Drop or Manual Input)
    if len(sys.argv) > 1:
        input_path = Path(sys.argv[1].strip('"'))
    else:
        user_input = input("Enter path to .ipynb file or folder: ").strip().strip('"')
        input_path = Path(user_input)

    if not input_path.exists():
        print(f"[ERROR] Path not found: {input_path}")
        input("\nPress Enter to exit...")
        return

    # 2. Determine target files
    if input_path.is_file():
        files_to_convert = [input_path]
        parent_dir = input_path.parent
    else:
        files_to_convert = list(input_path.glob("*.ipynb"))
        parent_dir = input_path

    if not files_to_convert:
        print("[INFO] No .ipynb files found.")
        input("\nPress Enter to exit...")
        return

    # 3. Setup Output Folder
    output_folder = parent_dir / "PDF_Outputs"
    output_folder.mkdir(exist_ok=True)
    print(f"Found {len(files_to_convert)} file(s). Saving to: {output_folder}\n")

    # 4. Process Loop
    success_count = 0
    for file in files_to_convert:
        if convert_notebook_to_pdf(file, output_folder):
            print(f"  [SUCCESS] {file.name}")
            success_count += 1
        else:
            print(f"  [FAILED]  {file.name}")

    print("\n" + "=" * 50)
    print(f"CONVERSION COMPLETE: {success_count}/{len(files_to_convert)} successful.")
    print("=" * 50)
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()