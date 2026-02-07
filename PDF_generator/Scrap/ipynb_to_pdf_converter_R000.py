import sys
import os
from pathlib import Path
import subprocess

def convert_notebook_to_pdf(input_path: Path):
    if input_path.suffix.lower() != '.ipynb':
        print(f"  [SKIP] Not a notebook: {input_path.name}")
        return False

    print(f"  Converting: {input_path.name}...")
    
    try:
        # Runs the nbconvert command via subprocess
        subprocess.run([
            "jupyter", "nbconvert", 
            "--to", "webpdf", 
            "--allow-chromium-download",
            str(input_path)
        ], check=True)
        
        print(f"  [SUCCESS] Created PDF for: {input_path.name}")
        return True
    except Exception as e:
        print(f"  [ERROR] Conversion failed: {e}")
        return False

if __name__ == "__main__":
    # Logic to handle drag-and-drop or manual input
    path_input = sys.argv[1] if len(sys.argv) > 1 else input("Enter .ipynb path: ")
    p = Path(path_input.strip('"'))
    
    if p.is_file():
        convert_notebook_to_pdf(p)
    else:
        print("Please provide a valid .ipynb file.")
    input("\nPress Enter to exit...")