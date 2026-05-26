"""Run the final econometric modeling stage from the prepared model-ready dataset.
Place this file in the project root and run:
    python run_models_only.py
Outputs are written to ./outputs.
"""
import subprocess, sys, pathlib, shutil
root = pathlib.Path(__file__).resolve().parent
for script in ["05_baseline_models.py", "06_robustness_models.py"]:
    src = root / "scripts" / script
    print(f"Running {script} ...")
    subprocess.run([sys.executable, str(src)], check=True)
print("Done. Check output Excel files.")
