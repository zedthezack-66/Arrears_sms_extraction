XTENDA SMS Extractor - GUI build

Files added:
- gui_runner.py   : Tkinter GUI that runs the SMS extraction on a selected .xlsx file
- build_exe.ps1   : PowerShell script to build a single-file Windows exe using PyInstaller

How to build (Windows):
1. Install Python 3.10+ on the machine and ensure `python` is on PATH.
2. Open PowerShell in this folder and run: .\build_exe.ps1
3. After completion, Xtenda_SMS_Extractor.exe will be placed in this folder.

How to use the exe:
- Double-click Xtenda_SMS_Extractor.exe, click "Choose Excel", select the loanbook .xlsx, then click Run.
- Output will be saved to an `output` folder next to the exe as Xtenda_SMS_Extracts.xlsx

Notes:
- The exe is built with --windowed (no console). To debug, run gui_runner.py with Python directly.
- If the build fails due to missing wheels or slow network, re-run the build script or install dependencies manually.
