"""
Build script for the Infant Report Generator.
Run this from inside the report_generator folder:

    python build.py

It will produce a 'dist/ReportGenerator' folder you can zip and send to testers.
"""

import subprocess
import sys
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# Locate the tkinterdnd2 package folder so PyInstaller bundles its DLL files
import tkinterdnd2
TKDND_PATH = os.path.dirname(tkinterdnd2.__file__)

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm",               # overwrite previous build without asking
    "--windowed",                # no console/terminal window when the app runs
    "--name", "ReportGenerator", # name of the output exe and folder
    "--icon", "NONE",            # no custom icon (add one later if you want)

    # Bundle the images folder so the reference diagrams are included
    "--add-data", f"images{os.pathsep}images",

    # Bundle the report template
    "--add-data", f"report_template.docx{os.pathsep}.",

    # Bundle tkinterdnd2's internal files (drag-and-drop support)
    "--add-data", f"{TKDND_PATH}{os.pathsep}tkinterdnd2",

    os.path.join(HERE, "app.py"),
]

print("Running PyInstaller with the following command:")
print(" ".join(cmd))
print()

result = subprocess.run(cmd, cwd=HERE)

if result.returncode == 0:
    dist_path = os.path.join(HERE, "dist", "ReportGenerator")
    print()
    print("=" * 60)
    print("BUILD SUCCESSFUL!")
    print(f"Your app is ready at:\n  {dist_path}")
    print()
    print("To share with a tester:")
    print("  1. Zip the entire 'ReportGenerator' folder")
    print("  2. Send the zip file")
    print("  3. They unzip it and double-click ReportGenerator.exe")
    print("=" * 60)
else:
    print()
    print("=" * 60)
    print("BUILD FAILED — check the error messages above.")
    print("=" * 60)
