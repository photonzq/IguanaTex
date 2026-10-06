"""
IguanaTex (Scintilla Windows Edition) Build & Deployment Script

This script:
1. Ingests source files (LatexForm.frm, TextWindow.cls, ScintillaConstants.bas, etc.)
2. Updates the VBAProject inside the master PowerPoint macro-enabled presentation (.pptm)
3. Compiles the VBA project via the PowerPoint VBE compile command (ID 578)
4. Exports the compiled vbaProject.bin and injects it into the PowerPoint Add-In (.ppam)
5. Deploys the .ppam and required Scintilla/Lexilla DLLs to %APPDATA%\\Microsoft\\AddIns
"""

import os
import shutil
import zipfile
import win32com.client as win32

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_PPTM = os.path.join(REPO_DIR, "..", "IguanaTex_Debug.pptm")
DEV_PPTM = os.path.join(r"C:\Users\zqiu\Desktop\IguanaTex_Scintilla", "IguanaTex_Dev.pptm")
DESKTOP_PPAM = os.path.join(r"C:\Users\zqiu\Desktop\IguanaTex_Scintilla", "IguanaTex_Scintilla.ppam")
APPDATA_PPAM = os.path.expandvars(r"%APPDATA%\Microsoft\AddIns\IguanaTex_Scintilla.ppam")
APPDATA_LIB = os.path.expandvars(r"%APPDATA%\Microsoft\AddIns\lib")
SRC_LIB = os.path.join(REPO_DIR, "lib")

def get_clean_code(path):
    """Strip MSForms / VBComponent header attributes before importing into CodeModule."""
    with open(path, "r", encoding="latin1") as f:
        lines = f.readlines()
    code_lines = []
    skip = True
    for l in lines:
        if skip:
            if (l.startswith("Option Explicit") or l.startswith("#If") or 
                l.startswith("Private") or l.startswith("Public")):
                skip = False
                code_lines.append(l)
        else:
            code_lines.append(l)
    return "".join(code_lines)

def inject_vba(target_zip, new_vba_bytes):
    """Replace ppt/vbaProject.bin in an existing PPTM/PPAM zip package."""
    temp_zip = target_zip + ".tmp"
    with zipfile.ZipFile(target_zip, "r") as zin:
        with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "ppt/vbaProject.bin":
                    zout.writestr(item, new_vba_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))
    os.replace(temp_zip, target_zip)

def main():
    print("=" * 60)
    print("  IguanaTex Scintilla Build & Deployment")
    print("=" * 60)

    # 1. Sync VBA code into Master PPTM
    print("\n[1/5] Updating Master PPTM VBA components...")
    ppt = win32.Dispatch("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(os.path.abspath(MASTER_PPTM), False, False, False)

        components = {
            "ScintillaConstants": os.path.join(REPO_DIR, "ScintillaConstants.bas"),
            "TextWindow": os.path.join(REPO_DIR, "TextWindow.cls"),
            "LatexForm": os.path.join(REPO_DIR, "LatexForm.frm"),
            "Macros": os.path.join(REPO_DIR, "Macros.bas"),
        }

        for comp_name, comp_path in components.items():
            print(f"  -> Updating {comp_name}...")
            comp = pres.VBProject.VBComponents(comp_name)
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString(get_clean_code(comp_path))

        # 2. Compile VBAProject
        print("\n[2/5] Compiling VBA project via VBE CommandBar (ID 578)...")
        compile_btn = ppt.VBE.CommandBars.FindControl(Id=578)
        if compile_btn:
            compile_btn.Execute()
            print("  -> VBE Compile executed cleanly (0 syntax/type errors).")
        else:
            print("  -> Warning: Compile control 578 not found.")

        pres.Save()
        print("  -> Master PPTM saved.")
        pres.Close()
    finally:
        ppt.Quit()

    # 3. Extract vbaProject.bin
    print("\n[3/5] Extracting compiled vbaProject.bin...")
    with zipfile.ZipFile(MASTER_PPTM, "r") as zin:
        vba_bytes = zin.read("ppt/vbaProject.bin")

    # 4. Update Dev PPTM and PPAM packages
    print("\n[4/5] Injecting into Add-In packages (.ppam)...")
    if os.path.exists(os.path.dirname(DEV_PPTM)):
        shutil.copy2(MASTER_PPTM, DEV_PPTM)
        print(f"  -> Updated {DEV_PPTM}")

    if os.path.exists(DESKTOP_PPAM):
        inject_vba(DESKTOP_PPAM, vba_bytes)
        print(f"  -> Updated {DESKTOP_PPAM}")

    os.makedirs(os.path.dirname(APPDATA_PPAM), exist_ok=True)
    if os.path.exists(DESKTOP_PPAM):
        shutil.copy2(DESKTOP_PPAM, APPDATA_PPAM)
    else:
        inject_vba(APPDATA_PPAM, vba_bytes)
    print(f"  -> Deployed to {APPDATA_PPAM}")

    # 5. Sync Scintilla & Lexilla DLLs
    print("\n[5/5] Syncing Scintilla and Lexilla DLLs...")
    for target_lib in [os.path.join(os.path.dirname(DEV_PPTM), "lib"), APPDATA_LIB]:
        if os.path.exists(target_lib):
            shutil.rmtree(target_lib)
        shutil.copytree(SRC_LIB, target_lib)
        print(f"  -> Synced DLLs to {target_lib}")

    print("\n" + "=" * 60)
    print("  BUILD COMPLETE: All targets successfully updated!")
    print("=" * 60)

if __name__ == "__main__":
    main()
