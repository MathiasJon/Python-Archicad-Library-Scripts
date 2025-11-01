"""
Script: GDL Object Compiler (Standalone Version)
Description: Recompiles HSF folders back to .gsm files using LP_XMLConverter
Features:
    - No Archicad/Tapir connection required
    - Configurable LP_XMLConverter path
    - Folder picker dialog for selecting HSF folder
    - Recognizes REAL HSF structure (scripts/, binaries/, metadata files)
    - Creates .gsm in the same location as the HSF folder
    - Backup system: COPIES original before compilation
Requirements:
    - LP_XMLConverter installed with Archicad
    - tkinter (included with Python)
"""

import os
import shutil
import subprocess
import platform
from tkinter import Tk, filedialog, messagebox
from pathlib import Path


# ============================================================================
# CONFIGURATION: Set your LP_XMLConverter path here
# ============================================================================

# Windows Example (Archicad 27):
# LP_XMLCONVERTER_PATH = r"C:\Program Files\Graphisoft\Archicad 27\LP_XMLConverter.exe"

# macOS Example (Archicad 27):
# LP_XMLCONVERTER_PATH = "/Applications/Graphisoft/Archicad 27/Archicad 27.app/Contents/MacOS/LP_XMLConverter.app/Contents/MacOS/LP_XMLConverter"

# Auto-detect based on OS (modify paths as needed):
if platform.system() == "Windows":
    LP_XMLCONVERTER_PATH = r"C:\Program Files\Graphisoft\Archicad 27\LP_XMLConverter.exe"
elif platform.system() == "Darwin":  # macOS
    LP_XMLCONVERTER_PATH = "/Applications/Graphisoft/Archicad 27/Archicad 27.app/Contents/MacOS/LP_XMLConverter.app/Contents/MacOS/LP_XMLConverter"
else:
    LP_XMLCONVERTER_PATH = None

# ============================================================================


def verify_converter_path():
    """Verify that LP_XMLConverter exists at the configured path"""
    if not LP_XMLCONVERTER_PATH:
        return False, "LP_XMLConverter path not configured for this operating system"
    
    if not os.path.exists(LP_XMLCONVERTER_PATH):
        return False, f"LP_XMLConverter not found at:\n{LP_XMLCONVERTER_PATH}\n\nPlease update the LP_XMLCONVERTER_PATH variable in the script."
    
    return True, LP_XMLCONVERTER_PATH


def select_hsf_folder():
    """Open a dialog to select an HSF folder"""
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring dialog to front

    folder_path = filedialog.askdirectory(
        title="Select HSF Folder to Compile",
        mustexist=True
    )

    root.destroy()
    return folder_path if folder_path else None


def validate_hsf_folder(hsf_path):
    """
    Validate that the selected folder is a valid HSF folder
    
    Real HSF structure created by LP_XMLConverter contains:
    - scripts/ folder (with .gdl script files)
    - binaries/ folder
    - Metadata files: ancestry, calledmacros, libpartdata, libpartdocs, paramlist
    """
    hsf_folder = Path(hsf_path)

    print(f"   Checking folder structure...")

    # Check for HSF structure indicators
    scripts_folder = hsf_folder / "scripts"
    binaries_folder = hsf_folder / "binaries"

    # HSF metadata files
    hsf_metadata_files = ["ancestry", "calledmacros",
                          "libpartdata", "libpartdocs", "paramlist"]
    found_metadata = [f for f in hsf_metadata_files if (
        hsf_folder / f).exists()]

    # Check if this looks like an HSF folder
    has_scripts = scripts_folder.exists() and scripts_folder.is_dir()
    has_binaries = binaries_folder.exists() and binaries_folder.is_dir()
    has_metadata = len(found_metadata) > 0

    if has_scripts or has_binaries or has_metadata:
        print(f"✓ Valid HSF structure detected!")
        if has_scripts:
            try:
                script_files = list(scripts_folder.iterdir())
                print(f"   • scripts/ folder: {len(script_files)} file(s)")
            except:
                print(f"   • scripts/ folder: found")
        if has_binaries:
            print(f"   • binaries/ folder: found")
        if has_metadata:
            print(f"   • Metadata files: {', '.join(found_metadata)}")
        return True

    # Maybe it's a parent folder? Check subfolders
    print(f"   HSF structure not found at root level")
    print(f"   Checking subfolders...")

    subfolders = [f for f in hsf_folder.iterdir()
                  if f.is_dir() and not f.name.startswith('.')]

    for subfolder in subfolders:
        scripts_folder = subfolder / "scripts"
        binaries_folder = subfolder / "binaries"
        found_metadata = [
            f for f in hsf_metadata_files if (subfolder / f).exists()]

        has_scripts = scripts_folder.exists() and scripts_folder.is_dir()
        has_binaries = binaries_folder.exists() and binaries_folder.is_dir()
        has_metadata = len(found_metadata) > 0

        if has_scripts or has_binaries or has_metadata:
            print(f"\n⚠️  HSF content found in subfolder: {subfolder.name}")
            if has_scripts:
                try:
                    script_files = list(scripts_folder.iterdir())
                    print(f"   • scripts/ folder: {len(script_files)} file(s)")
                except:
                    print(f"   • scripts/ folder: found")
            if has_binaries:
                print(f"   • binaries/ folder: found")
            if has_metadata:
                print(f"   • Metadata files: {', '.join(found_metadata)}")
            print(f"   Using subfolder: {subfolder}")
            return str(subfolder)

    # Show what we found for debugging
    print(f"\n   ❌ No HSF structure found")
    print(f"   Current folder contents:")
    try:
        items = list(hsf_folder.iterdir())[:15]
        for item in sorted(items):
            item_type = "📁" if item.is_dir() else "📄"
            print(f"      {item_type} {item.name}")
        if len(list(hsf_folder.iterdir())) > 15:
            print(
                f"      ... and {len(list(hsf_folder.iterdir())) - 15} more items")
    except Exception as e:
        print(f"      Error listing contents: {e}")

    print(f"\n   💡 Expected HSF structure:")
    print(f"      📁 scripts/")
    print(f"      📁 binaries/")
    print(f"      📄 ancestry")
    print(f"      📄 calledmacros")
    print(f"      📄 libpartdata")
    print(f"      📄 etc.")

    return False


def get_output_gsm_path(hsf_path):
    """
    Determine the output .gsm file path
    Looks for the original .gsm file in the parent directory and uses its exact name
    Falls back to reconstructing the name if no original file is found
    """
    hsf_folder = Path(hsf_path)
    parent_dir = hsf_folder.parent

    # Get the HSF folder name
    folder_name = hsf_folder.name

    # Strategy 1: Look for existing .gsm file in parent directory
    print(f"   Looking for original .gsm file in parent directory...")

    # Try to find a .gsm file that matches the folder structure
    possible_names = []

    # If folder ends with _GDLextract, remove it to get base name
    if folder_name.endswith("_GDLextract"):
        base_name = folder_name[:-11]
        possible_names.append(base_name)

    # Also check the parent folder name if it ends with _GDLextract
    parent_folder_name = parent_dir.name
    if parent_folder_name.endswith("_GDLextract"):
        base_name = parent_folder_name[:-11]
        possible_names.append(base_name)

    # Add the HSF folder name itself as a possibility
    possible_names.append(folder_name)

    # Look for existing .gsm files in parent directory
    gsm_files = list(parent_dir.glob("*.gsm"))

    # Try to match with possible names
    for gsm_file in gsm_files:
        gsm_name_without_ext = gsm_file.stem
        for possible_name in possible_names:
            if gsm_name_without_ext == possible_name:
                print(f"   ✓ Found original file: {gsm_file.name}")
                return str(gsm_file)

    # Strategy 2: Reconstruct name from folder structure
    if parent_folder_name.endswith("_GDLextract"):
        object_name = parent_folder_name[:-11]
        print(f"   Reconstructed name from parent folder: {object_name}")
    elif folder_name.endswith("_GDLextract"):
        object_name = folder_name[:-11]
        print(f"   Reconstructed name from folder: {object_name}")
    else:
        object_name = folder_name
        print(f"   Using HSF folder name: {object_name}")

    # Use reconstructed name
    output_gsm = parent_dir / f"{object_name}.gsm"
    print(f"   Will create: {output_gsm.name}")

    return str(output_gsm)


def backup_existing_file(file_path):
    """
    Create a backup of an existing file by COPYING it to .bak
    (NOT renaming, to preserve references)
    
    If .bak already exists, create .bak.1, .bak.2, etc.
    Returns True if backup was created, False if file didn't exist
    """
    original_file = Path(file_path)

    if not original_file.exists():
        return False

    # Find an available backup name
    backup_file = Path(str(original_file) + ".bak")
    counter = 1

    while backup_file.exists():
        backup_file = Path(f"{original_file}.bak.{counter}")
        counter += 1

    # COPY (not rename) the original file to backup
    try:
        shutil.copy2(original_file, backup_file)
        print(f"   💾 Backup created: {backup_file.name}")
        return True
    except Exception as e:
        print(f"   ⚠️  Could not create backup: {e}")
        return False


def compile_hsf_to_gsm(converter_path, hsf_path, output_gsm):
    """
    Compile HSF folder to .gsm file using LP_XMLConverter
    
    Command: "LP_XMLConverter" hsf2libpart "hsfFolderPath" "binaryFilePath"
    """
    try:
        # Build the command
        cmd = [
            converter_path,
            "hsf2libpart",
            hsf_path,
            output_gsm
        ]

        print("\n" + "="*60)
        print("COMPILATION COMMAND:")
        print(" ".join(f'"{arg}"' if " " in arg else arg for arg in cmd))
        print("="*60 + "\n")

        # Run the conversion
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True
        )

        # Print output
        if result.stdout:
            print("CONVERTER OUTPUT:")
            print(result.stdout)

        if result.stderr:
            print("CONVERTER MESSAGES:")
            print(result.stderr)

        return True

    except subprocess.CalledProcessError as e:
        print(f"\n❌ Compilation failed!")
        print(f"Error code: {e.returncode}")
        if e.stdout:
            print(f"Output: {e.stdout}")
        if e.stderr:
            print(f"Error: {e.stderr}")
        return False
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        return False


def main():
    """Main execution function"""
    print("\n" + "="*60)
    print("GDL OBJECT COMPILER (Standalone)")
    print("HSF → GSM")
    print("="*60)

    # Step 1: Verify LP_XMLConverter path
    print("\n1️⃣  Verifying LP_XMLConverter...")
    is_valid, result = verify_converter_path()
    
    if not is_valid:
        print(f"❌ {result}")
        print("\n💡 How to configure:")
        print("   1. Open this script in a text editor")
        print("   2. Find the LP_XMLCONVERTER_PATH variable at the top")
        print("   3. Set it to your LP_XMLConverter path")
        print("\n📝 Example paths:")
        print("   Windows: C:\\Program Files\\Graphisoft\\Archicad 27\\LP_XMLConverter.exe")
        print("   macOS: /Applications/Graphisoft/Archicad 27/Archicad 27.app/Contents/MacOS/LP_XMLConverter.app/Contents/MacOS/LP_XMLConverter")
        
        # Show error dialog
        root = Tk()
        root.withdraw()
        messagebox.showerror(
            "Configuration Error",
            f"{result}\n\nPlease edit the script and set the LP_XMLCONVERTER_PATH variable."
        )
        root.destroy()
        return

    converter_path = result
    print(f"✓ LP_XMLConverter found: {converter_path}")

    # Step 2: Select HSF folder
    print("\n2️⃣  Select HSF folder to compile...")
    hsf_path = select_hsf_folder()

    if not hsf_path:
        print("❌ No folder selected")
        return

    print(f"✓ Selected: {hsf_path}")

    # Step 3: Validate HSF folder
    print("\n3️⃣  Validating HSF folder...")
    validation = validate_hsf_folder(hsf_path)

    if not validation:
        print("\n❌ Selected folder does not appear to be a valid HSF folder")
        print("\n💡 HSF folders should contain:")
        print("   • scripts/ folder with .gdl files")
        print("   • binaries/ folder")
        print("   • Metadata files (ancestry, calledmacros, etc.)")
        return
    elif isinstance(validation, str):
        # Validation returned a corrected path
        hsf_path = validation

    print(f"✓ Valid HSF folder confirmed")

    # Step 4: Determine output path
    print("\n4️⃣  Determining output path...")
    output_gsm = get_output_gsm_path(hsf_path)
    print(f"✓ Output will be: {output_gsm}")

    # Step 4b: Backup existing file if present
    if os.path.exists(output_gsm):
        print(f"\n   ⚠️  File already exists: {Path(output_gsm).name}")
        print(f"   Creating backup (copy) before overwriting...")
        backup_existing_file(output_gsm)

    # Step 5: Compile
    print("\n5️⃣  Compiling HSF to .gsm...")
    success = compile_hsf_to_gsm(converter_path, hsf_path, output_gsm)

    # Summary
    print("\n" + "="*60)
    if success:
        print("✅ COMPILATION COMPLETED SUCCESSFULLY!")
        print("\n📦 Your GDL object has been updated:")
        print(f"   {output_gsm}")

        # Check file size
        if os.path.exists(output_gsm):
            size = os.path.getsize(output_gsm)
            print(f"   Size: {size:,} bytes ({size/1024:.1f} KB)")

        # Check for backup
        backup_path = Path(str(output_gsm) + ".bak")
        if backup_path.exists():
            print(f"\n💾 Original file backed up as:")
            print(f"   {backup_path.name}")

        print("\n💡 Next steps:")
        print("   • Load the .gsm file in Archicad")
        print("   • Test your modifications")
        print("   • Your original file is safely backed up (.bak)")
        print("   • Repeat edit → compile as needed")

        print("\n📝 Fast development workflow:")
        print("   1. Edit scripts in HSF folder")
        print("   2. Run this compiler")
        print("   3. Reload the library in Archicad")
        print("   4. Test immediately")
        print("   5. Repeat!")
    else:
        print("❌ COMPILATION FAILED")
        print("   Check the error messages above for details")
        print("\n💡 Common issues:")
        print("   • Missing or corrupted script files")
        print("   • Invalid GDL syntax in scripts")
        print("   • Missing required parameters")
        print("   • Corrupted HSF structure")
    print("="*60 + "\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Operation cancelled by user")
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
