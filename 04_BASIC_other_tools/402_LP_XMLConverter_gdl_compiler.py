"""
Script: GDL Object Compiler (Version corrigée)
Description: Recompiles HSF folders back to .gsm files using LP_XMLConverter
Features:
    - Auto-detects Archicad and LP_XMLConverter location
    - Folder picker dialog for selecting HSF folder
    - Recognizes REAL HSF structure (scripts/, binaries/, metadata files)
    - Creates .gsm in the same location as the HSF folder
Requirements:
    - archicad-api package
    - Tapir Add-On installed
    - tkinter (included with Python)
"""

from archicad import ACConnection
import os
import subprocess
import platform
from tkinter import Tk, filedialog
from pathlib import Path


def get_archicad_location():
    """Get the Archicad executable location using Tapir Add-On"""
    try:
        conn = ACConnection.connect()
        acc = conn.commands
        act = conn.types

        response = acc.ExecuteAddOnCommand(
            act.AddOnCommandId('TapirCommand', 'GetArchicadLocation')
        )

        if 'archicadLocation' in response:
            return response['archicadLocation']
        else:
            return None
    except Exception as e:
        print(f"Error connecting to Archicad: {e}")
        return None


def get_lp_xmlconverter_path(archicad_path):
    """
    Get the LP_XMLConverter executable path based on Archicad location

    Windows: Same folder as Archicad executable
    macOS: Inside LP_XMLConverter.app bundle in MacOS folder
    """
    if not archicad_path:
        return None

    system = platform.system()

    if system == "Windows":
        archicad_dir = os.path.dirname(archicad_path)
        converter_path = os.path.join(archicad_dir, "LP_XMLConverter.exe")
        
    elif system == "Darwin":  # macOS
        # On macOS, archicad_path returns the .app bundle:
        # /Applications/Graphisoft/Archicad 27/Archicad 27.app
        
        # Determine the MacOS directory
        if archicad_path.endswith('.app'):
            # archicadLocation is the .app bundle itself
            macos_dir = os.path.join(archicad_path, "Contents", "MacOS")
        else:
            # archicadLocation is the executable inside MacOS
            archicad_dir = os.path.dirname(archicad_path)
            if archicad_dir.endswith("MacOS"):
                macos_dir = archicad_dir
            else:
                macos_dir = archicad_dir
        
        # LP_XMLConverter is inside LP_XMLConverter.app bundle
        # Path: .../MacOS/LP_XMLConverter.app/Contents/MacOS/LP_XMLConverter
        converter_path = os.path.join(
            macos_dir, "LP_XMLConverter.app", "Contents", "MacOS", "LP_XMLConverter")
        
    else:
        print(f"Unsupported operating system: {system}")
        return None

    if os.path.exists(converter_path):
        return converter_path
    else:
        print(f"LP_XMLConverter not found at: {converter_path}")
        return None


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
    Places it in the parent folder of the HSF folder with the HSF folder name
    """
    hsf_folder = Path(hsf_path)
    parent_dir = hsf_folder.parent

    # Remove _GDLextract suffix if present
    folder_name = hsf_folder.name
    suffix = "_GDLextract"
    if folder_name.endswith(suffix):
        # Use removesuffix (Python 3.9+) or manual removal
        try:
            object_name = folder_name.removesuffix(suffix)
        except AttributeError:
            # Python < 3.9: remove suffix manually
            object_name = folder_name[:-len(suffix)]
    else:
        object_name = folder_name

    output_gsm = parent_dir / f"{object_name}_compiled.gsm"

    # If file exists, add a number
    counter = 1
    while output_gsm.exists():
        output_gsm = parent_dir / f"{object_name}_compiled_{counter}.gsm"
        counter += 1

    return str(output_gsm)


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
    print("GDL OBJECT COMPILER (HSF → GSM)")
    print("="*60)

    # Step 1: Get Archicad location
    print("\n1️⃣  Locating Archicad...")
    archicad_path = get_archicad_location()

    if not archicad_path:
        print("❌ Could not locate Archicad")
        print("   Make sure Archicad is running and Tapir Add-On is installed")
        return

    print(f"✓ Archicad found: {archicad_path}")

    # Step 2: Get LP_XMLConverter location
    print("\n2️⃣  Locating LP_XMLConverter...")
    converter_path = get_lp_xmlconverter_path(archicad_path)

    if not converter_path:
        print("❌ Could not locate LP_XMLConverter")
        print("   Make sure it's installed in the Archicad directory")
        return

    print(f"✓ LP_XMLConverter found: {converter_path}")

    # Step 3: Select HSF folder
    print("\n3️⃣  Select HSF folder to compile...")
    hsf_path = select_hsf_folder()

    if not hsf_path:
        print("❌ No folder selected")
        return

    print(f"✓ Selected: {hsf_path}")

    # Step 4: Validate HSF folder
    print("\n4️⃣  Validating HSF folder...")
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

    # Step 5: Determine output path
    print("\n5️⃣  Determining output path...")
    output_gsm = get_output_gsm_path(hsf_path)
    print(f"✓ Output will be: {output_gsm}")

    # Step 6: Compile
    print("\n6️⃣  Compiling HSF to .gsm...")
    success = compile_hsf_to_gsm(converter_path, hsf_path, output_gsm)

    # Summary
    print("\n" + "="*60)
    if success:
        print("✅ COMPILATION COMPLETED SUCCESSFULLY!")
        print("\n📦 Your compiled GDL object:")
        print(f"   {output_gsm}")
        print("\n💡 You can now:")
        print("   • Load it in Archicad")
        print("   • Test the modifications you made")
        print("   • Share it with others")
        print("   • Add it to your library")

        # Check file size
        if os.path.exists(output_gsm):
            size = os.path.getsize(output_gsm)
            print(f"\n📊 File size: {size:,} bytes ({size/1024:.1f} KB)")
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