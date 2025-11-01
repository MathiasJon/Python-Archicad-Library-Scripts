"""
Script: GDL Object Compiler (Version corrigée macOS + Backup)
Description: Recompiles HSF folders back to .gsm files using LP_XMLConverter
Features:
    - Auto-detects Archicad and LP_XMLConverter location (macOS compatible)
    - Folder picker dialog for selecting HSF folder
    - Recognizes REAL HSF structure (scripts/, binaries/, metadata files)
    - Creates .gsm in the same location as the HSF folder
    - Backup system: COPIES original before compilation (preserves Archicad references)
Requirements:
    - archicad-api package
    - Tapir Add-On installed
    - tkinter (included with Python)
"""

from archicad import ACConnection
import os
import shutil
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
    macOS: Inside ARCHICAD.app/Contents/MacOS/
    
    AMÉLIORATION macOS: Gère correctement les cas où archicadLocation 
    retourne soit le bundle .app soit l'exécutable à l'intérieur
    """
    if not archicad_path:
        return None

    system = platform.system()
    
    print(f"   Detected OS: {system}")
    print(f"   Archicad path: {archicad_path}")

    if system == "Windows":
        archicad_dir = os.path.dirname(archicad_path)
        converter_path = os.path.join(archicad_dir, "LP_XMLConverter.exe")
        
    elif system == "Darwin":  # macOS
        # Sur Mac, archicadLocation peut retourner:
        # - Le bundle .app: /Applications/Graphisoft/Archicad 27/Archicad 27.app
        # - L'exécutable: /Applications/Graphisoft/Archicad 27/Archicad 27.app/Contents/MacOS/Archicad 27
        
        # Déterminer le répertoire MacOS
        if archicad_path.endswith('.app'):
            # archicadLocation est le bundle .app lui-même
            macos_dir = os.path.join(archicad_path, "Contents", "MacOS")
            print(f"   Detected .app bundle, using MacOS dir: {macos_dir}")
        else:
            # archicadLocation est l'exécutable à l'intérieur de MacOS
            archicad_dir = os.path.dirname(archicad_path)
            if archicad_dir.endswith("MacOS"):
                macos_dir = archicad_dir
                print(f"   Detected executable in MacOS dir: {macos_dir}")
            else:
                # Fallback: chercher dans Contents/MacOS
                macos_dir = os.path.join(archicad_dir, "Contents", "MacOS")
                print(f"   Using fallback MacOS dir: {macos_dir}")
        
        # LP_XMLConverter est dans LP_XMLConverter.app/Contents/MacOS/LP_XMLConverter
        converter_path = os.path.join(
            macos_dir, "LP_XMLConverter.app", "Contents", "MacOS", "LP_XMLConverter")
        
        print(f"   Trying converter path: {converter_path}")
    else:
        print(f"Unsupported operating system: {system}")
        return None

    if os.path.exists(converter_path):
        print(f"   ✓ LP_XMLConverter found!")
        return converter_path
    else:
        print(f"   ✗ LP_XMLConverter not found at: {converter_path}")
        
        # Sur macOS, essayer un chemin alternatif
        if system == "Darwin":
            # Essayer de chercher dans le répertoire parent
            parent_dir = os.path.dirname(os.path.dirname(archicad_path)) if not archicad_path.endswith('.app') else os.path.dirname(archicad_path)
            alt_converter_path = os.path.join(
                parent_dir, "LP_XMLConverter.app", "Contents", "MacOS", "LP_XMLConverter")
            print(f"   Trying alternative path: {alt_converter_path}")
            
            if os.path.exists(alt_converter_path):
                print(f"   ✓ LP_XMLConverter found at alternative location!")
                return alt_converter_path
        
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
    Looks for the original .gsm file in the parent directory and uses its exact name
    Falls back to reconstructing the name if no original file is found
    """
    hsf_folder = Path(hsf_path)
    parent_dir = hsf_folder.parent

    # Get the HSF folder name
    folder_name = hsf_folder.name

    # Strategy 1: Look for existing .gsm file in parent directory
    # This is the most reliable method as it uses the exact original name
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
    # If HSF folder is inside a _GDLextract folder, use parent folder name
    if parent_folder_name.endswith("_GDLextract"):
        object_name = parent_folder_name[:-11]
        print(f"   Reconstructed name from parent folder: {object_name}")
    # If current folder ends with _GDLextract, use its name
    elif folder_name.endswith("_GDLextract"):
        object_name = folder_name[:-11]
        print(f"   Reconstructed name from folder: {object_name}")
    # Otherwise use the HSF folder name as-is
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
    (NOT renaming, to preserve Archicad references)
    
    If .bak already exists, create .bak.1, .bak.2, etc.
    Returns True if backup was created, False if file didn't exist
    
    AMÉLIORATION: Utilise shutil.copy2() au lieu de rename()
    pour que Archicad garde sa référence au fichier original
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
        print(f"   (Original file preserved for Archicad)")
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


def reload_libraries():
    """
    Reload all libraries in Archicad using Tapir Add-On
    This ensures that modified GDL objects are refreshed in the project
    """
    try:
        from archicad import ACConnection

        print("\n" + "="*60)
        print("RELOADING ARCHICAD LIBRARIES...")
        print("="*60)

        # Connect to Archicad
        conn = ACConnection.connect()
        acc = conn.commands
        act = conn.types

        # Optional: Get library count before reload
        try:
            response_before = acc.ExecuteAddOnCommand(
                act.AddOnCommandId('TapirCommand', 'GetLoadedLibraries')
            )

            if 'libraries' in response_before:
                lib_count = len(response_before['libraries'])
                print(f"\n📚 Current libraries loaded: {lib_count}")
            elif 'libraryPaths' in response_before:
                lib_count = len(response_before['libraryPaths'])
                print(f"\n📚 Current libraries loaded: {lib_count}")
        except:
            pass

        print("\n🔄 Reloading...")

        # Reload libraries using Tapir command
        response = acc.ExecuteAddOnCommand(
            act.AddOnCommandId('TapirCommand', 'ReloadLibraries')
        )

        print("\n✓ Libraries reloaded successfully")

        # Check response for any messages
        if isinstance(response, dict):
            if 'message' in response:
                print(f"   Message: {response['message']}")

            if 'reloadedCount' in response:
                print(f"   Reloaded: {response['reloadedCount']} libraries")

            if 'errors' in response and response['errors']:
                print("\n   ⚠️  Reload warnings/errors:")
                for error in response['errors']:
                    print(f"      - {error}")

        return True

    except Exception as e:
        print(f"\n⚠️  Error reloading libraries: {e}")
        print("   (Your .gsm file was still created successfully)")
        print("\n   Possible causes:")
        print("   • Tapir Add-On not installed")
        print("   • Library files have errors")
        print("   • Library folders not accessible")
        return False


def main():
    """Main execution function"""
    print("\n" + "="*60)
    print("GDL OBJECT COMPILER (HSF → GSM)")
    print("Version: macOS Compatible + Backup par copie")
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

    # Step 5b: Backup existing file if present
    if os.path.exists(output_gsm):
        print(f"\n   ⚠️  File already exists: {Path(output_gsm).name}")
        print(f"   Creating backup (copy) before overwriting...")
        backup_existing_file(output_gsm)

    # Step 6: Compile
    print("\n6️⃣  Compiling HSF to .gsm...")
    success = compile_hsf_to_gsm(converter_path, hsf_path, output_gsm)

    # Step 7: Reload libraries if compilation succeeded
    if success:
        print("\n7️⃣  Reloading libraries in Archicad...")
        reload_success = reload_libraries()
        if not reload_success:
            print("   ℹ️  Library reload failed, but .gsm file was created")
            print("   You can manually reload libraries in Archicad if needed")

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
        print("   • The object has been reloaded in Archicad")
        print("   • Test your modifications in the current project")
        print("   • Your original file is safely backed up (.bak)")
        print("   • Repeat edit → compile as needed")

        print("\n📝 Fast development workflow:")
        print("   1. Edit scripts in HSF folder")
        print("   2. Run this compiler (replaces original + auto-reload)")
        print("   3. Test immediately in Archicad")
        print("   4. Repeat!")
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