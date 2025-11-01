"""
Script: GDL Object Extractor
Description: Extracts .gsm files to HSF format using LP_XMLConverter
Features:
    - Auto-detects Archicad and LP_XMLConverter location
    - File picker dialog for selecting .gsm files
    - Creates organized extraction folder
    - Extracts to HSF format with images
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


def select_gsm_file():
    """Open a file dialog to select a .gsm file"""
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring dialog to front

    file_path = filedialog.askopenfilename(
        title="Select GDL Object to Extract",
        filetypes=[
            ("Archicad GDL Objects", "*.gsm"),
            ("All files", "*.*")
        ]
    )

    root.destroy()
    return file_path if file_path else None


def create_extraction_folder(gsm_path):
    """
    Create extraction folder in same directory as source file
    Folder name: [object_name]_GDLextract
    """
    gsm_file = Path(gsm_path)
    gsm_dir = gsm_file.parent
    gsm_name = gsm_file.stem  # filename without extension

    extract_folder = gsm_dir / f"{gsm_name}_GDLextract"

    # Create folder if it doesn't exist
    extract_folder.mkdir(exist_ok=True)

    return str(extract_folder)


def extract_gdl_to_hsf(converter_path, gsm_path, output_folder):
    """
    Extract .gsm file to HSF format using LP_XMLConverter

    Command: "LP_XMLConverter" libpart2hsf "binaryFilePath" "hsfFolderPath"
    """
    try:
        # Build the command
        cmd = [
            converter_path,
            "libpart2hsf",
            gsm_path,
            output_folder
        ]

        print("\n" + "="*60)
        print("EXTRACTION COMMAND:")
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
        print(f"\n❌ Conversion failed!")
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
    print("GDL OBJECT EXTRACTOR")
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

    # Step 3: Select GDL file
    print("\n3️⃣  Select GDL object to extract...")
    gsm_path = select_gsm_file()

    if not gsm_path:
        print("❌ No file selected")
        return

    print(f"✓ Selected: {gsm_path}")

    # Step 4: Create extraction folder
    print("\n4️⃣  Creating extraction folder...")
    output_folder = create_extraction_folder(gsm_path)
    print(f"✓ Extraction folder: {output_folder}")

    # Step 5: Extract
    print("\n5️⃣  Extracting GDL object...")
    success = extract_gdl_to_hsf(converter_path, gsm_path, output_folder)

    # Summary
    print("\n" + "="*60)
    if success:
        print("✅ EXTRACTION COMPLETED SUCCESSFULLY!")
        print("\n📁 Output location:")
        print(f"   {output_folder}")
        print("\n💡 The HSF folder contains:")
        print("   • Script files (.gdl)")
        print("   • Parameter definitions")
        print("   • Images folder")
        print("   • All editable components")
        print("\n📝 Next steps:")
        print("   • Edit the extracted files as needed")
        print("   • Use xml2libpart or hsf2libpart to rebuild the .gsm")
    else:
        print("❌ EXTRACTION FAILED")
        print("   Check the error messages above for details")
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