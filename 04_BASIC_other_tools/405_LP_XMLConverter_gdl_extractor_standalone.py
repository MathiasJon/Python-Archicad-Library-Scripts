"""
Script: GDL Object Extractor (Standalone Version)
Description: Extracts .gsm files to HSF format using LP_XMLConverter
Features:
    - No Archicad/Tapir connection required
    - Configurable LP_XMLConverter path
    - File picker dialog for selecting .gsm files
    - Creates organized extraction folder
    - Extracts to HSF format with images
Requirements:
    - LP_XMLConverter installed with Archicad
    - tkinter (included with Python)
"""

import os
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
    print("GDL OBJECT EXTRACTOR (Standalone)")
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

    # Step 2: Select GDL file
    print("\n2️⃣  Select GDL object to extract...")
    gsm_path = select_gsm_file()

    if not gsm_path:
        print("❌ No file selected")
        return

    print(f"✓ Selected: {gsm_path}")

    # Step 3: Create extraction folder
    print("\n3️⃣  Creating extraction folder...")
    output_folder = create_extraction_folder(gsm_path)
    print(f"✓ Extraction folder: {output_folder}")

    # Step 4: Extract
    print("\n4️⃣  Extracting GDL object...")
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
        print("   • Use the compiler script to rebuild the .gsm")
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
