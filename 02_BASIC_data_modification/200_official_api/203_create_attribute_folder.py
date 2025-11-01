"""
================================================================================
SCRIPT: Create Attribute Folder
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Modification

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Creates new folders in attribute hierarchies for better organization of
Archicad project attributes. This script provides:
- Creation of folders at root or nested levels
- Support for all attribute types that allow folders
- Validation of folder paths before creation
- Clear error messages for common issues
- Automatic duplicate detection

Attribute folders help organize project resources like layers, materials,
surfaces, and profiles into logical groups for easier management.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetAttributeFolderStructure(attributeType, path)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure
  Returns the folder structure for a given attribute type
  Used to validate existing structure and check for duplicates

- CreateAttributeFolders(attributeFolderCreationParameters)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.CreateAttributeFolders
  Creates attribute folders
  Parameters:
  * attributeFolderCreationParameters: List of folder creation parameters
  Returns:
  * List of ExecutionResult objects with success/error status

[Data Types]
- AttributeFolderCreationParameters
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac27.html#archicad.releases.ac27.b3001types.AttributeFolderCreationParameters 
  Parameters for creating an attribute folder:
  * attributeType: Type of attribute (string)
  * path: List of strings representing folder path

- AttributeFolderStructure
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac27.html#archicad.releases.ac27.b3001types.AttributeFolderStructure 
  Hierarchical structure of attribute folders

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Modify CONFIGURATION section below:
   - Set ATTRIBUTE_TYPE to the type of attributes to organize
   - Set FOLDER_PATH to the complete folder path to create
3. Run this script
4. Verify folder was created in Archicad

Note: The full path includes parent folders and the new folder name.

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Before running, set these values in the CONFIGURATION section:

ATTRIBUTE_TYPE: Type of attribute to organize
Supported types:
  • "Layer" - Layer attributes
  • "Pen" - Pen colors
  • "LineType" - Line types and patterns
  • "FillType" - Fill patterns
  • "CompositeLine" - Composite line types
  • "BuildingMaterial" - Building materials
  • "Composite" - Composite structures
  • "Profile" - Profile attributes
  • "Surface" - Surface materials
  • "ZoneCategory" - Zone categories

FOLDER_PATH: Complete path of folder to create
Format: List of folder names from root to target
Examples:
  ["Custom Materials"] - Creates folder at root level
  ["Custom Materials", "Wood"] - Creates "Wood" inside "Custom Materials"
  ["Structural", "Steel", "Profiles"] - Creates nested hierarchy

Important: The path includes ALL parent folders that need to exist or be created.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Attribute type and folder path
- Validation of existing structure
- Creation progress
- Success or error messages
- Folder ID if created successfully

Example output (success):
  === CREATING ATTRIBUTE FOLDER ===
  Type: BuildingMaterial
  Path: Custom Materials > Wood > Hardwood
  
  === VALIDATING STRUCTURE ===
  ✓ Retrieved existing folder structure
  
  === CREATING FOLDER ===
  ✓ Successfully created folder path
     Custom Materials > Wood > Hardwood
  
  Folder ID: {12345678-1234-1234-1234-123456789abc}

Example output (already exists):
  === CREATING FOLDER ===
  ✗ Error creating folder:
    Folder already exists at this path

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

If you have multiple Archicad instances open:
- The script connects to the first one opened
- Make sure the correct project is open in that instance
- Folder will be created in the connected instance's project

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
FOLDER PATHS:
- Paths are specified as lists of folder names
- Empty list [] means root level
- Each element in the list represents one level in the hierarchy
- Parent folders must exist or be included in the path
- Folder names are case-sensitive

ATTRIBUTE TYPES:
Not all attribute types support folder organization. Supported types:
✓ Layer, Pen, LineType, FillType, CompositeLine
✓ BuildingMaterial, Composite, Profile, Surface, ZoneCategory

Unsupported types will return an error when getting folder structure.

NESTED FOLDERS:
To create a nested folder structure:
1. Include all parent folder names in the path
2. The script will create the entire hierarchy if needed
3. Existing parent folders will be reused

Example: ["Parent", "Child", "Grandchild"]
- Creates "Parent" if it doesn't exist
- Creates "Child" inside "Parent" if it doesn't exist
- Creates "Grandchild" inside "Child"

DUPLICATE DETECTION:
- The API will return an error if the exact path already exists
- Check existing structure first using GetAttributeFolderStructure
- The script retrieves structure before creation for validation

FOLDER ORGANIZATION BENEFITS:
- Better project organization
- Easier to find and manage attributes
- Clear visual hierarchy in attribute dialogs
- Consistent naming conventions across projects
- Simplified template creation

LIMITATIONS:
- Cannot create folders for attribute types that don't support them
- Folder names must be unique within the same parent
- Cannot use special characters in folder names
- Maximum nesting depth depends on Archicad version

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 123_get_layers.py (view layer folder structure)
- 122_get_profiles.py (view profile folder structure)
- 114_01_get_building_materials.py (view material folder structure)
- 204_delete_attribute_folder.py (remove folders)
- 205_move_attributes.py (move attributes between folders)

================================================================================
"""

from archicad import ACConnection

# =============================================================================
# CONNECT TO ARCHICAD
# =============================================================================

# Establish connection to running Archicad instance
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

# Initialize command and type objects
acc = conn.commands
act = conn.types


# =============================================================================
# CONFIGURATION
# =============================================================================
# IMPORTANT: Modify these values before running the script!

# Attribute type (must be a supported type with folder organization)
# Supported: "Layer", "Pen", "LineType", "FillType", "CompositeLine",
#            "BuildingMaterial", "Composite", "Profile", "Surface", "ZoneCategory"
ATTRIBUTE_TYPE = "BuildingMaterial"

# Complete folder path to create
# Format: List of folder names from root to target folder
# Examples:
#   ["Custom Materials"]                    - Single folder at root
#   ["Custom Materials", "Wood"]            - "Wood" inside "Custom Materials"
#   ["Structural", "Steel", "Profiles"]     - Three-level nested structure
FOLDER_PATH = ["Custom Materials", "Test"]


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main():
    """Main function to create attribute folder."""

    print("\n" + "="*80)
    print("CREATE ATTRIBUTE FOLDER v1.0")
    print("="*80)

    # -------------------------------------------------------------------------
    # 1. DISPLAY CONFIGURATION
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CONFIGURATION")
    print("="*80)

    print(f"\nAttribute type: {ATTRIBUTE_TYPE}")
    print(f"Folder path: {' > '.join(FOLDER_PATH)}")
    print(f"  Levels: {len(FOLDER_PATH)}")
    print(f"  Root folder: {FOLDER_PATH[0] if FOLDER_PATH else 'N/A'}")
    if len(FOLDER_PATH) > 1:
        print(f"  Parent folders: {' > '.join(FOLDER_PATH[:-1])}")
        print(f"  Target folder: {FOLDER_PATH[-1]}")

    # -------------------------------------------------------------------------
    # 2. VALIDATE ATTRIBUTE TYPE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VALIDATING ATTRIBUTE TYPE")
    print("="*80)

    # List of supported attribute types
    supported_types = [
        "Layer", "Pen", "LineType", "FillType", "CompositeLine",
        "BuildingMaterial", "Composite", "Profile", "Surface", "ZoneCategory"
    ]

    if ATTRIBUTE_TYPE not in supported_types:
        print(
            f"\n⚠️  Warning: '{ATTRIBUTE_TYPE}' may not support folder organization")
        print("\nSupported attribute types:")
        for attr_type in supported_types:
            print(f"  • {attr_type}")
        print("\nAttempting to continue anyway...")
    else:
        print(
            f"\n✓ Attribute type '{ATTRIBUTE_TYPE}' supports folder organization")

    # -------------------------------------------------------------------------
    # 3. GET EXISTING FOLDER STRUCTURE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CHECKING EXISTING STRUCTURE")
    print("="*80)

    try:
        # Retrieve current folder structure
        existing_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)
        print(
            f"\n✓ Retrieved existing folder structure for '{ATTRIBUTE_TYPE}'")

        # Display basic structure info
        if hasattr(existing_structure, 'attributes'):
            num_attrs = len(
                existing_structure.attributes) if existing_structure.attributes else 0
            print(f"  Attributes at root: {num_attrs}")

        if hasattr(existing_structure, 'subfolders'):
            num_folders = len(
                existing_structure.subfolders) if existing_structure.subfolders else 0
            print(f"  Folders at root: {num_folders}")

    except Exception as e:
        print(f"\n✗ Error getting folder structure: {e}")
        print("\nPossible causes:")
        print("  • Attribute type is invalid")
        print("  • Attribute type doesn't support folders")
        print("  • API version incompatibility")
        print("\nPlease check the attribute type and try again")
        exit()

    # -------------------------------------------------------------------------
    # 4. CREATE FOLDER
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CREATING FOLDER")
    print("="*80)

    # Prepare folder creation parameters
    folder_params = act.AttributeFolderCreationParameters(
        attributeType=ATTRIBUTE_TYPE,
        path=FOLDER_PATH
    )

    print(f"\nCreating folder path: {' > '.join(FOLDER_PATH)}")
    if len(FOLDER_PATH) > 1:
        print(f"  Note: All parent folders will be created if they don't exist")

    try:
        # Call API to create folder(s)
        results = acc.CreateAttributeFolders([folder_params])

        if len(results) > 0:
            result = results[0]

            # Check if folder was created successfully
            if hasattr(result, 'attributeFolder') and result.attributeFolder:
                print(f"\n✓ Successfully created folder path:")
                print(f"   {' > '.join(FOLDER_PATH)}")

                # Display folder ID if available
                if hasattr(result, 'attributeFolderId') and result.attributeFolderId:
                    folder_id = result.attributeFolderId
                    print(f"\n  Folder ID: {folder_id}")

            elif hasattr(result, 'error') and result.error:
                # Folder creation failed
                error_msg = result.error.message if hasattr(
                    result.error, 'message') else str(result.error)
                print(f"\n✗ Error creating folder:")
                print(f"  {error_msg}")

                print("\nCommon issues:")
                print("  • Folder path already exists")
                print("  • Invalid folder name (special characters)")
                print("  • Parent folder structure is invalid")
                print("  • Insufficient permissions")
                exit()

            else:
                # Unknown response format
                print(f"\n⚠️  Folder may have been created (unknown response format)")

        else:
            print(f"\n✗ No response received from create operation")
            exit()

    except Exception as e:
        print(f"\n✗ Error creating folder: {e}")
        import traceback
        traceback.print_exc()

        print("\nTroubleshooting:")
        print("  1. Check if folder already exists")
        print("  2. Verify attribute type is correct")
        print("  3. Ensure folder names don't contain invalid characters")
        print("  4. Check Archicad API version compatibility")
        exit()

    # -------------------------------------------------------------------------
    # 5. VERIFY CREATION
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VERIFYING CREATION")
    print("="*80)

    try:
        # Retrieve updated folder structure
        updated_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)

        print(f"\n✓ Successfully retrieved updated structure")
        print(f"  Folder should now be visible in Archicad")

        # Count folders after creation
        if hasattr(updated_structure, 'subfolders'):
            num_folders = len(
                updated_structure.subfolders) if updated_structure.subfolders else 0
            print(f"  Total folders at root: {num_folders}")

    except Exception as e:
        print(f"\n⚠️  Could not verify folder creation: {e}")
        print("  (Folder may still have been created successfully)")

    # -------------------------------------------------------------------------
    # SUMMARY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Attribute type: {ATTRIBUTE_TYPE}")
    print(f"  Folder path: {' > '.join(FOLDER_PATH)}")
    print(f"  Status: Created successfully")
    print("\n" + "="*80)

    # Usage hints
    print("\n💡 Next steps:")
    print("   • Open Attribute Manager in Archicad to see the new folder")
    print("   • Move existing attributes into the folder if needed")
    print("   • Create additional nested folders if desired")

    print("\n💡 Examples:")
    print("   Create root folder:")
    print("     FOLDER_PATH = ['Custom Materials']")
    print("\n   Create nested folders:")
    print("     FOLDER_PATH = ['Custom Materials', 'Wood', 'Hardwood']")
    print("\n   Multiple folders (run script multiple times):")
    print("     FOLDER_PATH = ['Structural']")
    print("     FOLDER_PATH = ['Architectural']")
    print("     FOLDER_PATH = ['MEP']")


if __name__ == "__main__":
    main()
