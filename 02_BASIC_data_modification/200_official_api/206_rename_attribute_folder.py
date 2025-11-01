"""
================================================================================
SCRIPT: Rename Attribute Folder
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Modification

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Renames attribute folders in the current Archicad project. This script:
- Renames folders in attribute hierarchies
- Preserves folder contents and structure
- Validates new folder names
- Provides clear feedback on rename operations
- Maintains folder position in hierarchy

Useful for improving folder organization and naming consistency.

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
  Returns folder structure for finding the folder to rename

- RenameAttributeFolders(attributeFolderParametersList)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.RenameAttributeFolders
  Renames attribute folders
  Parameters:
  * attributeFolderParametersList: List of AttributeFolderRenameParameters
  Returns:
  * List of ExecutionResult objects with success/error status

[Data Types]
- AttributeFolderRenameParameters
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac27.html#archicad.releases.ac27.b3001types.AttributeFolderRenameParameters
  Used to rename an attribute folder. The folder is identified by its ID.
  Variables:
  * attributeFolderId: The identifier of an attribute folder
  * newName: The name of the folder. Legal names are not empty and do not 
             begin or end with whitespace

- AttributeFolderId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac26.html#archicad.releases.ac26.b3005types.AttributeFolder.attributeFolderId 
  The identifier of an attribute folder

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Modify CONFIGURATION section below:
   - Set ATTRIBUTE_TYPE to the type of folder
   - Set FOLDER_PATH to the folder to rename
   - Set NEW_FOLDER_NAME to the desired new name
3. Run this script
4. Verify folder was renamed in Archicad

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Before running, set these values in the CONFIGURATION section:

ATTRIBUTE_TYPE: Type of attribute folder to rename
Supported types:
  • "Layer", "Pen", "LineType", "FillType", "CompositeLine"
  • "BuildingMaterial", "Composite", "Profile", "Surface", "ZoneCategory"

FOLDER_PATH: Path to the folder to rename
Format: List of folder names from root to target
Examples:
  ["Old Name"] - Rename folder at root level
  ["Parent", "Old Name"] - Rename nested folder

NEW_FOLDER_NAME: New name for the folder
- Must not be empty
- Must not begin or end with whitespace
- Must be unique within parent folder
Example: "New Folder Name"

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Current folder path and type
- New folder name
- Validation results
- Rename operation status

Example output:
  === RENAME ATTRIBUTE FOLDER ===
  Type: BuildingMaterial
  Current path: Custom > Temporary
  New name: Archive
  
  === FINDING FOLDER ===
  ✓ Found folder to rename
    Current name: Temporary
    Parent: Custom
  
  === VALIDATING NEW NAME ===
  ✓ New name is valid
  
  === RENAMING FOLDER ===
  ✓ Successfully renamed folder
    From: Temporary
    To: Archive

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
RENAME BEHAVIOR:
- Only the folder name changes
- Folder contents remain unchanged
- Folder position in hierarchy stays the same
- Subfolders and attributes are not affected
- References to folder are automatically updated

FOLDER NAME RULES:
Valid folder names:
✓ Not empty
✓ No leading whitespace
✓ No trailing whitespace
✓ Unique within parent folder
✓ Can contain spaces in middle
✓ Can contain numbers, letters, symbols

Invalid folder names:
✗ Empty string ""
✗ Only whitespace "   "
✗ Leading spaces "  Name"
✗ Trailing spaces "Name  "
✗ Duplicate of sibling folder name

NAME VALIDATION:
The script validates:
- Name is not empty
- Name doesn't start/end with whitespace
- Folder to rename exists
- Folder can be identified by its path

Note: API validates uniqueness, not this script

FOLDER IDENTIFICATION:
Folders are identified by:
1. Path from root (list of folder names)
2. AttributeFolderId obtained from structure

The script searches for the folder using its path,
then renames it using its ID.

COMMON USE CASES:
- Fixing typos in folder names
- Updating naming conventions
- Improving folder organization
- Standardizing folder names across projects
- Clarifying folder purposes

LIMITATIONS:
- Cannot rename system/default folders
- Cannot rename root folder
- New name must be unique in parent
- Cannot rename locked folders
- Name validation rules enforced by API

BEST PRACTICES:
- Use clear, descriptive names
- Follow consistent naming convention
- Avoid special characters if possible
- Keep names concise but meaningful
- Test on sample project first

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 203_create_attribute_folder.py (create folders)
- 204_delete_attribute_folder.py (delete folders)
- 205_move_attributes.py (move folders)
- 123_get_layers.py (view folder structure)

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

# Attribute type containing the folder to rename
ATTRIBUTE_TYPE = "BuildingMaterial"

# Path to the folder to rename
# Example: ["Temporary"], ["Custom", "Old Name"]
FOLDER_PATH = ["Temporary Materials"]

# New name for the folder
# Rules: Not empty, no leading/trailing whitespace
NEW_FOLDER_NAME = "Archive Materials"


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def find_folder_in_structure(structure, path, current_path=[]):
    """
    Recursively search for a folder in the attribute structure.

    Args:
        structure: AttributeFolderStructure object
        path: List of folder names to search for
        current_path: Current path in recursion

    Returns:
        Tuple of (AttributeFolderId, current_name, parent_path) if found,
        else (None, None, None)
    """
    if not path:
        # Found the target folder
        folder_id = structure.attributeFolderId if hasattr(
            structure, 'attributeFolderId') else None
        folder_name = structure.name if hasattr(structure, 'name') else None
        return (folder_id, folder_name, current_path)

    # Search in subfolders
    target_name = path[0]
    remaining_path = path[1:]

    if hasattr(structure, 'subfolders') and structure.subfolders:
        for subfolder_item in structure.subfolders:
            if hasattr(subfolder_item, 'attributeFolder'):
                subfolder = subfolder_item.attributeFolder
                if hasattr(subfolder, 'name') and subfolder.name == target_name:
                    return find_folder_in_structure(
                        subfolder,
                        remaining_path,
                        current_path + [target_name]
                    )

    return (None, None, None)


def validate_folder_name(name):
    """
    Validate folder name according to API rules.

    Args:
        name: Folder name to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not name:
        return (False, "Name cannot be empty")

    if name != name.strip():
        return (False, "Name cannot begin or end with whitespace")

    return (True, None)


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main():
    """Main function to rename attribute folder."""

    print("\n" + "="*80)
    print("RENAME ATTRIBUTE FOLDER v1.0")
    print("="*80)

    # -------------------------------------------------------------------------
    # 1. DISPLAY CONFIGURATION
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CONFIGURATION")
    print("="*80)

    print(f"\nAttribute type: {ATTRIBUTE_TYPE}")
    print(f"Current path: {' > '.join(FOLDER_PATH)}")
    print(f"New name: {NEW_FOLDER_NAME}")

    # -------------------------------------------------------------------------
    # 2. VALIDATE NEW NAME
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VALIDATING NEW NAME")
    print("="*80)

    is_valid, error_msg = validate_folder_name(NEW_FOLDER_NAME)

    if not is_valid:
        print(f"\n✗ Invalid folder name: {error_msg}")
        print("\nFolder name rules:")
        print("  • Must not be empty")
        print("  • Must not begin with whitespace")
        print("  • Must not end with whitespace")
        print("  • Must be unique within parent folder")
        exit()

    print(f"\n✓ New name is valid: '{NEW_FOLDER_NAME}'")

    # -------------------------------------------------------------------------
    # 3. FIND FOLDER TO RENAME
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("FINDING FOLDER")
    print("="*80)

    try:
        # Get folder structure
        print(f"\nSearching for folder in '{ATTRIBUTE_TYPE}' hierarchy...")
        root_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)

        # Search for target folder
        folder_id, current_name, parent_path = find_folder_in_structure(
            root_structure,
            FOLDER_PATH
        )

        if not folder_id:
            print(f"\n✗ Folder not found: {' > '.join(FOLDER_PATH)}")
            print("\nPlease check:")
            print("  • Folder path is correct")
            print("  • Folder exists in this attribute type")
            print("  • Path is case-sensitive")
            print(f"\n💡 Tip: Run script 123 or similar to see folder structure")
            exit()

        print(f"\n✓ Found folder to rename")
        print(f"  Current name: {current_name}")
        if parent_path:
            print(f"  Parent: {' > '.join(parent_path)}")
        else:
            print(f"  Location: Root level")

    except Exception as e:
        print(f"\n✗ Error getting folder structure: {e}")
        import traceback
        traceback.print_exc()
        exit()

    # -------------------------------------------------------------------------
    # 4. CHECK IF NAME CHANGE IS NEEDED
    # -------------------------------------------------------------------------

    if current_name == NEW_FOLDER_NAME:
        print(f"\n⚠️  Folder already has the name '{NEW_FOLDER_NAME}'")
        print("No renaming needed.")
        exit()

    # -------------------------------------------------------------------------
    # 5. RENAME FOLDER
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("RENAMING FOLDER")
    print("="*80)

    # Prepare rename parameters
    rename_params = act.AttributeFolderRenameParameters(
        attributeFolderId=folder_id,
        newName=NEW_FOLDER_NAME
    )

    print(f"\nRenaming folder...")
    print(f"  From: '{current_name}'")
    print(f"  To: '{NEW_FOLDER_NAME}'")

    try:
        # Rename the folder
        results = acc.RenameAttributeFolders([rename_params])

        if len(results) > 0:
            result = results[0]

            # Check if rename was successful
            if hasattr(result, 'success') and result.success:
                print(f"\n✓ Successfully renamed folder")
                print(f"  New name: {NEW_FOLDER_NAME}")

            elif hasattr(result, 'error') and result.error:
                # Rename failed
                error_msg = result.error.message if hasattr(
                    result.error, 'message') else str(result.error)
                print(f"\n✗ Error renaming folder:")
                print(f"  {error_msg}")

                print("\nCommon reasons for failure:")
                print("  • Name already exists in parent folder")
                print("  • Name contains invalid characters")
                print("  • Folder is locked or protected")
                print("  • Insufficient permissions")
                exit()
            else:
                print(f"\n⚠️  Unknown rename result")
        else:
            print(f"\n✗ No response received from rename operation")
            exit()

    except Exception as e:
        print(f"\n✗ Error renaming folder: {e}")
        import traceback
        traceback.print_exc()
        exit()

    # -------------------------------------------------------------------------
    # 6. VERIFY RENAME
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VERIFYING RENAME")
    print("="*80)

    try:
        # Get updated structure
        updated_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)

        # Build new path with renamed folder
        new_path = parent_path + \
            [NEW_FOLDER_NAME] if parent_path else [NEW_FOLDER_NAME]

        # Try to find folder with new name
        verify_id, verify_name, verify_parent = find_folder_in_structure(
            updated_structure,
            new_path
        )

        if verify_id and verify_name == NEW_FOLDER_NAME:
            print(f"\n✓ Rename verified successfully")
            print(f"  New path: {' > '.join(new_path)}")
        else:
            print(f"\n⚠️  Could not verify new name")
            print("  (Folder may still have been renamed)")

    except Exception as e:
        print(f"\n⚠️  Could not verify rename: {e}")
        print("  (Folder may still have been renamed)")

    # -------------------------------------------------------------------------
    # SUMMARY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Attribute type: {ATTRIBUTE_TYPE}")
    print(f"  Old name: {current_name}")
    print(f"  New name: {NEW_FOLDER_NAME}")
    if parent_path:
        print(f"  Location: {' > '.join(parent_path)}")
    else:
        print(f"  Location: Root level")
    print("\n" + "="*80)

    print("\n💡 Next steps:")
    print("   • Verify in Archicad Attribute Manager")
    print("   • Check that references are still valid")
    print("   • Rename additional folders if needed")

    print("\n💡 Tips:")
    print("   • Use consistent naming conventions")
    print("   • Keep names clear and descriptive")
    print("   • Avoid special characters")


if __name__ == "__main__":
    main()
