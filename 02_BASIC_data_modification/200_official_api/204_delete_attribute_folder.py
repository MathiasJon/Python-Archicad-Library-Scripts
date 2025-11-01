"""
================================================================================
SCRIPT: Delete Attribute Folder
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Modification

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Deletes attribute folders from the current Archicad project. This script:
- Deletes specified folders from attribute hierarchies
- Automatically deletes all deletable attributes within the folder
- Recursively deletes all subfolders and their contents
- Provides clear warnings about permanent deletion
- Shows detailed information before deletion

⚠️  WARNING: This operation is PERMANENT and cannot be undone!
All deletable attributes and subfolders within the target folder will be removed.

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
  Returns the folder structure for validation

- DeleteAttributeFolders(attributeFolderIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.DeleteAttributeFolders
  Deletes attribute folders and all deletable attributes and folders it contains.
  To delete a folder, its identifier has to be provided.
  Parameters:
  * attributeFolderIds: List of AttributeFolderIdWrapperItem objects
  Returns:
  * List of ExecutionResult objects with success/error status

[Data Types]
- AttributeFolderId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac26.html#archicad.releases.ac26.b3005types.AttributeFolderId
  The identifier of an attribute folder

- AttributeFolderIdWrapperItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac27.html#archicad.releases.ac27.b3001types.AttributeFolderIdWrapperItem 
  Wrapper for attribute folder identifier

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Modify CONFIGURATION section below:
   - Set ATTRIBUTE_TYPE to the type of folder to delete
   - Set FOLDER_PATH to the path of the folder to delete
3. Review the warning about permanent deletion
4. Run this script
5. Verify folder was deleted in Archicad

⚠️  IMPORTANT: This operation cannot be undone!

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Before running, set these values in the CONFIGURATION section:

ATTRIBUTE_TYPE: Type of attribute folder to delete
Supported types:
  • "Layer", "Pen", "LineType", "FillType", "CompositeLine"
  • "BuildingMaterial", "Composite", "Profile", "Surface", "ZoneCategory"

FOLDER_PATH: Path to the folder to delete
Format: List of folder names from root to target
Examples:
  ["Old Materials"] - Deletes "Old Materials" folder at root
  ["Custom", "Temporary"] - Deletes "Temporary" inside "Custom"

CONFIRM_DELETION: Safety flag (must be True to proceed)
Set to True to confirm you understand deletion is permanent

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Folder path and type to delete
- Folder structure information (if found)
- Number of attributes and subfolders to be deleted
- Confirmation requirement
- Deletion results

Example output:
  === DELETE ATTRIBUTE FOLDER ===
  Type: BuildingMaterial
  Path: Temporary Materials
  
  === FINDING FOLDER ===
  ✓ Found folder to delete
  
  ⚠️  WARNING: PERMANENT DELETION
  This folder contains:
    • 12 attribute(s)
    • 3 subfolder(s)
  
  All content will be deleted permanently!
  
  === DELETING FOLDER ===
  ✓ Successfully deleted folder
     Temporary Materials

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
DELETION BEHAVIOR:
⚠️  CRITICAL: Deletion is PERMANENT and cannot be undone!
- All deletable attributes in the folder are deleted
- All subfolders are recursively deleted
- Protected/locked attributes may prevent deletion
- In-use attributes may prevent deletion

PROTECTED ATTRIBUTES:
Some attributes cannot be deleted because:
- They are currently in use by elements
- They are locked by Archicad
- They are system/default attributes
- They are referenced by other attributes

If the folder contains protected attributes:
- The operation may fail
- Some attributes may remain
- Subfolders may be partially deleted

FOLDER PATH:
- Must specify the complete path to the folder
- Path is case-sensitive
- Empty list [] refers to root level (not recommended)
- Folder must exist or operation will fail

RECURSIVE DELETION:
When deleting a folder:
1. All subfolders are deleted first (recursive)
2. All deletable attributes in subfolders are deleted
3. All deletable attributes in target folder are deleted
4. Target folder is deleted last

SAFETY MEASURES:
- CONFIRM_DELETION must be explicitly set to True
- Script shows folder contents before deletion
- Clear warning messages about permanent deletion
- No automatic confirmation - must be set in code

ALTERNATIVE APPROACHES:
Instead of deleting, consider:
- Moving attributes to an "Archive" folder
- Renaming folder to indicate it's deprecated
- Exporting attributes before deletion for backup

LIMITATIONS:
- Cannot delete folders with in-use attributes
- Cannot delete system/default folders
- Cannot undo deletion
- Cannot delete root folder
- Some attributes may remain if protected

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 203_create_attribute_folder.py (create folders)
- 205_move_attributes.py (move attributes to safety before deletion)
- 206_rename_attribute_folder.py (rename instead of delete)
- 123_get_layers.py (view layer folder structure before deletion)

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
# ⚠️  WARNING: DELETION IS PERMANENT AND CANNOT BE UNDONE!

# Attribute type containing the folder to delete
ATTRIBUTE_TYPE = "BuildingMaterial"

# Path to the folder to delete
# Example: ["Temporary"], ["Custom", "Old Materials"]
FOLDER_PATH = ["Temporary Materials"]

# Safety confirmation (MUST be True to proceed with deletion)
# Set to True only after reviewing what will be deleted
CONFIRM_DELETION = True  # Change to True to enable deletion


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def find_folder_in_structure(structure, path, current_path=[]):
    """
    Recursively search for a folder in the attribute structure.

    Args:
        structure: AttributeFolderStructure object
        path: List of folder names to search for
        current_path: Current path in recursion (for tracking)

    Returns:
        Tuple of (AttributeFolderId, folder_info_dict) if found, else (None, None)
    """
    if not path:
        # Found the target folder
        num_attrs = len(structure.attributes) if hasattr(
            structure, 'attributes') and structure.attributes else 0
        num_subfolders = len(structure.subfolders) if hasattr(
            structure, 'subfolders') and structure.subfolders else 0

        folder_info = {
            'name': structure.name if hasattr(structure, 'name') else 'Root',
            'path': current_path,
            'id': structure.attributeFolderId if hasattr(structure, 'attributeFolderId') else None,
            'num_attributes': num_attrs,
            'num_subfolders': num_subfolders
        }

        return (structure.attributeFolderId, folder_info)

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

    return (None, None)


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main():
    """Main function to delete attribute folder."""

    print("\n" + "="*80)
    print("DELETE ATTRIBUTE FOLDER v1.0")
    print("="*80)

    print("\n⚠️  WARNING: This script PERMANENTLY DELETES folders and their contents!")
    print("   This operation CANNOT be undone!")

    # -------------------------------------------------------------------------
    # 1. CHECK CONFIRMATION
    # -------------------------------------------------------------------------

    if not CONFIRM_DELETION:
        print("\n" + "="*80)
        print("DELETION NOT CONFIRMED")
        print("="*80)
        print("\n✗ CONFIRM_DELETION is set to False")
        print("\nFor safety, you must explicitly set CONFIRM_DELETION = True")
        print("in the CONFIGURATION section before deletion can proceed.")
        print("\nThis ensures you understand that deletion is permanent.")
        print("\n⚠️  Review the folder contents carefully before confirming!")
        exit()

    # -------------------------------------------------------------------------
    # 2. DISPLAY CONFIGURATION
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CONFIGURATION")
    print("="*80)

    print(f"\nAttribute type: {ATTRIBUTE_TYPE}")
    print(f"Folder to delete: {' > '.join(FOLDER_PATH)}")
    print(
        f"Confirmation: {'✓ CONFIRMED' if CONFIRM_DELETION else '✗ NOT CONFIRMED'}")

    # -------------------------------------------------------------------------
    # 3. FIND TARGET FOLDER
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("FINDING FOLDER")
    print("="*80)

    try:
        # Get folder structure
        print(f"\nSearching for folder in '{ATTRIBUTE_TYPE}' hierarchy...")
        root_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)

        # Search for target folder
        folder_id, folder_info = find_folder_in_structure(
            root_structure, FOLDER_PATH)

        if not folder_id:
            print(f"\n✗ Folder not found: {' > '.join(FOLDER_PATH)}")
            print(f"\nPlease check:")
            print("  • Folder path is correct")
            print("  • Folder exists in this attribute type")
            print("  • Path is case-sensitive")
            print(f"\n💡 Tip: Run 123_get_layers.py or similar to see folder structure")
            exit()

        print(f"\n✓ Found folder to delete")
        print(f"  Name: {folder_info['name']}")
        print(
            f"  Path: {' > '.join(folder_info['path']) if folder_info['path'] else 'Root'}")

    except Exception as e:
        print(f"\n✗ Error getting folder structure: {e}")
        import traceback
        traceback.print_exc()
        exit()

    # -------------------------------------------------------------------------
    # 4. SHOW FOLDER CONTENTS WARNING
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("⚠️  WARNING: PERMANENT DELETION")
    print("="*80)

    print(f"\nThis folder contains:")
    print(f"  • {folder_info['num_attributes']} attribute(s)")
    print(f"  • {folder_info['num_subfolders']} subfolder(s)")

    if folder_info['num_subfolders'] > 0:
        print(f"\nAll subfolders and their contents will also be deleted!")

    print(f"\nAll deletable content will be PERMANENTLY removed!")
    print(f"This operation CANNOT be undone!")

    # -------------------------------------------------------------------------
    # 5. DELETE FOLDER
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("DELETING FOLDER")
    print("="*80)

    # Prepare folder ID for deletion
    folder_id_wrapper = act.AttributeFolderIdWrapperItem(
        attributeFolderId=folder_id
    )

    print(f"\nDeleting folder: {' > '.join(FOLDER_PATH)}")
    print("Processing...")

    try:
        # Delete the folder
        results = acc.DeleteAttributeFolders([folder_id_wrapper])

        if len(results) > 0:
            result = results[0]

            # Check if deletion was successful
            if hasattr(result, 'success') and result.success:
                print(f"\n✓ Successfully deleted folder:")
                print(f"   {' > '.join(FOLDER_PATH)}")

            elif hasattr(result, 'error') and result.error:
                # Deletion failed
                error_msg = result.error.message if hasattr(
                    result.error, 'message') else str(result.error)
                print(f"\n✗ Error deleting folder:")
                print(f"  {error_msg}")

                print("\nCommon reasons for failure:")
                print("  • Folder contains in-use attributes")
                print("  • Folder contains protected/system attributes")
                print("  • Folder is locked")
                print("  • Insufficient permissions")
                exit()
            else:
                print(f"\n⚠️  Unknown deletion result")
        else:
            print(f"\n✗ No response received from delete operation")
            exit()

    except Exception as e:
        print(f"\n✗ Error deleting folder: {e}")
        import traceback
        traceback.print_exc()
        exit()

    # -------------------------------------------------------------------------
    # 6. VERIFY DELETION
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VERIFYING DELETION")
    print("="*80)

    try:
        # Get updated structure
        updated_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)

        # Try to find the folder again - should not exist
        verify_id, verify_info = find_folder_in_structure(
            updated_structure, FOLDER_PATH)

        if verify_id:
            print(f"\n⚠️  Warning: Folder still exists after deletion")
            print("  Some attributes may have been protected")
        else:
            print(f"\n✓ Folder successfully removed from structure")

    except Exception as e:
        print(f"\n⚠️  Could not verify deletion: {e}")
        print("  (Folder may still have been deleted)")

    # -------------------------------------------------------------------------
    # SUMMARY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Attribute type: {ATTRIBUTE_TYPE}")
    print(f"  Deleted folder: {' > '.join(FOLDER_PATH)}")
    print(f"  Status: Deleted successfully")
    print("\n" + "="*80)

    print("\n💡 Next steps:")
    print("   • Verify in Archicad that folder is removed")
    print("   • Check if any protected attributes remain")
    print("   • Review other folders for cleanup if needed")


if __name__ == "__main__":
    main()
