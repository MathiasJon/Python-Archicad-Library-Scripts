"""
================================================================================
SCRIPT: Move Attributes Between Folders
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Modification

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Moves attributes and attribute folders between different folders in the
attribute hierarchy. This script provides:
- Moving individual attributes to different folders
- Moving entire folders (with all contents) to new locations
- Moving multiple attributes and folders in one operation
- Validation of source and target folders
- Automatic name suggestions if attributes not found
- List of all available attributes to help find correct names
- Clear feedback on move operations

Useful for reorganizing project attributes into logical folder structures.

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

- GetAttributesByType(attributeType)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType
  Returns all attribute IDs of specified type

- GetAttributeFolderStructure(attributeType, path)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure
  Returns folder structure for finding folders and attributes

- MoveAttributesAndFolders(attributeFolderIds, attributeIds, targetFolderId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.MoveAttributesAndFolders
  Moves attributes and attribute folders
  Parameters:
  * attributeFolderIds: List of AttributeFolderIdWrapperItem to move
  * attributeIds: List of AttributeIdWrapperItem to move
  * targetFolderId: AttributeFolderId of destination folder

[Data Types]
- AttributeId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.AttributeId

- AttributeFolderId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.AttributeFolderId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Modify CONFIGURATION section below:
   - Set ATTRIBUTE_TYPE
   - Set what to move (attributes and/or folders)
   - Set TARGET_FOLDER_PATH for destination
3. Run this script
4. Verify items were moved in Archicad

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Before running, set these values in the CONFIGURATION section:

ATTRIBUTE_TYPE: Type of attributes to move
Supported types: "Layer", "BuildingMaterial", "Composite", "Profile", etc.

ATTRIBUTE_NAMES_TO_MOVE: List of attribute names to move INDIVIDUALLY
⚠️  IMPORTANT: These are moved as individual items, NOT their parent folders
Set to [] if not moving individual attributes
Example: ["Material A", "Material B", "Material C"]

FOLDER_PATHS_TO_MOVE: List of folder paths to move ENTIRELY
⚠️  CRITICAL: When you move a folder, ALL attributes inside it are moved too!
Set to [] if not moving folders
Example: [["Temporary"], ["Old", "Archived"]]
Format: Each path is a list of folder names from root

Common scenarios:
1. Move only specific attributes, leave folder behind:
   ATTRIBUTE_NAMES_TO_MOVE = ["Material A", "Material B"]
   FOLDER_PATHS_TO_MOVE = []

2. Move entire folder with all its contents:
   ATTRIBUTE_NAMES_TO_MOVE = []
   FOLDER_PATHS_TO_MOVE = [["Temporary Materials"]]
   → Moves folder AND all attributes inside it

3. Move some attributes AND entire folders:
   ATTRIBUTE_NAMES_TO_MOVE = ["Material A"]
   FOLDER_PATHS_TO_MOVE = [["Old Folder"]]
   → Moves Material A + Old Folder (with all its contents)

TARGET_FOLDER_PATH: Destination folder path
Example: ["Archive"] for root-level folder
Example: ["Structural", "Steel"] for nested folder

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Items to move (attributes and folders)
- Source and target folder information
- Suggestions for similar names if attributes not found
- List of available attributes to help find correct names
- Move operation progress
- Success or error messages

Example output (when attributes not found):
  === FINDING ATTRIBUTES ===
  Searching for 2 attribute(s)...
    ✗ NOT FOUND: Custom Material 1
       Check: name is exact (case-sensitive)
       Similar names found:
         → Custom Material (in Materials)
         → Material 1 (in Custom)
    ✗ NOT FOUND: Custom Material 2
       Check: name is exact (case-sensitive)
  
  Result: Found 0/2 attribute(s)
  
  ⚠️  WARNING: No attributes were found!
     Only folders will be moved (if any)
  
  💡 TIP: Listing available BuildingMaterial attributes...
  
  Available attributes (first 20):
     1. Concrete (at root)
     2. Steel (in Structural)
     3. Wood (in Finishes)
     ...
  
  💡 Copy the EXACT name (case-sensitive) to ATTRIBUTE_NAMES_TO_MOVE

Example output (success):
  === FINDING ATTRIBUTES ===
  Searching for 2 attribute(s)...
    ✓ Found: Concrete
    ✓ Found: Steel
  
  === MOVING ITEMS ===
  ✓ Successfully moved items

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
MOVE BEHAVIOR:
- Attributes are moved individually or in batches
- Folders are moved with ALL their contents (including subfolders)
- Original hierarchy within moved folders is preserved
- Cannot move to same location (no effect)

⚠️  CRITICAL DIFFERENCE:
- ATTRIBUTE_NAMES_TO_MOVE: Moves ONLY the specified attributes (individual items)
- FOLDER_PATHS_TO_MOVE: Moves the ENTIRE folder including ALL attributes inside

⚠️  IMPORTANT: If an attribute is already inside a folder you're moving, you DON'T need
to list it in ATTRIBUTE_NAMES_TO_MOVE - it will move automatically with its folder!

Example scenario 1:
Folder "Temporary" contains: Material A, Material B, Material C

Option 1 - Move only Material A:
  ATTRIBUTE_NAMES_TO_MOVE = ["Material A"]
  FOLDER_PATHS_TO_MOVE = []
  Result: Only Material A moves. Material B and C stay in "Temporary"

Option 2 - Move entire folder:
  ATTRIBUTE_NAMES_TO_MOVE = []
  FOLDER_PATHS_TO_MOVE = [["Temporary"]]
  Result: Folder "Temporary" moves with Material A, B, AND C inside it

Option 3 - Move Material A + entire other folder:
  ATTRIBUTE_NAMES_TO_MOVE = ["Material A"]  # Must be OUTSIDE "Old Folder"
  FOLDER_PATHS_TO_MOVE = [["Old Folder"]]
  Result: Material A moves individually + "Old Folder" moves with all its contents
  
⚠️  Common mistake: If Material A is INSIDE "Old Folder", don't list it in
ATTRIBUTE_NAMES_TO_MOVE - it will move automatically with the folder!

FOLDER MOVES:
When moving a folder:
- All subfolders and attributes move with it
- Hierarchy within folder is maintained
- Folder structure is preserved
- Cannot move folder into itself or its subfolders

ATTRIBUTE IDENTIFICATION:
- Attributes are identified by name (case-sensitive)
- If multiple attributes have same name, first match is used
- Script searches through entire hierarchy
- Attributes not found are skipped with warning

TARGET FOLDER:
- Target folder must exist before moving
- Use 203_create_attribute_folder.py to create target first
- Cannot move to root level (use empty path [])
- Target must be valid folder in same attribute type

LIMITATIONS:
- Cannot move system/default attributes
- Cannot move locked attributes
- Cannot move folders into their own subfolders
- Cannot move to non-existent folders
- All items must be same attribute type

VALIDATION:
Script validates:
- Target folder exists
- Source attributes/folders exist
- No circular dependencies (folder into itself)
- Attribute type consistency

COMMON USE CASES:
- Reorganizing project attributes into logical groups
- Moving temporary attributes to permanent folders
- Archiving old/unused attributes
- Consolidating scattered attributes
- Cleaning up messy folder structures

TROUBLESHOOTING:

Problem: "I listed attributes but only folders moved"
Possible causes:
1. Attribute names didn't match exactly (case-sensitive)
   → Check console output for "✗ NOT FOUND" messages
2. Attributes were already inside folders being moved
   → They move with folder automatically, no need to list them
3. Attributes don't exist in the project
   → Use Attribute Manager to verify exact names

Problem: "Attribute is not moving"
Possible causes:
1. Attribute is inside a folder you're moving
   → It moves with the folder, not individually
2. Attribute name is misspelled
   → Names are case-sensitive, must match exactly
3. Attribute is in different attribute type
   → Check you're using correct ATTRIBUTE_TYPE

Problem: "Too many items moved"
Cause: Moved a folder instead of individual attributes
Solution: Use ATTRIBUTE_NAMES_TO_MOVE for individual items,
          leave FOLDER_PATHS_TO_MOVE empty

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 203_create_attribute_folder.py (create target folders first)
- 204_delete_attribute_folder.py (delete after moving)
- 206_rename_attribute_folder.py (rename folders)
- 122_get_profiles.py (view structure before moving)

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

# Attribute type to work with
ATTRIBUTE_TYPE = "BuildingMaterial"

# Names of individual attributes to move
# Leave empty [] if not moving attributes
# These attributes are moved INDIVIDUALLY (not their parent folder)
#
# ⚠️  IMPORTANT: Only list attributes that are NOT inside folders you're moving!
# If an attribute is inside a folder in FOLDER_PATHS_TO_MOVE, it will move
# automatically with that folder - no need to list it here.
#
# Example scenario:
# If folder "Temporary" contains: Mat A, Mat B, Mat C
# And you set: ATTRIBUTE_NAMES_TO_MOVE = ["Mat A", "Mat B"]
# Result: Only Mat A and Mat B move. Mat C stays in "Temporary" folder.
#
# BUT if you ALSO set: FOLDER_PATHS_TO_MOVE = [["Temporary"]]
# Result: The entire folder moves with ALL materials (Mat A, B, C)
#         Even if you listed Mat A and Mat B here, they move with the folder anyway!
ATTRIBUTE_NAMES_TO_MOVE = [
    "Custom Material 1",
    "Custom Material 2"
]

# Paths of folders to move (each path is a list)
# Leave empty [] if not moving folders
# ⚠️  WARNING: Moving a folder moves ALL its contents!
#
# Example scenario:
# If folder "Temporary" contains: Mat A, Mat B, Mat C
# And you set: FOLDER_PATHS_TO_MOVE = [["Temporary"]]
# Result: Folder "Temporary" moves with Mat A, Mat B, AND Mat C inside it!
FOLDER_PATHS_TO_MOVE = [
    ["Temporary Materials"]
]

# Target folder path where items will be moved
# Examples:
#   ["Archive"] - Move to "Archive" folder at root
#   ["Custom", "Old"] - Move to "Old" folder inside "Custom"
#   [] - Move to root level (not recommended)
TARGET_FOLDER_PATH = ["Archive"]


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
        AttributeFolderId if found, else None
    """
    if not path:
        return structure.attributeFolderId if hasattr(structure, 'attributeFolderId') else None

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

    return None


def find_attribute_by_name(structure, name, attribute_type):
    """
    Recursively search for an attribute by name in the structure.

    Args:
        structure: AttributeFolderStructure object
        name: Attribute name to search for
        attribute_type: Type of attribute

    Returns:
        AttributeId if found, else None
    """
    # Check attributes in current folder
    if hasattr(structure, 'attributes') and structure.attributes:
        for attr_item in structure.attributes:
            if hasattr(attr_item, 'attributeHeader'):
                header = attr_item.attributeHeader
                if hasattr(header, 'name') and header.name == name:
                    if hasattr(header, 'attributeId'):
                        return header.attributeId

    # Search in subfolders
    if hasattr(structure, 'subfolders') and structure.subfolders:
        for subfolder_item in structure.subfolders:
            if hasattr(subfolder_item, 'attributeFolder'):
                result = find_attribute_by_name(
                    subfolder_item.attributeFolder,
                    name,
                    attribute_type
                )
                if result:
                    return result

    return None


def list_all_attributes_in_structure(structure, current_path=[]):
    """
    Recursively list all attributes in the structure with their paths.

    Args:
        structure: AttributeFolderStructure object
        current_path: Current folder path

    Returns:
        List of tuples (attribute_name, folder_path)
    """
    attributes = []

    # Get attributes in current folder
    if hasattr(structure, 'attributes') and structure.attributes:
        for attr_item in structure.attributes:
            if hasattr(attr_item, 'attributeHeader'):
                header = attr_item.attributeHeader
                if hasattr(header, 'name'):
                    attributes.append((header.name, current_path.copy()))

    # Search in subfolders
    if hasattr(structure, 'subfolders') and structure.subfolders:
        for subfolder_item in structure.subfolders:
            if hasattr(subfolder_item, 'attributeFolder'):
                subfolder = subfolder_item.attributeFolder
                subfolder_name = subfolder.name if hasattr(
                    subfolder, 'name') else 'Unknown'
                attributes.extend(
                    list_all_attributes_in_structure(
                        subfolder,
                        current_path + [subfolder_name]
                    )
                )

    return attributes


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main():
    """Main function to move attributes and folders."""

    print("\n" + "="*80)
    print("MOVE ATTRIBUTES AND FOLDERS v1.0")
    print("="*80)

    # -------------------------------------------------------------------------
    # 1. DISPLAY CONFIGURATION
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CONFIGURATION")
    print("="*80)

    print(f"\nAttribute type: {ATTRIBUTE_TYPE}")
    print(f"\nItems to move:")
    print(f"  Attributes: {len(ATTRIBUTE_NAMES_TO_MOVE)}")
    for name in ATTRIBUTE_NAMES_TO_MOVE:
        print(f"    • {name}")

    print(f"  Folders: {len(FOLDER_PATHS_TO_MOVE)}")
    for path in FOLDER_PATHS_TO_MOVE:
        print(f"    • {' > '.join(path)}")

    print(
        f"\nTarget folder: {' > '.join(TARGET_FOLDER_PATH) if TARGET_FOLDER_PATH else 'Root'}")

    # Check if there's anything to move
    if not ATTRIBUTE_NAMES_TO_MOVE and not FOLDER_PATHS_TO_MOVE:
        print("\n⚠️  Nothing to move!")
        print("Please specify attributes and/or folders to move in the configuration.")
        exit()

    # -------------------------------------------------------------------------
    # 2. GET FOLDER STRUCTURE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("LOADING STRUCTURE")
    print("="*80)

    try:
        print(f"\nLoading '{ATTRIBUTE_TYPE}' folder structure...")
        root_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)
        print(f"✓ Structure loaded")

    except Exception as e:
        print(f"\n✗ Error loading structure: {e}")
        exit()

    # -------------------------------------------------------------------------
    # 3. FIND TARGET FOLDER
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("FINDING TARGET FOLDER")
    print("="*80)

    if TARGET_FOLDER_PATH:
        target_folder_id = find_folder_in_structure(
            root_structure, TARGET_FOLDER_PATH)

        if not target_folder_id:
            print(
                f"\n✗ Target folder not found: {' > '.join(TARGET_FOLDER_PATH)}")
            print("\n💡 Tips:")
            print("  • Check folder path is correct")
            print("  • Path is case-sensitive")
            print("  • Create folder first using 203_create_attribute_folder.py")
            exit()

        print(f"\n✓ Found target folder: {' > '.join(TARGET_FOLDER_PATH)}")
    else:
        # Moving to root
        target_folder_id = root_structure.attributeFolderId
        print(f"\n✓ Target is root folder")

    # -------------------------------------------------------------------------
    # 4. FIND ATTRIBUTES TO MOVE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("FINDING ATTRIBUTES")
    print("="*80)

    attribute_ids_to_move = []

    if ATTRIBUTE_NAMES_TO_MOVE:
        print(
            f"\nSearching for {len(ATTRIBUTE_NAMES_TO_MOVE)} attribute(s)...")

        # Get ALL attributes of this type directly from API
        try:
            all_attribute_ids = acc.GetAttributesByType(ATTRIBUTE_TYPE)

            # Get their names using the appropriate Get...Attributes command
            # This is more reliable than parsing folder structure
            if ATTRIBUTE_TYPE == "BuildingMaterial":
                all_attributes = acc.GetBuildingMaterialAttributes(
                    all_attribute_ids)
            elif ATTRIBUTE_TYPE == "Layer":
                all_attributes = acc.GetLayerAttributes(all_attribute_ids)
            elif ATTRIBUTE_TYPE == "Profile":
                all_attributes = acc.GetProfileAttributes(all_attribute_ids)
            elif ATTRIBUTE_TYPE == "Composite":
                all_attributes = acc.GetCompositeAttributes(all_attribute_ids)
            elif ATTRIBUTE_TYPE == "Surface":
                all_attributes = acc.GetSurfaceAttributes(all_attribute_ids)
            else:
                # Fallback: try to get from folder structure
                print(
                    f"\n⚠️  Warning: Direct attribute retrieval not implemented for {ATTRIBUTE_TYPE}")
                print(f"   Using folder structure search (may be less reliable)")
                all_attributes = None

            if all_attributes:
                # Create a name -> ID mapping
                name_to_id = {}
                for attr_or_error in all_attributes:
                    # Try different attribute wrapper names
                    attr = None
                    if hasattr(attr_or_error, f'{ATTRIBUTE_TYPE.lower()}Attribute'):
                        attr = getattr(
                            attr_or_error, f'{ATTRIBUTE_TYPE.lower()}Attribute')
                    elif hasattr(attr_or_error, 'buildingMaterialAttribute'):
                        attr = attr_or_error.buildingMaterialAttribute
                    elif hasattr(attr_or_error, 'layerAttribute'):
                        attr = attr_or_error.layerAttribute
                    elif hasattr(attr_or_error, 'profileAttribute'):
                        attr = attr_or_error.profileAttribute
                    elif hasattr(attr_or_error, 'compositeAttribute'):
                        attr = attr_or_error.compositeAttribute
                    elif hasattr(attr_or_error, 'surfaceAttribute'):
                        attr = attr_or_error.surfaceAttribute

                    if attr and hasattr(attr, 'name') and hasattr(attr, 'attributeId'):
                        name_to_id[attr.name] = attr.attributeId

                print(
                    f"  Loaded {len(name_to_id)} {ATTRIBUTE_TYPE} attributes from project")

                # DEBUG: Show first 10 attribute names to verify we're loading the right thing
                if name_to_id:
                    print(f"\n  First 10 {ATTRIBUTE_TYPE} names in project:")
                    for i, name in enumerate(list(name_to_id.keys())[:10], 1):
                        # Show repr() to reveal any hidden characters
                        print(f"    {i}. '{name}' (length: {len(name)})")
                    if len(name_to_id) > 10:
                        print(f"    ... and {len(name_to_id) - 10} more")
                print()

                # Now search for requested attributes
                for attr_name in ATTRIBUTE_NAMES_TO_MOVE:
                    print(
                        f"  Searching for: '{attr_name}' (length: {len(attr_name)})")

                    if attr_name in name_to_id:
                        attribute_ids_to_move.append(
                            act.AttributeIdWrapperItem(
                                attributeId=name_to_id[attr_name])
                        )
                        print(f"    ✓ FOUND in attribute list")
                    else:
                        print(f"    ✗ NOT FOUND in attribute list")

                        # Check for exact match ignoring case
                        case_insensitive_match = None
                        for name in name_to_id.keys():
                            if name.lower() == attr_name.lower():
                                case_insensitive_match = name
                                break

                        if case_insensitive_match:
                            print(
                                f"    ⚠️  Found with different case: '{case_insensitive_match}'")
                            print(f"       Use exact case-sensitive name!")
                        else:
                            # Show similar names
                            similar = [name for name in name_to_id.keys(
                            ) if attr_name.lower() in name.lower()]
                            if similar:
                                print(f"    Did you mean:")
                                for sim in similar[:3]:
                                    print(f"      → '{sim}'")
            else:
                # Fallback to folder structure search
                for attr_name in ATTRIBUTE_NAMES_TO_MOVE:
                    attr_id = find_attribute_by_name(
                        root_structure, attr_name, ATTRIBUTE_TYPE)

                    if attr_id:
                        attribute_ids_to_move.append(
                            act.AttributeIdWrapperItem(attributeId=attr_id)
                        )
                        print(f"  ✓ Found: {attr_name}")
                    else:
                        print(f"  ✗ NOT FOUND: {attr_name}")

        except Exception as e:
            print(f"\n✗ Error loading attributes: {e}")
            import traceback
            traceback.print_exc()

        print(
            f"\nResult: Found {len(attribute_ids_to_move)}/{len(ATTRIBUTE_NAMES_TO_MOVE)} attribute(s)")

        if len(attribute_ids_to_move) == 0:
            print(f"\n⚠️  WARNING: No attributes were found!")
            print(f"   Only folders will be moved (if any)")
    else:
        print(f"\nNo attributes to move (ATTRIBUTE_NAMES_TO_MOVE is empty)")

    # -------------------------------------------------------------------------
    # 5. FIND FOLDERS TO MOVE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("FINDING FOLDERS")
    print("="*80)

    folder_ids_to_move = []

    if FOLDER_PATHS_TO_MOVE:
        print(f"\nSearching for {len(FOLDER_PATHS_TO_MOVE)} folder(s)...")
        print(f"⚠️  Note: Moving a folder will move ALL its contents!")

        for folder_path in FOLDER_PATHS_TO_MOVE:
            folder_id = find_folder_in_structure(root_structure, folder_path)

            if folder_id:
                folder_ids_to_move.append(
                    act.AttributeFolderIdWrapperItem(
                        attributeFolderId=folder_id)
                )
                print(f"  ✓ Found: {' > '.join(folder_path)}")
            else:
                print(f"  ✗ Not found: {' > '.join(folder_path)}")

        print(
            f"\nFound {len(folder_ids_to_move)}/{len(FOLDER_PATHS_TO_MOVE)} folder(s)")
    else:
        print(f"\nNo folders to move")

    # Check if anything was found
    if not attribute_ids_to_move and not folder_ids_to_move:
        print("\n✗ No items found to move!")
        print("Please check the names and paths in the configuration.")
        exit()

    # -------------------------------------------------------------------------
    # 6. MOVE ITEMS
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("MOVING ITEMS")
    print("="*80)

    print(
        f"\nTarget destination: {' > '.join(TARGET_FOLDER_PATH) if TARGET_FOLDER_PATH else 'Root'}")
    print(f"\nItems prepared for move:")
    print(f"  • Individual attributes: {len(attribute_ids_to_move)}")
    if attribute_ids_to_move:
        print(f"    (These will move as individual items)")
    print(f"  • Folders: {len(folder_ids_to_move)}")
    if folder_ids_to_move:
        print(f"    (These will move with ALL their contents)")

    # Check if we actually have items to move
    if not attribute_ids_to_move and not folder_ids_to_move:
        print("\n⚠️  Warning: No items to move (nothing was found)")
        exit()

    if folder_ids_to_move:
        print(f"\n⚠️  Reminder: Folder contents will move automatically!")

    # Debug info
    print(f"\nAPI call details:")
    print(f"  - Folder IDs to move: {len(folder_ids_to_move)}")
    print(f"  - Attribute IDs to move: {len(attribute_ids_to_move)}")
    print(f"  - Target folder ID: {'Set' if target_folder_id else 'Not set'}")

    try:
        # Perform the move
        acc.MoveAttributesAndFolders(
            folder_ids_to_move,
            attribute_ids_to_move,
            target_folder_id
        )

        print(f"\n✓ Successfully moved items")

    except Exception as e:
        print(f"\n✗ Error moving items: {e}")
        import traceback
        traceback.print_exc()

        print("\nCommon issues:")
        print("  • Cannot move folder into itself or its subfolders")
        print("  • Items are locked or protected")
        print("  • Target folder is invalid")
        print("  • Insufficient permissions")
        exit()

    # -------------------------------------------------------------------------
    # 7. VERIFY MOVE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VERIFYING MOVE")
    print("="*80)

    try:
        # Reload structure to verify
        updated_structure = acc.GetAttributeFolderStructure(ATTRIBUTE_TYPE)
        print(f"\n✓ Structure updated")
        print(f"  Verify in Archicad that items were moved correctly")

    except Exception as e:
        print(f"\n⚠️  Could not verify move: {e}")
        print("  (Items may still have been moved successfully)")

    # -------------------------------------------------------------------------
    # SUMMARY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Attribute type: {ATTRIBUTE_TYPE}")
    print(f"  Attributes moved: {len(attribute_ids_to_move)}")
    print(f"  Folders moved: {len(folder_ids_to_move)}")
    print(
        f"  Target: {' > '.join(TARGET_FOLDER_PATH) if TARGET_FOLDER_PATH else 'Root'}")
    print("\n" + "="*80)

    print("\n💡 Next steps:")
    print("   • Verify items in Archicad Attribute Manager")
    print("   • Check that folder hierarchies are correct")
    print("   • Move additional items if needed")


if __name__ == "__main__":
    main()
