"""
================================================================================
SCRIPT: Set Element Classification
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Modification

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Sets classification for selected elements in the current Archicad project.
This script provides powerful classification assignment with:
- Recursive search through entire classification hierarchy
- Search by classification ID OR name (case-insensitive)
- Support for multiple selected elements of different types
- Automatic verification of applied classifications
- Detailed progress reporting

Classification systems organize building elements according to standards
like Uniclass, Omniclass, or custom systems. This script simplifies the
process of assigning classifications to multiple elements at once.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Elements must be selected in Archicad before running

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements
  Returns the list of currently selected elements

- GetAllClassificationSystems()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationSystems
  Returns all classification systems in the project

- GetAllClassificationsInSystem(classificationSystemId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationsInSystem
  Returns all classification items in a specific system

- SetClassificationsOfElements(elementClassifications)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.SetClassificationsOfElements
  Sets the classifications of elements

- GetClassificationsOfElements(elements, classificationSystemIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetClassificationsOfElements
  Returns the classifications of elements

- GetDetailsOfClassificationItems(classificationItemIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfClassificationItems
  Returns detailed information about classification items

- GetTypesOfElements(elements)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements
  Returns the type of elements

[Data Types]
- ClassificationId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationId
  Combines classification system ID and classification item ID

- ElementClassification
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementClassification 
  Contains element ID and its classification ID

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select elements you want to classify
3. Modify CONFIGURATION section below:
   - Set CLASSIFICATION_SYSTEM to your system name
   - Set CLASSIFICATION_SEARCH to classification ID or name
4. Run this script
5. Verify the classifications were applied

To find available systems and classifications:
- Run 101_get_all_classification_systems.py
- Run 102_get_classification_system_with_hierarchy.py

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Before running, set these values in the CONFIGURATION section:

CLASSIFICATION_SYSTEM: Name of the classification system to use
Example: "Uniclass 2015", "Omniclass", "Classification Archicad"

CLASSIFICATION_SEARCH: ID or name of the classification to apply
- Can be the classification ID (e.g., "Ss_25_10_20")
- Can be the classification name (e.g., "Wall")
- Search is case-insensitive
- Searches recursively through all hierarchy levels

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Number of selected elements
- Search progress through hierarchy
- Found classification details
- Element types breakdown
- Application progress
- Verification results

Example output:
  === SETTING CLASSIFICATION ===
  Elements selected: 15
  ✓ Found classification system: Uniclass 2015
  
  === SEARCHING THROUGH HIERARCHY ===
  ✓ FOUND: 'Ss_25_10_20' | 'External walls'
  
  === ELEMENTS TO CLASSIFY ===
    Wall: 12
    Column: 3
  
  === APPLYING CLASSIFICATION ===
  ✓ Successfully set classification for 15 element(s)
  
  === VERIFYING CHANGES ===
  ✓ Verification successful

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

If you have multiple Archicad instances open:
- The script connects to the first one opened
- Make sure the correct project is open in that instance
- Select elements in the correct instance before running

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
CLASSIFICATION HIERARCHY:
- Classifications are organized in tree structures
- This script searches recursively through all levels
- Parent classifications can contain child classifications
- Search finds matches at any depth level

SEARCH BEHAVIOR:
- Case-insensitive matching for both ID and name
- Searches through ID field first, then name field
- Returns first match found in hierarchy
- Stops searching after first match

ELEMENT SELECTION:
- Elements must be selected before running the script
- Can select multiple element types simultaneously
- All selected elements receive the same classification
- Script shows breakdown by element type

VERIFICATION:
- Script automatically verifies the applied classification
- Checks first element in selection as sample
- Verification failure doesn't mean operation failed
- Manual verification recommended for critical operations

CLASSIFICATION STRUCTURE:
- ClassificationId combines system ID and item ID
- Both IDs required for proper classification assignment
- Same classification can exist in multiple systems
- System must match the intended classification standard

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 101_get_all_classification_systems.py (list available systems)
- 102_get_classification_system_with_hierarchy.py (browse hierarchy)
- 107_get_elements_by_classification.py (find classified elements)
- 111_get_element_full_info.py (verify element classifications)

================================================================================
"""

from archicad import ACConnection

# =============================================================================
# CONNECT TO ARCHICAD
# =============================================================================

# Establish connection to running Archicad instance
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

# Initialize command, type, and utility objects
acc = conn.commands
act = conn.types
acu = conn.utilities


# =============================================================================
# CONFIGURATION
# =============================================================================
# IMPORTANT: Modify these values before running the script!
# To find available systems and classifications, run:
# - 101_get_all_classification_systems.py
# - 102_get_classification_system_with_hierarchy.py

# Name of the classification system to use
CLASSIFICATION_SYSTEM = "Classification Archicad"

# Classification to search for (can be ID or name, case-insensitive)
# Examples:
#   - By ID: "Ss_25_10_20"
#   - By name: "External walls"
#   - By name: "Cloison" (partition in French)
CLASSIFICATION_SEARCH = "Cloison"


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def search_classification_recursive(item, search_term, level=0):
    """
    Recursively search through classification hierarchy.

    Searches both ID and name fields (case-insensitive) at all hierarchy levels.
    Returns the first match found.

    Args:
        item: ClassificationItemArrayItem object to search
        search_term: String to search for (ID or name)
        level: Current depth level in hierarchy (for display)

    Returns:
        Tuple of (classificationItemId, name) if found, else (None, None)
    """
    # Get the classification item from wrapper
    class_item = item.classificationItem

    # Extract ID and name
    item_id = class_item.id
    item_name = class_item.name

    # Convert search term to lowercase for case-insensitive comparison
    search_lower = search_term.lower()

    # Check if this item matches the search term
    if (item_id.lower() == search_lower or
            item_name.lower() == search_lower):
        # Use ID as fallback name if name is empty
        display_name = item_name if item_name else item_id

        # Display match with indentation showing hierarchy level
        indent = "  " * level
        print(f"{indent}✓ FOUND: '{item_id}' | '{display_name}'")

        return (class_item.classificationItemId, display_name)

    # Recursively search children if they exist
    if hasattr(class_item, 'children') and class_item.children:
        for child in class_item.children:
            result = search_classification_recursive(
                child, search_term, level + 1)
            if result[0] is not None:
                return result

    # No match found at this level or below
    return (None, None)


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main():
    """Main function to set element classifications."""

    print("\n" + "="*80)
    print("SET ELEMENT CLASSIFICATION v1.0")
    print("="*80)

    # -------------------------------------------------------------------------
    # 1. GET SELECTED ELEMENTS
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SELECTED ELEMENTS")
    print("="*80)

    selected_elements = acc.GetSelectedElements()

    if len(selected_elements) == 0:
        print("\n⚠️  No elements selected")
        print("\nPlease:")
        print("  1. Select elements in Archicad")
        print("  2. Run this script again")
        exit()

    print(f"\n✓ Found {len(selected_elements)} selected element(s)")
    print(f"  Searching for classification: '{CLASSIFICATION_SEARCH}'")

    # Get element types for information
    element_types = acc.GetTypesOfElements(
        [elem.elementId for elem in selected_elements]
    )

    # Count elements by type
    type_counts = {}
    for elem_type in element_types:
        type_name = elem_type.typeOfElement.elementType
        type_counts[type_name] = type_counts.get(type_name, 0) + 1

    print("\n  Element types:")
    for elem_type, count in sorted(type_counts.items()):
        print(f"    • {elem_type}: {count}")

    # -------------------------------------------------------------------------
    # 2. FIND CLASSIFICATION SYSTEM
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CLASSIFICATION SYSTEM")
    print("="*80)

    # Get all classification systems in project
    classification_systems = acc.GetAllClassificationSystems()

    # Search for the specified system
    target_system_guid = None
    for system in classification_systems:
        if system.name == CLASSIFICATION_SYSTEM:
            target_system_guid = system.classificationSystemId.guid
            break

    # Check if system was found
    if not target_system_guid:
        print(f"\n✗ Classification system '{CLASSIFICATION_SYSTEM}' not found")
        print("\nAvailable systems:")
        for system in classification_systems:
            print(f"  • {system.name}")
        print("\n💡 Tip: Run 101_get_all_classification_systems.py to see all systems")
        exit()

    print(f"\n✓ Found classification system: {CLASSIFICATION_SYSTEM}")

    # -------------------------------------------------------------------------
    # 3. SEARCH CLASSIFICATION HIERARCHY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SEARCHING CLASSIFICATION HIERARCHY")
    print("="*80)

    # Get all root classifications in the system
    all_classifications = acc.GetAllClassificationsInSystem(
        act.ClassificationSystemId(target_system_guid)
    )

    print(f"\nSearching for: '{CLASSIFICATION_SEARCH}'")
    print(f"(Case-insensitive, searches ID and name fields)\n")

    # Search through all root classifications and their children
    target_classification = None
    classification_name = None

    for root_item in all_classifications:
        result = search_classification_recursive(
            root_item,
            CLASSIFICATION_SEARCH,
            level=0
        )
        if result[0] is not None:
            target_classification, classification_name = result
            break

    # Check if classification was found
    if not target_classification:
        print(
            f"\n✗ Classification '{CLASSIFICATION_SEARCH}' not found in hierarchy")
        print("\nPossible reasons:")
        print("  • Classification ID or name is incorrect")
        print("  • Classification doesn't exist in this system")
        print("  • Typo in search term")
        print("\n💡 Tip: Run 102_get_classification_system_with_hierarchy.py")
        print("        to see all available classifications")
        exit()

    print(f"\n✓ Using classification: {classification_name}")

    # -------------------------------------------------------------------------
    # 4. APPLY CLASSIFICATION
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("APPLYING CLASSIFICATION")
    print("="*80)

    # Prepare classification data for each element
    element_classifications = []

    for element in selected_elements:
        # Create ClassificationId combining system and item
        classification_id = act.ClassificationId(
            classificationSystemId=act.ClassificationSystemId(
                target_system_guid),
            classificationItemId=target_classification
        )

        # Create ElementClassification object
        element_classifications.append(
            act.ElementClassification(
                elementId=element.elementId,
                classificationId=classification_id
            )
        )

    # Apply classifications to all selected elements
    print(f"\nApplying classification '{classification_name}'...")
    print(f"  To {len(selected_elements)} element(s)...")

    try:
        acc.SetClassificationsOfElements(element_classifications)
        print(
            f"\n✓ Successfully set classification for {len(selected_elements)} element(s)")

    except Exception as e:
        print(f"\n✗ Error setting classification: {e}")
        import traceback
        traceback.print_exc()
        exit()

    # -------------------------------------------------------------------------
    # 5. VERIFY CHANGES
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VERIFYING CHANGES")
    print("="*80)

    try:
        # Verify by checking the first element as a sample
        classifications = acc.GetClassificationsOfElements(
            [selected_elements[0].elementId],
            [act.ClassificationSystemId(target_system_guid)]
        )

        if len(classifications) > 0 and len(classifications[0].classificationIds) > 0:
            print(f"\n✓ Verification successful - classification applied")

            # Show the applied classification details
            for class_id in classifications[0].classificationIds:
                class_details = acc.GetDetailsOfClassificationItems(
                    [class_id.classificationItemId]
                )
                if len(class_details) > 0:
                    class_item = class_details[0].classificationItem
                    print(f"\n  Applied classification:")
                    print(f"    • ID: {class_item.id}")
                    print(f"    • Name: {class_item.name}")
        else:
            print(f"\n⚠️  Could not verify classification")
            print("  (Classification may still have been applied successfully)")

    except Exception as e:
        print(f"\n⚠️  Verification failed: {e}")
        print("  (Classification may still have been applied successfully)")

    # -------------------------------------------------------------------------
    # SUMMARY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Classification: {classification_name}")
    print(f"  System: {CLASSIFICATION_SYSTEM}")
    print(f"  Elements modified: {len(selected_elements)}")
    print("\n" + "="*80)

    # Usage hints
    print("\n💡 Next steps:")
    print("   • Use 107_get_elements_by_classification.py to find classified elements")
    print("   • Use 111_get_element_full_info.py to verify element details")
    print("   • Use Element Info Palette in Archicad to check classifications")


if __name__ == "__main__":
    main()
