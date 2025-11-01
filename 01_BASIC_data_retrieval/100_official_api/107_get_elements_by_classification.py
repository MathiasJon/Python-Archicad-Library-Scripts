"""
================================================================================
SCRIPT: Get Elements by Classification
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves all elements with a specific classification. Searches through the
entire classification hierarchy (including children) by ID or name, with
case-insensitive matching.

This script is useful for finding all elements that have been assigned a 
specific classification, regardless of their element type.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- CLASSIFICATION_SYSTEM      : Name of the classification system (e.g., "Classification Archicad")
- CLASSIFICATION_SEARCH      : Classification ID or name to search for
- MAX_ELEMENTS_TO_DISPLAY    : Maximum number of elements to display with details (default: 10)

The search is case-insensitive and works with both classification IDs and names.

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAllClassificationSystems()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationSystems
  
- GetAllClassificationsInSystem()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationsInSystem
  
- GetElementsByClassification()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetElementsByClassification
  
- GetTypesOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements
  
- GetAllPropertyIdsOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyIdsOfElements
  
- GetPropertyValuesOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetPropertyValuesOfElements

[Data Types]
- ClassificationSystem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationSystem
  
- ClassificationSystemId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationId.classificationSystemId
  
- ClassificationItemArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationItemArrayItem
  
- ElementIdArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementIdArrayItem

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Run 102_get_classification_system_with_hierarchy.py to find classification IDs
2. Modify CLASSIFICATION_SYSTEM and CLASSIFICATION_SEARCH in this script
3. Run this script
4. View elements with the specified classification

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. SEARCH PARAMETERS:
   - Classification system being searched
   - Search term used

2. CLASSIFICATION HIERARCHY SEARCH:
   - Visual feedback when classification is found
   - Total classifications searched if not found

3. ELEMENTS BY TYPE:
   - Count of elements by type with the classification

4. ELEMENT DETAILS:
   - Element type and GUID
   - First few property values (when available)
   - Limited by MAX_ELEMENTS_TO_DISPLAY

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

If you have multiple Archicad instances open:
- The script connects to the first one opened
- Make sure the correct project is open in that instance
- Close other instances if you need to target a specific project

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Search is case-insensitive for both ID and name
- Searches through entire classification hierarchy recursively
- If classification not found, displays first 20 available classifications
- Property display is limited to first 3 properties per element

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 101_get_all_classification_systems.py
- 102_get_classification_system_with_hierarchy.py
- 201_set_element_classification.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# ============================================================================
# CONFIGURATION
# ============================================================================
CLASSIFICATION_SYSTEM = "Classification ARCHICAD"
CLASSIFICATION_SEARCH = "Mur"  # Can be an ID or a name
MAX_ELEMENTS_TO_DISPLAY = 10   # Number of elements to display with details

print(f"\n=== SEARCHING FOR CLASSIFIED ELEMENTS ===")
print(f"System: {CLASSIFICATION_SYSTEM}")
print(f"Search term: {CLASSIFICATION_SEARCH}")

# ============================================================================
# FIND CLASSIFICATION SYSTEM
# ============================================================================
classification_systems = acc.GetAllClassificationSystems()
target_system_guid = None

for system in classification_systems:
    if system.name == CLASSIFICATION_SYSTEM:
        target_system_guid = system.classificationSystemId.guid
        break

if not target_system_guid:
    print(f"\n✗ Classification system '{CLASSIFICATION_SYSTEM}' not found")
    print("\nAvailable classification systems:")
    for system in classification_systems:
        print(f"  - {system.name}")
    exit()

# ============================================================================
# GET ALL CLASSIFICATIONS IN SYSTEM
# ============================================================================
all_classifications = acc.GetAllClassificationsInSystem(
    act.ClassificationSystemId(target_system_guid)
)

print(f"\n=== SEARCHING THROUGH HIERARCHY ===")
print(f"(Searching for: '{CLASSIFICATION_SEARCH}')\n")

target_classification_item_id = None
classification_name = None
search_term_lower = CLASSIFICATION_SEARCH.lower()
all_found_items = []  # Store all items for display


def search_classification_recursive(item, level=0):
    """
    Recursively search through classification hierarchy

    Args:
        item: ClassificationItemArrayItem object
        level: Current depth level (for display)

    Returns:
        tuple: (classificationItemId, name) if found, else (None, None)
    """
    global target_classification_item_id, classification_name

    # Get the classification item
    class_item = item.classificationItem

    # Get the ID and name
    item_id = class_item.id
    item_name = class_item.name

    # Store for display
    indent = "  " * level
    all_found_items.append((level, item_id, item_name))

    # Check if this is the item we're searching for
    if (item_id.lower() == search_term_lower or
            item_name.lower() == search_term_lower):
        target_classification_item_id = class_item.classificationItemId
        classification_name = item_name
        print(f"{indent}✓ FOUND: '{item_id}' | '{item_name}'")
        return (class_item.classificationItemId, item_name)

    # Recursively search children
    if hasattr(class_item, 'children') and class_item.children:
        for child in class_item.children:
            result = search_classification_recursive(child, level + 1)
            if result[0] is not None:
                return result

    return (None, None)


# ============================================================================
# SEARCH THROUGH HIERARCHY
# ============================================================================
for root_item in all_classifications:
    result = search_classification_recursive(root_item, level=0)
    if result[0] is not None:
        break

if not target_classification_item_id:
    print(
        f"\n✗ Classification '{CLASSIFICATION_SEARCH}' not found in hierarchy")
    print(f"\nTotal classifications searched: {len(all_found_items)}")
    print(f"\nFirst 20 classifications found:")
    for i, (level, item_id, item_name) in enumerate(all_found_items[:20]):
        indent = "  " * level
        print(f"{indent}{item_id}: {item_name}")
    if len(all_found_items) > 20:
        print(f"... and {len(all_found_items) - 20} more")
    exit()

print(f"\n✓ Using classification: {classification_name}")

# ============================================================================
# GET ELEMENTS WITH THIS CLASSIFICATION
# ============================================================================
print(f"\n=== RETRIEVING ELEMENTS ===")

elements = acc.GetElementsByClassification(target_classification_item_id)

print(f"Total elements found: {len(elements)}")

if len(elements) == 0:
    print("\n⚠️  No elements have this classification")
    print("Try assigning the classification to elements first")
    exit()

# ============================================================================
# GET ELEMENT TYPES AND COUNT BY TYPE
# ============================================================================
element_ids = [elem.elementId for elem in elements]
element_types = acc.GetTypesOfElements(element_ids)

# Count by type
type_counts = {}
for elem_type in element_types:
    type_name = elem_type.typeOfElement.elementType
    type_counts[type_name] = type_counts.get(type_name, 0) + 1

# Display results
print(f"\n=== ELEMENTS BY TYPE ===")
for elem_type, count in sorted(type_counts.items()):
    print(f"  {elem_type}: {count}")

# ============================================================================
# SHOW ELEMENT DETAILS
# ============================================================================
print(
    f"\n=== FIRST {min(MAX_ELEMENTS_TO_DISPLAY, len(elements))} ELEMENTS ===")
for i in range(min(MAX_ELEMENTS_TO_DISPLAY, len(elements))):
    element = elements[i]
    elem_type = element_types[i]

    print(f"\n{i+1}. Type: {elem_type.typeOfElement.elementType}")
    print(f"   GUID: {element.elementId.guid}")

    # Try to get some properties for context
    try:
        property_ids = acc.GetAllPropertyIdsOfElements([element.elementId])
        if property_ids and property_ids[0].propertyIds:
            # Get first few properties
            first_props = property_ids[0].propertyIds[:3]
            prop_values = acc.GetPropertyValuesOfElements(
                [element.elementId], first_props)

            if prop_values:
                print(f"   Properties:")
                for prop_val in prop_values[0].propertyValues:
                    try:
                        # Try to extract value
                        if hasattr(prop_val.propertyValue, 'value'):
                            print(f"     - {prop_val.propertyValue.value}")
                    except:
                        pass
    except:
        pass

if len(elements) > MAX_ELEMENTS_TO_DISPLAY:
    print(f"\n... and {len(elements) - MAX_ELEMENTS_TO_DISPLAY} more elements")

print(f"\n{'='*70}")
print("SEARCH COMPLETE")
print(f"\n💡 TIP: Use 201_set_element_classification.py to")
print(f"         assign classifications to more elements")
