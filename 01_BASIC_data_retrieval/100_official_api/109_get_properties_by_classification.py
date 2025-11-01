"""
================================================================================
SCRIPT: Get Properties by Classification
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves classification systems and shows which properties are available for
each classification item. Displays classification items with their associated
properties organized by classification system.

This script is useful for understanding which properties are available for
different classification items before assigning values.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- MAX_ITEMS_PER_SYSTEM    : Maximum classification items to display per system
                            (default: None = show all)
- MAX_SAMPLE_PROPERTIES   : Number of sample properties to display per item
                            (default: 3)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetClassificationSystemIds()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetClassificationSystemIds
  
- GetClassificationSystems()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetClassificationSystems
  
- GetAllClassificationsInSystem()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationsInSystem
  
- GetClassificationItemAvailability()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetClassificationItemAvailability
  
- GetDetailsOfProperties()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfProperties

[Data Types]
- ClassificationSystem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationSystem
  
- ClassificationItemArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationItemArrayItem
  
- PropertyId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.PropertyId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project with classification systems
2. Adjust MAX_ITEMS_PER_SYSTEM and MAX_SAMPLE_PROPERTIES if needed
3. Run this script
4. View classification items with their available properties

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. CLASSIFICATION SYSTEMS SUMMARY:
   - Total number of systems found

2. FOR EACH SYSTEM:
   - System name
   - Total classification items
   - Item ID and name
   - Number of available properties
   - Sample properties (limited by MAX_SAMPLE_PROPERTIES)

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
- Property details are retrieved one by one to avoid total failure if one fails
- If property details cannot be retrieved, GUIDs are displayed instead
- Only sample properties are shown per item to avoid cluttering console
- Error handling ensures script continues even if individual items fail

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 101_get_all_classification_systems.py
- 102_get_classification_system_with_hierarchy.py
- 108_get_all_properties.py
- 110_get_classifications_by_property.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities

# ============================================================================
# CONFIGURATION
# ============================================================================
MAX_ITEMS_PER_SYSTEM = None  # Set to a number to limit display, None = show all
MAX_SAMPLE_PROPERTIES = 3     # Number of sample properties to display per item


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
def get_all_classification_systems():
    """
    Get all classification systems in the project.

    Returns:
        List of classification system details
    """
    # Get classification system IDs
    system_ids = acc.GetClassificationSystemIds()

    if not system_ids:
        return []

    # Get details of classification systems
    systems = acc.GetClassificationSystems(system_ids)

    return [system.classificationSystem for system in systems if hasattr(system, 'classificationSystem')]


def get_classification_items(system_id):
    """
    Get all classification items in a system.

    Args:
        system_id: Classification system ID

    Returns:
        List of classification items
    """
    try:
        items = acc.GetAllClassificationsInSystem(system_id)
        return items
    except:
        return []


def get_properties_for_classification_item(classification_item_id):
    """
    Get all properties available for a specific classification item.

    Args:
        classification_item_id: ClassificationItemId object

    Returns:
        List of PropertyId objects (extracted from PropertyIdArrayItem)
    """
    try:
        # Get property availability for this classification item
        availability = acc.GetClassificationItemAvailability([
            act.ClassificationItemIdArrayItem(classification_item_id)
        ])

        if availability and len(availability) > 0:
            avail_item = availability[0]
            if hasattr(avail_item, 'classificationItemAvailability'):
                # Extract propertyId from each PropertyIdArrayItem
                prop_array_items = avail_item.classificationItemAvailability.availableProperties
                return [item.propertyId for item in prop_array_items]

        return []
    except Exception as e:
        return []


def get_property_details_safe(property_ids):
    """
    Get details of properties with error handling.

    Args:
        property_ids: List of PropertyId objects

    Returns:
        List of property definitions
    """
    if not property_ids:
        return []

    result = []

    # Try to get details one by one to avoid total failure
    for prop_id in property_ids:
        try:
            # Create PropertyIdArrayItem from PropertyId
            prop_array = [act.PropertyIdArrayItem(prop_id)]
            properties = acc.GetDetailsOfProperties(prop_array)

            if properties and len(properties) > 0:
                prop_result = properties[0]
                if hasattr(prop_result, 'propertyDefinition'):
                    result.append(prop_result.propertyDefinition)
        except:
            # Skip properties that fail
            continue

    return result


# ============================================================================
# MAIN SCRIPT
# ============================================================================
print("\n" + "="*80)
print("PROPERTIES BY CLASSIFICATION")
print("="*80)

# Get all classification systems
print("\n📊 Retrieving classification systems...")
systems = get_all_classification_systems()

if not systems:
    print("⚠️  No classification systems found.")
    exit()

print(f"Found {len(systems)} classification system(s)\n")

# ============================================================================
# PROCESS EACH CLASSIFICATION SYSTEM
# ============================================================================
for system in systems:
    print(f"\n{'='*80}")
    print(f"Classification System: {system.name}")
    print(f"{'='*80}")

    # Get classification items
    items = get_classification_items(system.classificationSystemId)

    if not items:
        print("  No classification items found.")
        continue

    print(f"  Total items: {len(items)}\n")

    # Determine how many items to display
    items_to_display = items if MAX_ITEMS_PER_SYSTEM is None else items[:MAX_ITEMS_PER_SYSTEM]

    # Show items with their available properties
    for idx, item in enumerate(items_to_display, 1):
        try:
            if hasattr(item, 'classificationItem'):
                item_obj = item.classificationItem
                item_id = item_obj.id if hasattr(item_obj, 'id') else 'Unknown'
                item_name = item_obj.name if hasattr(
                    item_obj, 'name') else item_id

                print(f"  {idx}. {item_id}: {item_name}")

                # Get available properties for this item
                if hasattr(item_obj, 'classificationItemId'):
                    available_props = get_properties_for_classification_item(
                        item_obj.classificationItemId
                    )

                    if available_props:
                        print(
                            f"     ✓ {len(available_props)} properties available")

                        # Try to get details for sample properties
                        prop_details = get_property_details_safe(
                            available_props[:MAX_SAMPLE_PROPERTIES])

                        if prop_details:
                            print(f"     Sample properties:")
                            for prop_def in prop_details:
                                print(
                                    f"       • {prop_def.group.name}: {prop_def.name}")
                        else:
                            # Show GUIDs if details fail
                            print(
                                f"     Property GUIDs (first {MAX_SAMPLE_PROPERTIES}):")
                            for i, prop_id in enumerate(available_props[:MAX_SAMPLE_PROPERTIES], 1):
                                print(f"       {i}. {prop_id.guid}")
                    else:
                        print(f"     No properties available")

                print()
        except Exception as e:
            print(f"     Error processing item {idx}: {e}")
            print()
            continue

    # Show if there are more items
    if MAX_ITEMS_PER_SYSTEM and len(items) > MAX_ITEMS_PER_SYSTEM:
        print(f"  ... and {len(items) - MAX_ITEMS_PER_SYSTEM} more items\n")

# ============================================================================
# SUMMARY
# ============================================================================
print("="*80)
print("✓ Property retrieval complete")
print("\n💡 TIP: Use these properties in scripts like:")
print("   - 202_set_property_value.py")
print("   - 303_list_element_property_values.py")
