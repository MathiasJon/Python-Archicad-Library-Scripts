"""
================================================================================
SCRIPT: Get Classifications by Property
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Shows which classification items have access to which properties by building
a reverse mapping from properties to classifications. Displays properties with
the classification items that can use them.

This script is useful for understanding property-classification relationships
and finding which classifications support specific properties.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- MAX_PROPERTIES_TO_DISPLAY          : Maximum number of properties to display
                                       (default: 10)
- MAX_CLASSIFICATIONS_PER_PROPERTY   : Maximum classifications shown per property
                                       (default: 5)

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
1. Open an Archicad project with classification systems and properties
2. Adjust MAX_PROPERTIES_TO_DISPLAY and MAX_CLASSIFICATIONS_PER_PROPERTY if needed
3. Run this script
4. View properties with their associated classifications

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. MAPPING PROCESS:
   - Progress indicator during analysis
   - Total classification items analyzed

2. SAMPLE PROPERTIES:
   - Property name (group: name format)
   - Number of classifications that can use it
   - List of classifications (limited by MAX_CLASSIFICATIONS_PER_PROPERTY)

3. STATISTICS:
   - Total properties with classification associations
   - Total property-classification associations
   - Average classifications per property

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
- This script may take a moment to complete for projects with many classifications
- Properties are identified by GUID and mapped to classification items
- Only properties that are available to at least one classification are shown
- Property details are retrieved safely with error handling

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 108_get_all_properties.py
- 109_get_properties_by_classification.py
- 101_get_all_classification_systems.py
- 102_get_classification_system_with_hierarchy.py

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
MAX_PROPERTIES_TO_DISPLAY = 10        # Number of properties to display
MAX_CLASSIFICATIONS_PER_PROPERTY = 5  # Classifications shown per property


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
        List of PropertyId objects
    """
    try:
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
    except:
        return []


def get_property_detail_safe(property_id):
    """
    Get details of a single property safely.

    Args:
        property_id: PropertyId object

    Returns:
        Property definition or None
    """
    try:
        prop_array = [act.PropertyIdArrayItem(property_id)]
        properties = acc.GetDetailsOfProperties(prop_array)

        if properties and len(properties) > 0:
            prop_result = properties[0]
            if hasattr(prop_result, 'propertyDefinition'):
                return prop_result.propertyDefinition
    except:
        pass

    return None


def build_property_to_classification_map():
    """
    Build a mapping of properties to classification items that can use them.

    Returns:
        tuple: (property_map, total_items_analyzed)
            - property_map: Dictionary mapping property GUID to list of classification items
            - total_items_analyzed: Total number of classification items processed
    """
    property_map = {}

    # Get all classification systems
    systems = get_all_classification_systems()

    total_items = 0
    for system in systems:
        # Get items in this system
        items = get_classification_items(system.classificationSystemId)
        total_items += len(items)

        for item in items:
            try:
                if hasattr(item, 'classificationItem'):
                    item_obj = item.classificationItem

                    if hasattr(item_obj, 'classificationItemId'):
                        # Get properties available for this item
                        available_props = get_properties_for_classification_item(
                            item_obj.classificationItemId
                        )

                        # Add this classification to each property's list
                        for prop_id in available_props:
                            prop_guid = prop_id.guid

                            if prop_guid not in property_map:
                                property_map[prop_guid] = {
                                    'property_id': prop_id,
                                    'classifications': []
                                }

                            property_map[prop_guid]['classifications'].append({
                                'system': system.name,
                                'item_id': item_obj.id if hasattr(item_obj, 'id') else 'Unknown',
                                'item_name': item_obj.name if hasattr(item_obj, 'name') else 'Unknown'
                            })
            except:
                continue

    return property_map, total_items


# ============================================================================
# MAIN SCRIPT
# ============================================================================
print("\n" + "="*80)
print("CLASSIFICATIONS BY PROPERTY")
print("="*80)

print("\n📊 Building property-to-classification mapping...")
print("This may take a moment...\n")

# ============================================================================
# BUILD THE MAPPING
# ============================================================================
property_map, total_items = build_property_to_classification_map()

if not property_map:
    print("⚠️  No property-classification relationships found.")
    exit()

print(f"✓ Analyzed {total_items} classification items")
print(
    f"✓ Found {len(property_map)} properties with classification associations\n")

# ============================================================================
# DISPLAY SAMPLE PROPERTIES AND THEIR CLASSIFICATIONS
# ============================================================================
print("="*80)
print("SAMPLE PROPERTIES AND THEIR CLASSIFICATIONS")
print("="*80 + "\n")

sample_count = min(MAX_PROPERTIES_TO_DISPLAY, len(property_map))

for idx, (prop_guid, prop_data) in enumerate(list(property_map.items())[:sample_count], 1):
    # Try to get property details
    prop_def = get_property_detail_safe(prop_data['property_id'])

    if prop_def:
        prop_name = f"{prop_def.group.name}: {prop_def.name}"
    else:
        prop_name = f"Property {prop_guid[:16]}..."

    classifications = prop_data['classifications']

    print(f"{idx}. {prop_name}")
    print(f"   ✓ Used by {len(classifications)} classification item(s):")

    # Show limited classifications
    display_count = min(MAX_CLASSIFICATIONS_PER_PROPERTY, len(classifications))
    for class_item in classifications[:display_count]:
        print(
            f"     • [{class_item['system']}] {class_item['item_id']}: {class_item['item_name']}")

    if len(classifications) > MAX_CLASSIFICATIONS_PER_PROPERTY:
        print(
            f"     ... and {len(classifications) - MAX_CLASSIFICATIONS_PER_PROPERTY} more")

    print()

if len(property_map) > sample_count:
    print(f"... and {len(property_map) - sample_count} more properties\n")

print("="*80)

# ============================================================================
# STATISTICS
# ============================================================================
total_associations = sum(len(v['classifications'])
                         for v in property_map.values())
avg_classifications = total_associations / \
    len(property_map) if property_map else 0

print(f"\n📈 STATISTICS:")
print(f"  Total properties: {len(property_map)}")
print(f"  Total associations: {total_associations}")
print(f"  Average classifications per property: {avg_classifications:.1f}")

print("\n" + "="*80)
print("✓ Analysis complete")
print("\n💡 TIP: Use this information to understand which properties work with")
print("         which classifications in your project")
