"""
================================================================================
SCRIPT: Get Element Full Info
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Gets complete information for selected elements including type, GUID,
classifications, property values, 2D/3D bounding boxes, and zone relationships.
Provides a comprehensive overview of element data.

This script is useful for detailed element inspection and debugging.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)
- At least one element selected in Archicad

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- MAX_ELEMENTS_TO_ANALYZE : Maximum number of elements to analyze (default: 10)
- MAX_ZONES_TO_CHECK      : Maximum number of zones to check for element containment
                            (default: 100, set to None to check all zones)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements
  
- GetTypesOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements
  
- GetAllClassificationSystems()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationSystems
  
- GetClassificationsOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetClassificationsOfElements
  
- GetAllClassificationsInSystem()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationsInSystem
  
- GetAllPropertyIds()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyIds
  
- GetDetailsOfProperties()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfProperties
  
- Get2DBoundingBoxes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.Get2DBoundingBoxes
  
- Get3DBoundingBoxes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.Get3DBoundingBoxes
  
- GetElementsRelatedToZones()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetElementsRelatedToZones
  
- GetElementsByType()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetElementsByType

[Utilities]
- GetPropertyValuesDictionary()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b3000utilities.Utilities.GetPropertyValuesDictionary

[Data Types]
- ElementIdArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementIdArrayItem
  
- TypeOfElementWrapper
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac26.html#archicad.releases.ac26.b3005types.TypeOfElementWrapper

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select one or more elements
3. Adjust MAX_ELEMENTS_TO_ANALYZE and MAX_ZONES_TO_CHECK if needed
4. Run this script
5. View complete element information in console

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
For each selected element, displays:

1. ELEMENT IDENTIFICATION:
   - Element number and GUID
   - Element type

2. CLASSIFICATIONS:
   - Classification ID and name for each system
   - Or "No classifications assigned"

3. PROPERTY VALUES:
   - All properties with non-empty values
   - Property group and name with value

4. 2D BOUNDING BOX:
   - Min/Max X and Y coordinates
   - Or "Not available" if cannot be retrieved

5. 3D BOUNDING BOX:
   - Min/Max X, Y, and Z coordinates
   - Calculated dimensions (Width x Depth x Height)
   - Or "Not available" if cannot be retrieved

6. ZONE RELATIONSHIPS:
   - If element is a Zone: List of elements contained in the zone with counts by type
   - If element is NOT a Zone: List of zones containing this element with zone names
   - Or "Not in any zone" / "No zones found"

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
- If no elements are selected, script exits with a warning
- Only properties with non-empty values are displayed
- Property names include group prefix to avoid confusion with duplicate names
- 2D/3D bounding boxes may not be available for all element types
- Zone relationship check is limited by MAX_ZONES_TO_CHECK for performance
- Set MAX_ZONES_TO_CHECK = None to check all zones (may be slow)
- For many elements, consider using MAX_ELEMENTS_TO_ANALYZE to limit output

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 105_get_selected_elements.py
- 303_list_element_property_values.py
- 304_list_element_property_values_from_all.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types
acu = conn.utilities

# ============================================================================
# CONFIGURATION
# ============================================================================
MAX_ELEMENTS_TO_ANALYZE = 10  # Maximum number of elements to analyze
# Maximum number of zones to check for element containment
MAX_ZONES_TO_CHECK = 100

# ============================================================================
# GET SELECTED ELEMENTS
# ============================================================================
selected_elements = acc.GetSelectedElements()

if len(selected_elements) == 0:
    print("\n⚠️  No elements selected")
    print("Please select at least one element in Archicad and run again")
    exit()

# Determine how many elements to analyze
elements_to_analyze = selected_elements[:MAX_ELEMENTS_TO_ANALYZE]

print(f"\n=== ANALYZING {len(elements_to_analyze)} ELEMENT(S) ===")
if len(selected_elements) > MAX_ELEMENTS_TO_ANALYZE:
    print(f"⚠️  {len(selected_elements) - MAX_ELEMENTS_TO_ANALYZE} elements not analyzed (limited by MAX_ELEMENTS_TO_ANALYZE)")
print()

# ============================================================================
# BUILD CLASSIFICATION LOOKUP (once for all elements)
# ============================================================================
classification_item_lookup = {}

try:
    classification_systems = acc.GetAllClassificationSystems()
    classification_system_ids = [
        system.classificationSystemId for system in classification_systems]

    def add_items_to_lookup(item_wrapper):
        """Recursively add classification items to lookup dictionary"""
        if hasattr(item_wrapper, 'classificationItem'):
            item = item_wrapper.classificationItem
            if hasattr(item, 'id') and hasattr(item, 'name') and hasattr(item, 'classificationItemId'):
                item_guid = item.classificationItemId.guid
                classification_item_lookup[item_guid] = (item.id, item.name)

            # Process children recursively
            if hasattr(item, 'children') and item.children:
                for child in item.children:
                    add_items_to_lookup(child)

    # Load all classification items from all systems
    for system in classification_systems:
        try:
            all_items = acc.GetAllClassificationsInSystem(
                system.classificationSystemId)
            for item_wrapper in all_items:
                add_items_to_lookup(item_wrapper)
        except:
            pass

    print(
        f"⏳ Loaded {len(classification_item_lookup)} classification items for lookup\n")
except:
    classification_systems = []
    classification_system_ids = []

# Get all property IDs once
all_property_ids = acc.GetAllPropertyIds()

# ============================================================================
# PROCESS EACH SELECTED ELEMENT
# ============================================================================
for idx, element in enumerate(elements_to_analyze):
    element_id = element.elementId

    print(f"{'='*70}")
    print(f"ELEMENT {idx+1}/{len(elements_to_analyze)}")
    print(f"{'='*70}")
    print(f"GUID: {element_id.guid}")

    # ========================================================================
    # GET ELEMENT TYPE
    # ========================================================================
    element_type = "Unknown"  # Initialize with default value
    try:
        element_types = acc.GetTypesOfElements([element_id])
        if element_types:
            element_type = element_types[0].typeOfElement.elementType
            print(f"Type: {element_type}")
        else:
            print(f"Type: Unknown")
    except:
        print(f"Type: Unknown")

    # ========================================================================
    # GET CLASSIFICATIONS
    # ========================================================================
    print("\n--- Classifications ---")
    if classification_system_ids:
        try:
            classifications = acc.GetClassificationsOfElements(
                [element_id], classification_system_ids)

            if classifications and classifications[0].classificationIds:
                elem_classifications = classifications[0].classificationIds

                has_classification = False
                for class_id_wrapper in elem_classifications:
                    if class_id_wrapper and hasattr(class_id_wrapper, 'classificationId'):
                        class_id = class_id_wrapper.classificationId
                        if hasattr(class_id, 'classificationItemId') and class_id.classificationItemId:
                            item_guid = class_id.classificationItemId.guid

                            # Lookup the classification item name
                            if item_guid in classification_item_lookup:
                                item_id, item_name = classification_item_lookup[item_guid]
                                if item_id and item_name:
                                    print(f"  {item_id} - {item_name}")
                                elif item_id:
                                    print(f"  {item_id}")
                                elif item_name:
                                    print(f"  {item_name}")
                                has_classification = True

                if not has_classification:
                    print("  No classifications assigned")
            else:
                print("  No classifications assigned")
        except:
            print("  Unable to retrieve classifications")
    else:
        print("  No classification systems available")

    # ========================================================================
    # GET PROPERTY VALUES
    # ========================================================================
    print("\n--- Property Values ---")
    try:
        # Get all property values for this element
        all_values_dict = acu.GetPropertyValuesDictionary(
            [element_id], all_property_ids)

        if element_id in all_values_dict:
            props_dict = all_values_dict[element_id]

            # Show only properties with values
            filled_props = {}
            for prop_id, value in props_dict.items():
                # Get property details including group
                try:
                    prop_details = acc.GetDetailsOfProperties([prop_id])
                    if prop_details:
                        prop_def = prop_details[0].propertyDefinition
                        prop_name = prop_def.name

                        # Get group name if available
                        if hasattr(prop_def, 'group') and prop_def.group:
                            try:
                                group_name = prop_def.group.name
                                # Create full property name with group
                                full_prop_name = f"{group_name}: {prop_name}"
                            except:
                                full_prop_name = prop_name
                        else:
                            full_prop_name = prop_name
                    else:
                        full_prop_name = str(prop_id)
                except:
                    full_prop_name = str(prop_id)

                # Check if value is filled
                if value is not None and str(value).strip():
                    filled_props[full_prop_name] = value

            if filled_props:
                for prop_name in sorted(filled_props.keys()):
                    value = filled_props[prop_name]
                    # Truncate long values
                    value_str = str(value)
                    if len(value_str) > 80:
                        value_str = value_str[:77] + "..."
                    print(f"  {prop_name}: {value_str}")
            else:
                print("  No property values set")
        else:
            print("  No properties available")
    except Exception as e:
        print(f"  Unable to retrieve properties")

    # ========================================================================
    # GET 2D BOUNDING BOX
    # ========================================================================
    print("\n--- 2D Bounding Box ---")
    try:
        bbox_2d = acc.Get2DBoundingBoxes([element_id])
        if bbox_2d and hasattr(bbox_2d[0], 'boundingBox2D'):
            box = bbox_2d[0].boundingBox2D
            print(f"  Min X: {box.xMin:.3f}, Max X: {box.xMax:.3f}")
            print(f"  Min Y: {box.yMin:.3f}, Max Y: {box.yMax:.3f}")
        else:
            print("  2D bounding box not available")
    except:
        print("  2D bounding box not available")

    # ========================================================================
    # GET 3D BOUNDING BOX
    # ========================================================================
    print("\n--- 3D Bounding Box ---")
    try:
        bbox_3d = acc.Get3DBoundingBoxes([element_id])
        if bbox_3d and hasattr(bbox_3d[0], 'boundingBox3D'):
            box = bbox_3d[0].boundingBox3D
            print(f"  Min X: {box.xMin:.3f}, Max X: {box.xMax:.3f}")
            print(f"  Min Y: {box.yMin:.3f}, Max Y: {box.yMax:.3f}")
            print(f"  Min Z: {box.zMin:.3f}, Max Z: {box.zMax:.3f}")

            # Calculate dimensions
            width = box.xMax - box.xMin
            depth = box.yMax - box.yMin
            height = box.zMax - box.zMin
            print(f"  Dimensions: {width:.3f} x {depth:.3f} x {height:.3f}")
        else:
            print("  3D bounding box not available")
    except:
        print("  3D bounding box not available")

    # ========================================================================
    # GET ZONE RELATIONSHIPS
    # ========================================================================
    print("\n--- Zone Relationships ---")
    try:
        if element_type == "Zone":
            # Case 1: Element IS a zone - list elements in this zone
            related_elements = acc.GetElementsRelatedToZones(
                [element_id], None)

            if related_elements and len(related_elements) > 0:
                if hasattr(related_elements[0], 'elements'):
                    elements_in_zone = related_elements[0].elements

                    if elements_in_zone:
                        print(
                            f"  This zone contains {len(elements_in_zone)} element(s):")

                        # Count by type
                        type_counts = {}
                        for elem in elements_in_zone:
                            elem_types = acc.GetTypesOfElements(
                                [elem.elementId])
                            if elem_types:
                                elem_type_name = elem_types[0].typeOfElement.elementType
                                type_counts[elem_type_name] = type_counts.get(
                                    elem_type_name, 0) + 1

                        # Display counts
                        for elem_type_name, count in sorted(type_counts.items()):
                            print(f"    • {elem_type_name}: {count}")
                    else:
                        print("  This zone contains no elements")
                else:
                    print("  Unable to retrieve elements in zone")
            else:
                print("  Unable to retrieve elements in zone")
        else:
            # Case 2: Element is NOT a zone - find which zone(s) it belongs to
            all_zones = acc.GetElementsByType("Zone")

            if all_zones:
                # Check which zones contain this element
                containing_zones = []

                # Limit zones to check based on configuration
                zones_to_check = all_zones[:MAX_ZONES_TO_CHECK] if MAX_ZONES_TO_CHECK else all_zones

                for zone in zones_to_check:
                    zone_id = zone.elementId
                    # Get elements related to this zone
                    try:
                        related = acc.GetElementsRelatedToZones(
                            [zone_id], None)

                        if related and hasattr(related[0], 'elements'):
                            zone_elements = related[0].elements
                            # Check if our element is in this zone
                            for zone_elem in zone_elements:
                                if zone_elem.elementId.guid == element_id.guid:
                                    containing_zones.append(zone_id)
                                    break
                    except:
                        continue

                if containing_zones:
                    print(
                        f"  This element is in {len(containing_zones)} zone(s):")
                    # Get zone names from properties
                    for zone_id in containing_zones:
                        try:
                            zone_props = acu.GetPropertyValuesDictionary(
                                [zone_id], all_property_ids)
                            if zone_id in zone_props:
                                zone_name = None
                                zone_number = None

                                for prop_id, value in zone_props[zone_id].items():
                                    try:
                                        prop_details = acc.GetDetailsOfProperties(
                                            [prop_id])
                                        if prop_details:
                                            prop_name = prop_details[0].propertyDefinition.name
                                            if prop_name.lower() in ["name", "nom", "zone name"]:
                                                zone_name = str(value)
                                            elif prop_name.lower() in ["number", "numéro", "zone number"]:
                                                zone_number = str(value)
                                    except:
                                        continue

                                # Display zone info
                                if zone_name and zone_number:
                                    print(f"    • {zone_number}: {zone_name}")
                                elif zone_name:
                                    print(f"    • {zone_name}")
                                elif zone_number:
                                    print(f"    • Zone {zone_number}")
                                else:
                                    print(f"    • Zone (GUID: {zone_id.guid})")
                        except:
                            print(f"    • Zone (GUID: {zone_id.guid})")
                else:
                    print("  This element is not in any zone")

                if MAX_ZONES_TO_CHECK and len(all_zones) > MAX_ZONES_TO_CHECK:
                    print(
                        f"  ⚠️  Only checked first {MAX_ZONES_TO_CHECK} of {len(all_zones)} zones")
            else:
                print("  No zones found in project")

    except Exception as e:
        print(f"  Unable to retrieve zone relationships")

    print()  # Empty line between elements

# ============================================================================
# SUMMARY
# ============================================================================
print(f"{'='*70}")
print("✓ Analysis complete")

if len(selected_elements) > MAX_ELEMENTS_TO_ANALYZE:
    print(f"\n⚠️  Only first {MAX_ELEMENTS_TO_ANALYZE} elements were analyzed")
    print("   To analyze all elements, increase MAX_ELEMENTS_TO_ANALYZE")

print("\n💡 TIP: Use other scripts for specific information:")
print("   - 303_list_element_property_values.py for property details")
print("   - 105_get_selected_elements.py for element type overview")
