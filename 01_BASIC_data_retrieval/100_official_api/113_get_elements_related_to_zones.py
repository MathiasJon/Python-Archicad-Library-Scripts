"""
================================================================================
SCRIPT: Get Elements Related to Zones
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Gets all elements that are associated with zones in the project. For each zone,
displays the zone information and lists all elements contained within it,
organized by element type.

This script is useful for understanding zone composition and verifying that
elements are correctly associated with their zones.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)
- At least one zone in the project

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- MAX_ZONES_TO_ANALYZE      : Maximum number of zones to analyze (default: 10)
- MAX_ELEMENTS_TO_DISPLAY   : Maximum number of elements to display per zone (default: 5)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetElementsByType()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetElementsByType
  
- GetElementsRelatedToZones()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetElementsRelatedToZones
  
- GetTypesOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements
  
- GetAllPropertyIds()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyIds
  
- GetDetailsOfProperties()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfProperties

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
1. Open an Archicad project with zones
2. Adjust MAX_ZONES_TO_ANALYZE and MAX_ELEMENTS_TO_DISPLAY if needed
3. Run this script
4. View zone composition in console

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
For each zone, displays:

1. ZONE IDENTIFICATION:
   - Zone GUID
   - Zone name from properties

2. ELEMENTS IN ZONE:
   - Total count of elements
   - Count by element type
   - List of first elements with type and GUID

3. SUMMARY:
   - Total zones analyzed

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
- Elements are 'related' to a zone if they are within the zone boundary
- GetElementsRelatedToZones requires None for elementTypes parameter
- Zone name search stops as soon as the name property is found
- For many zones, consider using MAX_ZONES_TO_ANALYZE to limit output

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 106_get_elements_by_type.py
- 111_get_element_full_info.py

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
MAX_ZONES_TO_ANALYZE = 10     # Maximum number of zones to analyze
MAX_ELEMENTS_TO_DISPLAY = 5   # Maximum number of elements to display per zone

# ============================================================================
# GET ALL ZONES
# ============================================================================
print("\n" + "="*70)
print("ZONES IN PROJECT")
print("="*70)

all_zones = acc.GetElementsByType("Zone")

if len(all_zones) == 0:
    print("\n⚠️  No zones found in the project")
    print("Create some zones in Archicad first")
    exit()

print(f"\nTotal zones found: {len(all_zones)}")

# Determine how many zones to analyze
zones_to_analyze = all_zones[:MAX_ZONES_TO_ANALYZE]

if len(all_zones) > MAX_ZONES_TO_ANALYZE:
    print(
        f"⚠️  Analyzing first {MAX_ZONES_TO_ANALYZE} zones only (limited by MAX_ZONES_TO_ANALYZE)")

print()

# Get all property IDs once for efficiency
all_property_ids = acc.GetAllPropertyIds()

# ============================================================================
# PROCESS EACH ZONE
# ============================================================================
for i, zone in enumerate(zones_to_analyze):
    zone_id = zone.elementId

    print(f"{'='*70}")
    print(f"ZONE {i+1}/{len(zones_to_analyze)}")
    print(f"{'='*70}")
    print(f"GUID: {zone_id.guid}")

    # ========================================================================
    # GET ZONE NAME
    # ========================================================================
    try:
        zone_props = acu.GetPropertyValuesDictionary(
            [zone_id], all_property_ids)

        if zone_id in zone_props:
            zone_name = None

            # Search until we find the zone name
            for prop_id, value in zone_props[zone_id].items():
                # Stop as soon as we found the zone name
                if zone_name:
                    break

                try:
                    prop_details = acc.GetDetailsOfProperties([prop_id])
                    if prop_details:
                        prop_name = prop_details[0].propertyDefinition.name

                        # Look for zone name property
                        if prop_name.lower() in ["name", "nom", "zone name"]:
                            zone_name = str(value) if value else None
                except:
                    continue

            # Display zone name
            if zone_name:
                print(f"Zone: {zone_name}")
    except:
        pass

    # ========================================================================
    # GET ELEMENTS RELATED TO THIS ZONE
    # ========================================================================
    try:
        # GetElementsRelatedToZones requires None for elementTypes parameter
        related_elements = acc.GetElementsRelatedToZones([zone_id], None)

        if related_elements:
            result = related_elements[0]

            # Check if it's an error or elements
            if hasattr(result, 'error'):
                print(f"\n✗ Error: {result.error}")
            elif hasattr(result, 'elements'):
                elements_wrapper = result.elements

                # ElementsWrapper might have an 'elements' attribute or be iterable
                zone_elements = None

                if hasattr(elements_wrapper, 'elements'):
                    zone_elements = elements_wrapper.elements
                elif hasattr(elements_wrapper, '__iter__'):
                    try:
                        zone_elements = list(elements_wrapper)
                    except:
                        pass

                if zone_elements:
                    print(f"\nElements in zone: {len(zone_elements)}")

                    # Get types of related elements
                    element_ids = [elem.elementId for elem in zone_elements]
                    element_types = acc.GetTypesOfElements(element_ids)

                    # Count by type
                    type_counts = {}
                    for elem_type in element_types:
                        type_name = elem_type.typeOfElement.elementType
                        type_counts[type_name] = type_counts.get(
                            type_name, 0) + 1

                    print("\nElements by type:")
                    for elem_type, count in sorted(type_counts.items()):
                        print(f"  {elem_type}: {count}")

                    # Show first few elements
                    print(
                        f"\nFirst {min(MAX_ELEMENTS_TO_DISPLAY, len(zone_elements))} elements:")
                    for j in range(min(MAX_ELEMENTS_TO_DISPLAY, len(zone_elements))):
                        elem_type = element_types[j].typeOfElement.elementType
                        elem_guid = zone_elements[j].elementId.guid
                        print(f"  {j+1}. {elem_type} - {elem_guid}")

                    if len(zone_elements) > MAX_ELEMENTS_TO_DISPLAY:
                        print(
                            f"  ... and {len(zone_elements) - MAX_ELEMENTS_TO_DISPLAY} more")
                else:
                    print("\n⚠️  No elements related to this zone")
            else:
                print(f"\n⚠️  Unexpected result structure")
        else:
            print("\n⚠️  No result from GetElementsRelatedToZones")

    except Exception as e:
        print(f"\n✗ Error getting related elements: {e}")

    print()

# ============================================================================
# SUMMARY
# ============================================================================
print(f"{'='*70}")
print("SUMMARY")
print(f"{'='*70}")
print(f"Total zones in project: {len(all_zones)}")
print(f"Zones analyzed: {len(zones_to_analyze)}")

if len(all_zones) > MAX_ZONES_TO_ANALYZE:
    print(f"\n⚠️  Only first {MAX_ZONES_TO_ANALYZE} zones were analyzed")
    print("   To analyze all zones, increase MAX_ZONES_TO_ANALYZE")

print("\n✓ Analysis complete")
print("\n💡 TIP: Elements are 'related' to a zone if they are within")
print("         the zone boundary or associated with it")
