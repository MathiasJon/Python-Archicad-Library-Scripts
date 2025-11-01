"""
================================================================================
SCRIPT: Get All Properties
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves all properties defined in the project and organizes them by property
groups. Displays property names, types, and descriptions for each group.

This script is useful for discovering available properties before working with
element property values.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- MAX_PROPERTIES_PER_GROUP : Maximum number of properties to display per group
                             (default: None = show all)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAllPropertyIds()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyIds
  
- GetDetailsOfProperties()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfProperties
  
- GetPropertyGroups()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetPropertyGroups

[Data Types]
- PropertyId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.PropertyId
  
- PropertyDefinition
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.PropertyDefinition
  
- PropertyGroup
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.PropertyGroup

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Adjust MAX_PROPERTIES_PER_GROUP if needed
3. Run this script
4. View all available properties organized by groups

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. TOTAL PROPERTY COUNT:
   - Number of properties found in the project

2. PROPERTIES BY GROUP:
   - Group name with property count
   - Property names with types and descriptions
   - Limited by MAX_PROPERTIES_PER_GROUP if set

3. SUMMARY STATISTICS:
   - Total property groups
   - Total properties
   - Top 5 groups by property count

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
- Properties are organized by their property groups
- Property descriptions are shown for first 5 properties per group only
- Long descriptions are truncated to 80 characters
- Properties without groups are listed under "No Group"

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 109_get_properties_by_classification.py
- 202_set_property_value.py
- 303_list_element_property_values.py

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
MAX_PROPERTIES_PER_GROUP = None  # Set to a number to limit display, None = show all

print(f"\n=== RETRIEVING ALL PROPERTIES ===")

# ============================================================================
# GET ALL PROPERTY IDS
# ============================================================================
all_property_ids = acc.GetAllPropertyIds()
print(f"Total properties found: {len(all_property_ids)}")

if len(all_property_ids) == 0:
    print("\n⚠️  No properties found in this project")
    exit()

# ============================================================================
# GET PROPERTY DETAILS
# ============================================================================
print(f"\n=== LOADING PROPERTY DETAILS ===")
property_details = acc.GetDetailsOfProperties(all_property_ids)

# ============================================================================
# ORGANIZE PROPERTIES BY GROUP
# ============================================================================
properties_by_group = {}

for prop_detail in property_details:
    prop_def = prop_detail.propertyDefinition

    # Get property information
    prop_name = prop_def.name

    # Get group information
    if hasattr(prop_def, 'group') and prop_def.group:
        # The group object has the name directly
        try:
            group_name = prop_def.group.name
        except AttributeError:
            # Fallback: try to get group details
            try:
                group_details = acc.GetPropertyGroups([prop_def.group])
                if group_details:
                    group_name = group_details[0].name
                else:
                    group_name = "Unknown Group"
            except:
                group_name = "Unknown Group"
    else:
        group_name = "No Group"

    # Add to dictionary
    if group_name not in properties_by_group:
        properties_by_group[group_name] = []

    # Store property info
    prop_info = {
        'name': prop_name,
        'id': prop_def.propertyId if hasattr(prop_def, 'propertyId') else None,
        'description': prop_def.description if hasattr(prop_def, 'description') else None,
        'type': None
    }

    # Try to get property type
    if hasattr(prop_def, 'defaultValue') and prop_def.defaultValue:
        default_val = prop_def.defaultValue
        if hasattr(default_val, 'type'):
            prop_info['type'] = default_val.type

    properties_by_group[group_name].append(prop_info)

# ============================================================================
# DISPLAY PROPERTIES BY GROUP
# ============================================================================
print(f"\n=== PROPERTIES BY GROUP ===")
print(f"Total groups: {len(properties_by_group)}")
print(f"Total properties: {len(property_details)}\n")

for group_name in sorted(properties_by_group.keys()):
    properties = properties_by_group[group_name]

    print(f"\n{'='*70}")
    print(f"📁 {group_name}")
    print(f"{'='*70}")
    print(f"Properties: {len(properties)}\n")

    # Determine how many properties to display
    props_to_display = properties if MAX_PROPERTIES_PER_GROUP is None else properties[
        :MAX_PROPERTIES_PER_GROUP]

    # Display properties in this group
    for i, prop in enumerate(props_to_display, 1):
        print(f"{i}. {prop['name']}")

        # Show type if available
        if prop['type']:
            print(f"   Type: {prop['type']}")

        # Show description if available (only for first 5 to avoid clutter)
        if prop['description'] and i <= 5:
            # Truncate long descriptions
            desc = prop['description']
            if len(desc) > 80:
                desc = desc[:77] + "..."
            print(f"   Description: {desc}")

    # Show if there are more properties
    if MAX_PROPERTIES_PER_GROUP and len(properties) > MAX_PROPERTIES_PER_GROUP:
        print(
            f"\n... and {len(properties) - MAX_PROPERTIES_PER_GROUP} more properties")

    # Add spacing between groups
    print()

# ============================================================================
# SUMMARY STATISTICS
# ============================================================================
print(f"\n{'='*70}")
print(f"=== SUMMARY ===")
print(f"{'='*70}")
print(f"Total property groups: {len(properties_by_group)}")
print(f"Total properties: {len(property_details)}")

# Show top 5 groups by property count
print(f"\n=== TOP 5 GROUPS BY PROPERTY COUNT ===")
sorted_groups = sorted(properties_by_group.items(),
                       key=lambda x: len(x[1]), reverse=True)
for i, (group_name, properties) in enumerate(sorted_groups[:5], 1):
    print(f"{i}. {group_name}: {len(properties)} properties")

print(f"\n{'='*70}")
print("✓ Property retrieval complete")
print(f"\n💡 TIP: Use these property names in scripts like:")
print(f"   - 109_get_properties_by_classification.py")
print(f"   - 202_set_property_value.py")
print(f"   - 303_list_element_property_values.py")
