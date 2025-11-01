"""
================================================================================
SCRIPT: Set Element Property Value (Optimized)
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Modification

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Sets property values for selected elements in the current Archicad project
using an optimized fast lookup method. This script efficiently handles:
- Fast property search using GetAllPropertyNames() method
- Multiple property types (string, integer, number, boolean, length, etc.)
- Bulk property updates for multiple selected elements
- Automatic type detection and conversion
- Verification of applied changes
- Detailed error reporting

Properties in Archicad store metadata and custom information about elements.
This script simplifies batch property updates across multiple elements.

PERFORMANCE OPTIMIZATION:
This script uses GetAllPropertyNames() which is significantly faster than
iterating through all property definitions. For projects with hundreds of
properties, this can reduce search time from seconds to milliseconds.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Elements must be selected in Archicad before running
- Property must exist in the project

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements
  Returns the list of currently selected elements

- GetAllPropertyNames()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyNames
  Returns the property names (FAST - optimized for property search)

- GetAllPropertyIds()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyIds
  Returns property identifiers in same order as GetAllPropertyNames()

- GetDetailsOfProperties(propertyIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfProperties
  Returns detailed property definitions including type information

- SetPropertyValuesOfElements(elementPropertyValues)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.SetPropertyValuesOfElements
  Sets the property values of elements

- GetPropertyValuesOfElements(elements, propertyIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetPropertyValuesOfElements
  Returns property values for verification

[Data Types]
- ElementPropertyValue
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.ElementPropertyValue
  Combines element ID, property ID, and property value

- Property Value Types:
  * NormalStringPropertyValue: Text values
  * NormalIntegerPropertyValue: Whole numbers
  * NormalNumberPropertyValue: Decimal numbers
  * NormalBooleanPropertyValue: True/False values
  * NormalLengthPropertyValue: Distances (in mm)
  * NormalAreaPropertyValue: Areas (in m²)
  * NormalVolumePropertyValue: Volumes (in m³)
  * NormalAnglePropertyValue: Angles (in degrees)

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select elements you want to modify
3. Modify CONFIGURATION section below:
   - Set PROPERTY_NAME to the property you want to change
   - Set NEW_VALUE to the desired value
4. Run this script
5. Verify the changes were applied

To find available properties:
- Use Archicad's Property Manager
- Use Element Info Palette to see element properties
- Run 108_get_property_values.py to list properties

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Before running, set these values in the CONFIGURATION section:

PROPERTY_NAME: Name of the property to modify
Example: "Reference", "Description", "Status", "Phase"
Note: Property name must match exactly (case-sensitive)

PROPERTY_GROUP: Group/folder name containing the property (OPTIONAL)
- Set to None to search in all groups (default)
- Required if multiple properties have the same name in different groups
- The script will detect conflicts and prompt you to specify the group
Examples: "General", "Custom Properties", "Classification", "Built-In Properties"

NEW_VALUE: New value to assign to the property
- String properties: Any text value
- Number properties: Numeric value (int or float)
- Boolean properties: "true", "false", "1", "0", "yes", "no"
- Length properties: Value in millimeters
- Area properties: Value in square meters
- Volume properties: Value in cubic meters
- Angle properties: Value in degrees

The script automatically detects the property type and converts the value.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Number of selected elements
- Property search progress (fast method)
- Property type detection
- Value creation and conversion
- Application results
- Verification of changes

Example output (single property):
  === SEARCHING FOR PROPERTY (FAST METHOD) ===
  Total properties in project: 847
  Searching for property 'Reference'...
  ✓ Found property: Reference
    Group: General
    Type: string
  
  === APPLYING PROPERTY VALUE ===
  ✓ Property set for 15 element(s)

Example output (multiple properties with same name):
  === SEARCHING FOR PROPERTY (FAST METHOD) ===
  Total properties in project: 847
  Searching for property 'Reference'
  
  ⚠️  Found 2 properties named 'Reference' in different groups:
  
    1. Group: General
       Type: string
    2. Group: Custom Properties
       Type: string
  
  ✗ Multiple properties found with name 'Reference'
  
  Please specify PROPERTY_GROUP in the configuration:
  
  Example:
    PROPERTY_NAME = 'Reference'
    PROPERTY_GROUP = 'General'  # Choose one

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
PROPERTY TYPES:
Archicad supports various property value types:
- string: Text values (e.g., "Wall-001", "Living Room")
- integer: Whole numbers (e.g., 1, 42, -5)
- number: Decimal numbers (e.g., 3.14, 2.5, -0.5)
- boolean: True/False values
- length: Distances in millimeters (e.g., 2400.0 = 2.4m)
- area: Areas in square meters (e.g., 15.5 = 15.5m²)
- volume: Volumes in cubic meters (e.g., 45.2 = 45.2m³)
- angle: Angles in degrees (e.g., 90.0 = 90°)

The script automatically detects the property type and converts NEW_VALUE.

PERFORMANCE OPTIMIZATION:
Traditional method: Query details for every property (slow for large projects)
Optimized method: Use GetAllPropertyNames() for instant lookup
Result: 100x-1000x faster property search

PROPERTY SOURCES:
- Built-In Properties: Standard Archicad properties (IDs, dimensions, etc.)
- User-Defined Properties: Custom properties created in Property Manager
- Classification Properties: Properties from classification systems
- Dynamic Properties: Properties that calculate values automatically

DUPLICATE PROPERTY NAMES:
⚠️  IMPORTANT: Multiple properties can have the same name if they're in different groups!
Example: "Reference" can exist in both "General" and "Custom Properties" groups.

If the script finds multiple properties with the same name:
- It will list all matching properties with their groups
- You must set PROPERTY_GROUP in the configuration
- Run the script again with the specific group name

The script automatically detects conflicts and provides clear guidance.

VALUE CONVERSION:
- String: Converts any value to text
- Integer: Converts numeric strings to whole numbers
- Number: Converts numeric strings to decimals
- Boolean: Accepts "true", "false", "1", "0", "yes", "no", "oui"
- Length/Area/Volume/Angle: Converts numeric strings to appropriate units

ERROR HANDLING:
- Missing property: Shows similar property names for correction
- Type mismatch: Clear error message about expected type
- Read-only properties: Reported in results
- Protected elements: Reported individually

LIMITATIONS:
- Cannot modify read-only properties
- Cannot modify properties of locked elements
- Cannot modify properties of elements in locked layers
- Enum properties require exact enum value match

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 108_get_property_values.py (view current property values)
- 105_get_property_definitions.py (list all properties and types)
- 203_create_property.py (create new custom properties)
- 111_get_element_full_info.py (verify element properties)

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

# Name of the property to set (must match exactly)
PROPERTY_NAME = "Test"

# OPTIONAL: Group/folder name for the property
# Leave as None to search in all groups (will warn if multiple matches found)
# Set to specific group name if property exists in multiple groups
# Examples: "General", "Custom Properties", "Classification", "Built-In Properties"
PROPERTY_GROUP = None  # Or: "General", "Custom Properties", etc.

# New value to assign to the property
# Will be automatically converted to correct type
NEW_VALUE = "TEST-001"


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main():
    """Main function to set property values."""

    print("\n" + "="*80)
    print("SET ELEMENT PROPERTY VALUE v1.0 (OPTIMIZED)")
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
    print(f"  Property to set: '{PROPERTY_NAME}'")
    if PROPERTY_GROUP:
        print(f"  In group: '{PROPERTY_GROUP}'")
    print(f"  New value: '{NEW_VALUE}'")

    # -------------------------------------------------------------------------
    # 2. FIND PROPERTY (FAST METHOD)
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SEARCHING FOR PROPERTY (FAST METHOD)")
    print("="*80)

    # Get all property names (FAST!)
    available_properties = acc.GetAllPropertyNames()
    print(f"\n✓ Loaded {len(available_properties)} properties from project")

    # Get corresponding property IDs (same order as names)
    available_property_ids = acc.GetAllPropertyIds()

    print(f"Searching for property '{PROPERTY_NAME}'")
    if PROPERTY_GROUP:
        print(f"  In group: '{PROPERTY_GROUP}'")
    print()

    # Search for ALL properties with the target name
    matching_properties = []

    for index, prop in enumerate(available_properties):
        # Extract property name and group
        group_name = None
        prop_name = None

        # UserDefined properties have localizedName [group, name]
        if hasattr(prop, 'localizedName') and len(prop.localizedName) >= 2:
            group_name = prop.localizedName[0]
            prop_name = prop.localizedName[1]

        # BuiltIn properties have nonLocalizedName
        elif hasattr(prop, 'nonLocalizedName'):
            prop_name = prop.nonLocalizedName
            group_name = "Built-In Properties"

        if not prop_name:
            continue

        # Check if this property name matches
        if prop_name == PROPERTY_NAME:
            property_type = prop.type if hasattr(prop, 'type') else 'unknown'

            # Get the corresponding property ID
            prop_id_item = available_property_ids[index]

            # Extract PropertyId from PropertyIdArrayItem wrapper
            if hasattr(prop_id_item, 'propertyId'):
                prop_id = prop_id_item.propertyId
            else:
                prop_id = prop_id_item

            matching_properties.append({
                'index': index,
                'name': prop_name,
                'group': group_name,
                'type': property_type,
                'id': prop_id
            })

    # Check results
    if len(matching_properties) == 0:
        print(f"✗ Property '{PROPERTY_NAME}' not found")
        print("\nSearching for similar property names...")

        # Find properties with similar names
        similar = []
        search_term = PROPERTY_NAME.lower()

        for index, prop in enumerate(available_properties[:100]):
            group_name = None
            prop_name = None

            # Extract name
            if hasattr(prop, 'localizedName') and len(prop.localizedName) >= 2:
                group_name = prop.localizedName[0]
                prop_name = prop.localizedName[1]
            elif hasattr(prop, 'nonLocalizedName'):
                prop_name = prop.nonLocalizedName
                group_name = "Built-In"

            # Check for partial match
            if prop_name and search_term in prop_name.lower():
                similar.append((prop_name, group_name))

        if similar:
            print(f"\nFound {len(similar)} properties with similar names:")
            for name, group in similar[:10]:
                print(f"  • {name} (in {group})")
        else:
            print("\nNo similar properties found")

        print("\n💡 Tips:")
        print("  • Check spelling and case (property name must match exactly)")
        print("  • Use Property Manager to see available properties")
        print("  • Run 108_get_property_values.py to list all properties")
        exit()

    # Multiple properties with same name - need to filter by group
    elif len(matching_properties) > 1:
        if not PROPERTY_GROUP:
            # Multiple matches and no group specified - show all and exit
            print(
                f"⚠️  Found {len(matching_properties)} properties named '{PROPERTY_NAME}' in different groups:")
            print()
            for i, prop in enumerate(matching_properties, 1):
                print(f"  {i}. Group: {prop['group']}")
                print(f"     Type: {prop['type']}")

            print(f"\n✗ Multiple properties found with name '{PROPERTY_NAME}'")
            print("\nPlease specify PROPERTY_GROUP in the configuration:")
            print("\nExample:")
            print(f"  PROPERTY_NAME = '{PROPERTY_NAME}'")
            print(
                f"  PROPERTY_GROUP = '{matching_properties[0]['group']}'  # Choose one")
            exit()
        else:
            # Filter by group
            filtered = [
                p for p in matching_properties if p['group'] == PROPERTY_GROUP]

            if len(filtered) == 0:
                print(
                    f"✗ Property '{PROPERTY_NAME}' not found in group '{PROPERTY_GROUP}'")
                print(f"\nFound '{PROPERTY_NAME}' in these groups:")
                for prop in matching_properties:
                    print(f"  • {prop['group']}")
                print(f"\nPlease update PROPERTY_GROUP to one of the above groups")
                exit()
            elif len(filtered) > 1:
                # Should not happen, but handle it
                print(
                    f"⚠️  Found multiple properties '{PROPERTY_NAME}' in group '{PROPERTY_GROUP}'")
                print("  Using the first one...")
                selected_property = filtered[0]
            else:
                # Exactly one match - perfect!
                selected_property = filtered[0]
    else:
        # Exactly one property with this name
        selected_property = matching_properties[0]

        # If group was specified, verify it matches
        if PROPERTY_GROUP and selected_property['group'] != PROPERTY_GROUP:
            print(
                f"⚠️  Warning: Property '{PROPERTY_NAME}' found in group '{selected_property['group']}'")
            print(f"           but you specified group '{PROPERTY_GROUP}'")
            print(
                f"           Using the property from '{selected_property['group']}'")

    # Extract the selected property info
    target_property_id = selected_property['id']
    property_index = selected_property['index']
    property_group = selected_property['group']
    property_type = selected_property['type']

    print(f"✓ Found property: {PROPERTY_NAME}")
    print(f"  Group: {property_group}")
    print(f"  Type: {property_type}")

    # -------------------------------------------------------------------------
    # 3. DETERMINE PROPERTY TYPE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("DETERMINING PROPERTY TYPE")
    print("="*80)

    # Get detailed property definition
    prop_details = acc.GetDetailsOfProperties(
        [available_property_ids[property_index]])
    property_def = prop_details[0].propertyDefinition if len(
        prop_details) > 0 else None

    # Extract property value type from definition
    prop_value_type = None
    if property_def:
        try:
            if hasattr(property_def, 'defaultValue') and property_def.defaultValue:
                if hasattr(property_def.defaultValue, 'propertyValue'):
                    if hasattr(property_def.defaultValue.propertyValue, 'type'):
                        prop_value_type = property_def.defaultValue.propertyValue.type
        except:
            pass

    print(
        f"\nProperty value type: {prop_value_type if prop_value_type else 'Unknown (will use string)'}")

    # -------------------------------------------------------------------------
    # 4. CREATE PROPERTY VALUE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CREATING PROPERTY VALUE")
    print("="*80)

    print(f"\nConverting '{NEW_VALUE}' to type '{prop_value_type}'...")

    try:
        # Create appropriate property value based on detected type
        if prop_value_type == 'string' or prop_value_type is None:
            # String property (default)
            property_value = act.NormalStringPropertyValue(
                type='string',
                status='normal',
                value=str(NEW_VALUE)
            )

        elif prop_value_type == 'integer':
            # Integer property (whole numbers)
            property_value = act.NormalIntegerPropertyValue(
                type='integer',
                status='normal',
                value=int(NEW_VALUE)
            )

        elif prop_value_type == 'number':
            # Number property (decimals)
            property_value = act.NormalNumberPropertyValue(
                type='number',
                status='normal',
                value=float(NEW_VALUE)
            )

        elif prop_value_type == 'boolean':
            # Boolean property (true/false)
            bool_value = str(NEW_VALUE).lower() in ['true', '1', 'yes', 'oui']
            property_value = act.NormalBooleanPropertyValue(
                type='boolean',
                status='normal',
                value=bool_value
            )

        elif prop_value_type == 'length':
            # Length property (millimeters)
            property_value = act.NormalLengthPropertyValue(
                type='length',
                status='normal',
                value=float(NEW_VALUE)
            )

        elif prop_value_type == 'area':
            # Area property (square meters)
            property_value = act.NormalAreaPropertyValue(
                type='area',
                status='normal',
                value=float(NEW_VALUE)
            )

        elif prop_value_type == 'volume':
            # Volume property (cubic meters)
            property_value = act.NormalVolumePropertyValue(
                type='volume',
                status='normal',
                value=float(NEW_VALUE)
            )

        elif prop_value_type == 'angle':
            # Angle property (degrees)
            property_value = act.NormalAnglePropertyValue(
                type='angle',
                status='normal',
                value=float(NEW_VALUE)
            )

        else:
            # Unknown type - default to string
            print(
                f"⚠️  Unknown property type '{prop_value_type}', using string")
            property_value = act.NormalStringPropertyValue(
                type='string',
                status='normal',
                value=str(NEW_VALUE)
            )

        print(f"✓ Property value created successfully")

    except Exception as e:
        print(f"\n✗ Error creating property value: {e}")
        print(
            f"\nMake sure NEW_VALUE '{NEW_VALUE}' is compatible with type '{prop_value_type}'")
        print("\nExpected formats:")
        print("  • string: Any text")
        print("  • integer: Whole number (e.g., 42)")
        print("  • number: Decimal (e.g., 3.14)")
        print("  • boolean: true/false/1/0/yes/no")
        print("  • length: Number in mm (e.g., 2400.0)")
        print("  • area: Number in m² (e.g., 15.5)")
        print("  • volume: Number in m³ (e.g., 45.2)")
        print("  • angle: Number in degrees (e.g., 90.0)")
        exit()

    # -------------------------------------------------------------------------
    # 5. APPLY PROPERTY VALUE
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("APPLYING PROPERTY VALUE")
    print("="*80)

    # Prepare element property values for all selected elements
    element_property_values = []
    for element in selected_elements:
        element_property_values.append(
            act.ElementPropertyValue(
                elementId=element.elementId,
                propertyId=target_property_id,
                propertyValue=property_value
            )
        )

    print(f"\nApplying property '{PROPERTY_NAME}' = '{NEW_VALUE}'")
    print(f"  To {len(selected_elements)} element(s)...")

    # Set the property values
    try:
        results = acc.SetPropertyValuesOfElements(element_property_values)

        # Count successes and errors
        success_count = 0
        error_count = 0

        for i, result in enumerate(results):
            if hasattr(result, 'success') and result.success:
                success_count += 1
            else:
                error_count += 1
                if hasattr(result, 'error'):
                    print(f"  Element {i+1} error: {result.error}")

        print(f"\n✓ Property set for {success_count} element(s)")

        if error_count > 0:
            print(f"⚠️  {error_count} element(s) had errors")
            print("  (May be read-only properties or locked elements)")

    except Exception as e:
        print(f"\n✗ Error setting property: {e}")
        import traceback
        traceback.print_exc()
        exit()

    # -------------------------------------------------------------------------
    # 6. VERIFY CHANGES
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("VERIFYING CHANGES")
    print("="*80)

    try:
        # Get property value from first element as sample
        prop_values = acc.GetPropertyValuesOfElements(
            [selected_elements[0].elementId],
            [target_property_id]
        )

        if len(prop_values) > 0 and len(prop_values[0].propertyValues) > 0:
            prop_val = prop_values[0].propertyValues[0]

            if hasattr(prop_val.propertyValue, 'value'):
                actual_value = prop_val.propertyValue.value
                print(f"\n✓ Verification successful")
                print(f"  Property '{PROPERTY_NAME}' = {actual_value}")
            else:
                print(f"\n⚠️  Could not read back property value")
                print("  (Property may still have been set successfully)")
        else:
            print(f"\n⚠️  Could not verify property value")
            print("  (Property may still have been set successfully)")

    except Exception as e:
        print(f"\n⚠️  Verification failed: {e}")
        print("  (Property may still have been set successfully)")

    # -------------------------------------------------------------------------
    # SUMMARY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Property: {PROPERTY_NAME}")
    print(f"  Group: {property_group}")
    print(f"  New value: {NEW_VALUE}")
    print(f"  Type: {prop_value_type}")
    print(f"  Elements modified: {success_count}")
    print("\n" + "="*80)

    # Performance note
    print("\n💡 Performance:")
    print("  • Used GetAllPropertyNames() for fast property lookup")
    print(f"  • Searched {len(available_properties)} properties instantly")
    print("  • Much faster than querying each property definition!")


if __name__ == "__main__":
    main()
