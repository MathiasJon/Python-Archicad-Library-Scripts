"""
================================================================================
SCRIPT: List All Properties with Details
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Retrieval - Properties

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Lists all properties available in the current Archicad project with comprehensive
details in a compact, one-line-per-property format. This script provides:
- Complete list of all properties organized by groups
- Property types, editability, and descriptions
- Enumeration values for dropdown properties
- Default values configured for properties
- Sample values from actual elements in the project
- Statistical summary of property types and distribution

Properties are the metadata fields that can be attached to building elements
(walls, slabs, doors, etc.) and used for documentation, scheduling, and BIM workflows.

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

- GetAllElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllElements
  Returns all elements in the project for value sampling

- GetAllPropertyIds()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyIds
  Returns the list of all available property IDs

- GetDetailsOfProperties(propertyIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfProperties
  Returns detailed information about specified properties

- GetPropertyValuesOfElements(elements, propertyIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetPropertyValuesOfElements
  Returns property values for specified elements

[Data Types]
- PropertyId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.PropertyId
  Identifier for properties

- PropertyDefinition
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.PropertyDefinition
  Contains property metadata (name, type, group, default values, etc.)

- PropertyValue
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.PropertyValue
  Contains actual property values from elements

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. Review the compact property listing with details
4. Note property names and types for use in other scripts

No configuration needed - script automatically samples elements for values.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Loading information (elements and properties count)
- Properties organized by groups
- For each property: Number, Name, Type, Editability, Description, Enum Values,
  Default Value, and Sample Values from elements
- Summary with statistics and type distribution

Example output:
  ════════════════════════════════════════════════════════════
  📋 General ({count})
  ────────────────────────────────────────────────────────────
    1. Element ID                        | string       | ✓ | Unique identifier            | | Def:                | Ex: 12345, 67890
    2. Status                            | singleEnum   | ✓ | Current element status       | [New, Existing, +2]  | Def:New             | Ex: Existing, New

  📊 Type Distribution:
     • string        :  125 properties
     • number        :   89 properties
     • singleEnum    :   45 properties

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
PROPERTY GROUPS:
- Properties are organized into groups like "General", "IFC", "Custom", etc.
- Groups help organize properties logically
- Custom properties created by users appear in their defined groups

PROPERTY TYPES:
- string: Text values
- number: Numeric values (integer or decimal)
- length: Distance measurements
- area: Area measurements
- volume: Volume measurements
- angle: Angular measurements
- boolean: True/False values
- singleEnum: Single selection from dropdown list
- multiEnum: Multiple selections from list

EDITABILITY:
- ✓ = Property can be edited by users
- ✗ = Property is read-only (calculated or system-managed)

SAMPLE VALUES:
- Script samples up to 30 elements to show example values
- Shows up to 3 unique values per property
- Helps understand what data is actually in the project
- <empty> indicates no values found in sampled elements
- Long values are truncated with "..." for readability

ENUM VALUES:
- Shows available options for dropdown properties
- Limited to first 3 options with count of additional options
- Format: [Option1, Option2, Option3, +5] means 3 shown + 5 more

DEFAULT VALUES:
- Shows the configured default value for the property
- Empty if no default is configured
- Useful for understanding property initialization

PERFORMANCE:
- Fast for typical projects (< 1000 properties)
- Samples only 30 elements for value examples
- Full property details retrieved for all properties
- May take a few seconds for large projects with many properties

LIMITATIONS:
- Sample values based on first 30 elements only
- May not represent full range of values in large projects
- Property values depend on element types (not all properties apply to all elements)

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 103_get_all_properties.py (basic property listing)
- 104_get_property_value.py (get specific property value)
- 202_set_property_value.py (set property values)
- 111_get_element_full_info.py (view element with all properties)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print("\n" + "="*180)
print("PROPERTIES: NAME | TYPE | EDIT | DESCRIPTION | ENUM VALUES | DEFAULT | SAMPLE VALUES")
print("="*180)

# Step 1: Get elements for value sampling
print("\n🔍 Loading data...")
all_elements = acc.GetAllElements()
sample_elements = all_elements[:30]  # Sample 30 elements
print(f"   Elements: {len(sample_elements)}/{len(all_elements)}")

# Step 2: Get all properties
all_property_ids = acc.GetAllPropertyIds()
all_property_details = acc.GetDetailsOfProperties(all_property_ids)
print(f"   Properties: {len(all_property_details)}")

# Step 3: Organize by group
properties_by_group = {}
for prop_detail in all_property_details:
    prop_def = prop_detail.propertyDefinition
    group_name = prop_def.group.name if hasattr(
        prop_def, 'group') and prop_def.group else "No Group"

    if group_name not in properties_by_group:
        properties_by_group[group_name] = []
    properties_by_group[group_name].append(prop_detail)

print(f"   Groups: {len(properties_by_group)}\n")

# Helper function to get property type


def get_property_type(prop_def):
    """Extract the property type directly from PropertyDefinition.type"""

    # The type is directly available as an attribute
    if hasattr(prop_def, 'type') and prop_def.type:
        return prop_def.type

    return 'unknown'

# Helper function to get sample values


def get_sample_values(prop_def, sample_elements):
    """Get up to 3 sample values from elements"""
    try:
        property_ids = [act.PropertyIdArrayItem(
            propertyId=prop_def.propertyId)]
        property_values = acc.GetPropertyValuesOfElements(
            sample_elements, property_ids)

        unique_values = set()
        for prop_val_or_error in property_values:
            if hasattr(prop_val_or_error, 'propertyValues'):
                for prop_val in prop_val_or_error.propertyValues:
                    if hasattr(prop_val, 'propertyValue') and hasattr(prop_val.propertyValue, 'value'):
                        value = prop_val.propertyValue.value

                        # Convert value to string
                        if value is None:
                            continue  # Skip empty values
                        elif isinstance(value, (str, int, float, bool)):
                            value_str = str(value)
                        elif hasattr(value, 'value'):
                            value_str = str(value.value)
                        else:
                            value_str = str(type(value).__name__)

                        # Truncate long values
                        if len(value_str) > 30:
                            value_str = value_str[:27] + "..."

                        unique_values.add(value_str)

                        if len(unique_values) >= 3:  # Limit to 3 values
                            break

            if len(unique_values) >= 3:
                break

        if unique_values:
            return ", ".join(sorted(unique_values))
        else:
            return "<empty>"

    except Exception as e:
        return f"<error>"


# Step 4: Display results in compact format
print("="*180)
total_count = 0

for group_name in sorted(properties_by_group.keys()):
    properties = properties_by_group[group_name]

    print(f"\n📁 {group_name} ({len(properties)})")
    print("─"*180)

    for i, prop_detail in enumerate(properties, 1):
        prop_def = prop_detail.propertyDefinition
        prop_name = prop_def.name
        prop_type = get_property_type(prop_def)
        sample_values = get_sample_values(prop_def, sample_elements)

        # Get additional properties
        description = prop_def.description if prop_def.description else ""
        if len(description) > 30:
            description = description[:27] + "..."

        editable = "✓" if prop_def.isEditable else "✗"

        # Get enum values
        enum_values = ""
        if prop_def.possibleEnumValues and len(prop_def.possibleEnumValues) > 0:
            enum_list = [
                ev.enumValue.displayValue for ev in prop_def.possibleEnumValues[:3]]
            enum_values = f"[{', '.join(enum_list)}]"
            if len(prop_def.possibleEnumValues) > 3:
                enum_values = enum_values[:-1] + \
                    f", +{len(prop_def.possibleEnumValues)-3}]"

        # Get default value
        # Structure: defaultValue -> basicDefaultValue -> value
        default_val = ""
        if prop_def.defaultValue:
            try:
                if hasattr(prop_def.defaultValue, 'basicDefaultValue'):
                    basic_default = prop_def.defaultValue.basicDefaultValue
                    if hasattr(basic_default, 'value'):
                        val = basic_default.value
                        if val is not None:
                            # Handle different value types
                            if isinstance(val, (str, int, float, bool)):
                                default_val = str(val)
                            elif hasattr(val, 'value'):
                                default_val = str(val.value)
                            else:
                                default_val = str(val)
                            
                            # Truncate if too long
                            if len(default_val) > 20:
                                default_val = default_val[:17] + "..."
            except (AttributeError, TypeError):
                default_val = ""

        # Format: number. name | type | edit | desc | enum | default | sample values
        # Truncate name if too long
        if len(prop_name) > 35:
            prop_name = prop_name[:32] + "..."

        print(f"{i:3}. {prop_name:38} | {prop_type:12} | {editable} | {description:32} | {enum_values:25} | Def:{default_val:15} | Ex: {sample_values}")
        total_count += 1

# Summary
print("\n" + "="*180)
print("SUMMARY")
print("="*180)
print(f"Groups: {len(properties_by_group)} | Properties: {total_count} | Elements sampled: {len(sample_elements)}/{len(all_elements)}")

# Type distribution
print("\n📊 Type Distribution:")
type_counts = {}
for group_name, properties in properties_by_group.items():
    for prop_detail in properties:
        prop_type = get_property_type(prop_detail.propertyDefinition)
        type_counts[prop_type] = type_counts.get(prop_type, 0) + 1

for prop_type, count in sorted(type_counts.items(), key=lambda x: -x[1]):
    print(f"   • {prop_type:15}: {count:4} properties")

print("\n" + "="*180)
print("✓ Complete")