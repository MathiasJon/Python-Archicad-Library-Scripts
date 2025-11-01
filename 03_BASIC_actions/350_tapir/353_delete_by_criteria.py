"""
Script: Delete by Property - Improved (CORRECT FORMAT)
API: Tapir Add-On v1.2.1+ & Official API
Description: Improved property filtering that handles enums, strings, etc.
Usage:
    1. Run 107_diagnose_property_values.py first to check values
    2. Configure property filter below
    3. Enable deletion
    4. Run script
    ⚠  WARNING: This permanently deletes elements!
Requirements:
    - archicad-api package
    - Tapir Add-On v1.2.1+ installed
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# === CONFIGURATION ===

# Property to filter by
PROPERTY_NAME = "Status"

# Value to match - can be:
# - String: "To Delete"
# - Number: 123
# - Boolean: True/False
PROPERTY_VALUE = "To Delete"

# Match mode for string comparison
MATCH_MODE = "EXACT"  # Options: "EXACT", "CONTAINS", "STARTS_WITH", "ENDS_WITH"

# Case sensitive?
CASE_SENSITIVE = False

# SAFETY: Set to True to enable deletion
ENABLE_DELETION = True

# =======================================

print(f"\n{'='*70}")
print(f"DELETE ELEMENTS BY PROPERTY (IMPROVED)")
print(f"{'='*70}")
print(f"Property: {PROPERTY_NAME}")
print(f"Looking for value: {repr(PROPERTY_VALUE)}")
print(f"Match mode: {MATCH_MODE}")
print(f"Case sensitive: {CASE_SENSITIVE}")

# Get all elements
all_elements = acc.GetAllElements()
print(f"\nTotal elements in project: {len(all_elements)}")

# Get all properties
all_property_ids = acc.GetAllPropertyIds()
prop_details = acc.GetDetailsOfProperties(all_property_ids)

# Find the target property
target_prop_id = None
target_prop_detail = None

for prop_detail in prop_details:
    if prop_detail.propertyDefinition.name == PROPERTY_NAME:
        target_prop_id = prop_detail.propertyDefinition.propertyId
        target_prop_detail = prop_detail
        break

if not target_prop_id:
    print(f"\n✗ Property '{PROPERTY_NAME}' not found")
    print("\nAvailable properties (first 20):")
    for i, prop_detail in enumerate(prop_details[:20]):
        print(f"  {i+1}. {prop_detail.propertyDefinition.name}")
    exit()

print(f"\n✓ Property found")
print(f"  Type: {target_prop_detail.propertyDefinition.type}")

# Check if enum
is_enum = target_prop_detail.propertyDefinition.possibleEnumValues is not None
if is_enum:
    print(f"  Property is an ENUM with values:")
    for enum_val in target_prop_detail.propertyDefinition.possibleEnumValues:
        print(f"    • {enum_val.enumValue.displayValue}")

# Helper function to extract value from property


def extract_value(prop_value):
    """Extract the actual value from a PropertyValue object"""
    if not hasattr(prop_value, 'value'):
        return None

    value = prop_value.value

    # If it's an enum or object with displayValue
    if hasattr(value, 'displayValue'):
        return value.displayValue

    # If it's an object with nonLocalizedValue
    if hasattr(value, 'nonLocalizedValue'):
        return value.nonLocalizedValue

    # Otherwise return as-is
    return value

# Helper function to check if value matches


def value_matches(actual_value, target_value, mode="EXACT", case_sensitive=False):
    """Check if actual value matches target value"""
    if actual_value is None:
        return False

    # Convert to strings for comparison
    actual_str = str(actual_value)
    target_str = str(target_value)

    # Handle case sensitivity
    if not case_sensitive:
        actual_str = actual_str.lower()
        target_str = target_str.lower()

    # Match based on mode
    if mode == "EXACT":
        return actual_str == target_str
    elif mode == "CONTAINS":
        return target_str in actual_str
    elif mode == "STARTS_WITH":
        return actual_str.startswith(target_str)
    elif mode == "ENDS_WITH":
        return actual_str.endswith(target_str)
    else:
        return actual_str == target_str


# Get property values for all elements
print(f"\n🔍 Scanning {len(all_elements)} elements...")

property_values = acc.GetPropertyValuesOfElements(
    all_elements,
    [act.PropertyIdArrayItem(propertyId=target_prop_id)]
)

# Filter elements
elements_to_delete = []
match_details = []

for i, (elem, prop_val_result) in enumerate(zip(all_elements, property_values)):
    try:
        # Extract property value
        if hasattr(prop_val_result, 'propertyValues') and prop_val_result.propertyValues:
            prop_val = prop_val_result.propertyValues[0]

            if hasattr(prop_val, 'propertyValue'):
                prop_value = prop_val.propertyValue

                # Extract the actual value
                actual_value = extract_value(prop_value)

                # Check if it matches
                if value_matches(actual_value, PROPERTY_VALUE, MATCH_MODE, CASE_SENSITIVE):
                    elements_to_delete.append(elem)
                    match_details.append({
                        'element': elem,
                        'value': actual_value
                    })

    except Exception as e:
        # Silently skip elements that cause errors
        pass

# Display results
print(f"\n{'='*70}")
print(f"DELETION PREVIEW")
print(f"{'='*70}")
print(f"Elements matching criteria: {len(elements_to_delete)}")

if len(elements_to_delete) == 0:
    print("\n✓ No elements match the criteria")
    print("\n💡 Tips:")
    print("  • Run 107_diagnose_property_values.py to check actual values")
    print("  • Try MATCH_MODE = 'CONTAINS' for partial matches")
    print("  • Set CASE_SENSITIVE = False for case-insensitive matching")
    exit()

# Get types of matched elements
element_ids = [elem.elementId for elem in elements_to_delete]
element_types = acc.GetTypesOfElements(element_ids)

# Count by type
type_counts = {}
for elem_type in element_types:
    type_name = elem_type.typeOfElement.elementType
    type_counts[type_name] = type_counts.get(type_name, 0) + 1

print("\nMatched elements by type:")
for elem_type, count in sorted(type_counts.items()):
    print(f"  • {elem_type}: {count}")

# Show first 10 with their values
print("\nFirst 10 matched elements:")
for i, (match, elem_type) in enumerate(zip(match_details[:10], element_types[:10])):
    elem_type_name = elem_type.typeOfElement.elementType
    elem_guid = str(match['element'].elementId.guid)
    value = match['value']
    print(
        f"  {i+1:3}. {elem_type_name:15} | Value: '{value}' | {elem_guid[:18]}...")

if len(elements_to_delete) > 10:
    print(f"  ... and {len(elements_to_delete) - 10} more")

# Safety check
print(f"\n{'='*70}")
print(f"⚠  CONFIRM DELETION")
print(f"{'='*70}")

if not ENABLE_DELETION:
    print("\n✋ DELETION DISABLED")
    print("\n💡 To proceed:")
    print("  1. Review the matched elements above")
    print("  2. Save your Archicad project!")
    print("  3. Set ENABLE_DELETION = True in the script")
    print("  4. Run again")
    exit()

# Build elements array
elements_array = []
for elem in elements_to_delete:
    elements_array.append({
        'elementId': {
            'guid': str(elem.elementId.guid)
        }
    })

try:
    # Delete elements
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'DeleteElements'),
        {
            'elements': elements_array
        }
    )

    print(f"\n{'='*70}")
    print(f"✓ DELETION COMPLETE")
    print(f"{'='*70}")
    print(f"Deleted {len(elements_to_delete)} element(s)")

    print("\nDeleted by type:")
    for elem_type, count in sorted(type_counts.items()):
        print(f"  ✓ {elem_type}: {count}")

    print(f"\n💡 Undo in Archicad: Ctrl+Z / Cmd+Z")

except Exception as e:
    print(f"\n{'='*70}")
    print(f"✗ ERROR")
    print(f"{'='*70}")
    print(f"{e}")

print(f"\n{'='*70}")
print("MATCH MODE EXAMPLES")
print(f"{'='*70}")
print("EXACT:       Value must match exactly")
print("CONTAINS:    Value contains the text anywhere")
print("STARTS_WITH: Value starts with the text")
print("ENDS_WITH:   Value ends with the text")
