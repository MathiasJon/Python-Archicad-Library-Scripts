"""
Script: Delete Elements (CORRECT FORMAT)
API: Tapir Add-On v1.2.1+
Description: Deletes selected elements from the project
Usage:
    1. Select elements to delete in Archicad
    2. Enable deletion in code (safety measure)
    3. Run script
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

# Get selected elements
selected_elements = acc.GetSelectedElements()

if len(selected_elements) == 0:
    print("\n⚠  No elements selected")
    print("Please select elements in Archicad and run again")
    exit()

print("\n" + "="*60)
print("⚠  WARNING: ELEMENT DELETION")
print("="*60)

# Get element types for display
element_ids = [elem.elementId for elem in selected_elements]
element_types = acc.GetTypesOfElements(element_ids)

# Count by type
type_counts = {}
for elem_type in element_types:
    type_name = elem_type.typeOfElement.elementType
    type_counts[type_name] = type_counts.get(type_name, 0) + 1

print(f"\nElements selected for deletion: {len(selected_elements)}")
print("\nElements by type:")
for elem_type, count in sorted(type_counts.items()):
    print(f"  {elem_type}: {count}")

# Show first 10 elements
print("\nFirst 10 elements:")
for i, (elem, elem_type) in enumerate(zip(selected_elements[:10], element_types[:10])):
    elem_type_name = elem_type.typeOfElement.elementType
    elem_guid = str(elem.elementId.guid)
    print(f"  {i+1}. {elem_type_name:15} | {elem_guid}")

if len(selected_elements) > 10:
    print(f"  ... and {len(selected_elements) - 10} more")

# Confirmation prompt
print("\n" + "="*60)
print("⚠  THIS ACTION WILL DELETE ELEMENTS!")
print("="*60)
print("\nTo proceed with deletion:")
print("  1. Save your Archicad project first!")
print("  2. Set ENABLE_DELETION = True below")
print("  3. Run the script again")

# === SAFETY: Set to True to enable deletion ===
ENABLE_DELETION = True  # Change to True to enable

if not ENABLE_DELETION:
    print("\n✋ DELETION DISABLED")
    print("\n💡 Set ENABLE_DELETION = True in the script to proceed")
    exit()

# Build elements array with CORRECT structure according to Tapir schema
elements_array = []
for elem in selected_elements:
    elements_array.append({
        'elementId': {
            'guid': str(elem.elementId.guid)  # Convert UUID to string
        }
    })

try:
    # Delete elements using Tapir command with CORRECT format
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'DeleteElements'),
        {
            'elements': elements_array  # NOT 'elementIds'!
        }
    )

    print(f"\n{'='*60}")
    print(f"✓ DELETION COMPLETE")
    print(f"{'='*60}")
    print(f"Deleted {len(selected_elements)} element(s)")

    print("\nDeleted elements by type:")
    for elem_type, count in sorted(type_counts.items()):
        print(f"  ✓ {elem_type}: {count}")

    print("\n💡 Use Archicad's Undo (Ctrl+Z / Cmd+Z) to restore if needed")

    # Check response
    if isinstance(response, dict):
        if response.get('success'):
            print("\n✓ All elements deleted successfully")
        elif 'error' in response:
            print(f"\n⚠ Deletion completed with errors:")
            print(f"  Error code: {response['error'].get('code')}")
            print(f"  Message: {response['error'].get('message')}")

except Exception as e:
    print(f"\n{'='*60}")
    print(f"✗ ERROR")
    print(f"{'='*60}")
    print(f"{e}")
    print("\nPossible causes:")
    print("  • Tapir Add-On is not installed or not loaded")
    print("  • Tapir Add-On version is too old (needs v1.2.1+)")
    print("  • Elements are locked or protected")
    print("  • Elements are in a locked layer")
    print("  • Teamwork reservation required")
    print("\nTo install/update Tapir:")
    print("  https://github.com/ENZYME-APD/tapir-archicad-automation/releases")

print("\n" + "="*60)
print("SAFETY TIPS")
print("="*60)
print("• ALWAYS save your project before deleting")
print("• Test on a copy of the project first")
print("• Use Archicad's Undo immediately if needed")
print("• Consider hiding elements instead of deleting")
print("• Keep backups of important projects")
print("• Use version control for team projects")
