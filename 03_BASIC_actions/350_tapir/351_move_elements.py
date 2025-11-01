"""
Script: Move Elements (CORRECT FORMAT)
API: Tapir Add-On v1.0.2+
Description: Moves selected elements by specified offset
Usage:
    1. Select elements in Archicad
    2. Set movement offset below
    3. Run script
Requirements:
    - archicad-api package
    - Tapir Add-On v1.0.2+ installed
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

# === CONFIGURATION ===
# Movement offset in meters
MOVE_X = 5.0   # Move 5 meters in X direction
MOVE_Y = 0.0   # No movement in Y
MOVE_Z = 0.0   # No movement in Z (vertical)

# Copy elements instead of moving them?
COPY_INSTEAD_OF_MOVE = False

print(f"\n{'='*60}")
print(f"{'COPYING' if COPY_INSTEAD_OF_MOVE else 'MOVING'} ELEMENTS")
print(f"{'='*60}")
print(f"Elements selected: {len(selected_elements)}")
print(f"Offset: X={MOVE_X}m, Y={MOVE_Y}m, Z={MOVE_Z}m")

# Get element types for display
element_ids = [elem.elementId for elem in selected_elements]
element_types = acc.GetTypesOfElements(element_ids)

print("\nElements to move:")
for i, elem_type in enumerate(element_types[:10]):  # Show first 10
    elem_type_name = elem_type.typeOfElement.elementType
    print(f"  {i+1}. {elem_type_name}")
if len(element_types) > 10:
    print(f"  ... and {len(element_types) - 10} more")

try:
    # Build elementsWithMoveVectors array according to Tapir schema
    elements_with_vectors = []

    for elem in selected_elements:
        elements_with_vectors.append({
            'elementId': {
                'guid': str(elem.elementId.guid)  # Convert UUID to string
            },
            'moveVector': {
                'x': MOVE_X,
                'y': MOVE_Y,
                'z': MOVE_Z
            },
            'copy': COPY_INSTEAD_OF_MOVE  # Optional parameter
        })

    # Move elements using Tapir command with CORRECT format
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'MoveElements'),
        {
            'elementsWithMoveVectors': elements_with_vectors
        }
    )

    print(f"\n{'='*60}")
    print(f"✓ SUCCESS")
    print(f"{'='*60}")

    if COPY_INSTEAD_OF_MOVE:
        print(f"Created {len(selected_elements)} copy/copies")
        print(f"Offset by: ({MOVE_X}, {MOVE_Y}, {MOVE_Z})")
    else:
        print(f"Moved {len(selected_elements)} element(s)")
        print(f"New position offset by: ({MOVE_X}, {MOVE_Y}, {MOVE_Z})")

    # Check for any failures in results
    if 'executionResults' in response:
        failed = sum(
            1 for r in response['executionResults'] if not r.get('success', True))
        if failed > 0:
            print(f"\n⚠ Warning: {failed} element(s) failed to move")

    print("\n💡 Tips:")
    print("  • Undo in Archicad if needed (Ctrl+Z / Cmd+Z)")
    print("  • Use negative values to move in opposite direction")
    print("  • Z offset moves elements vertically")
    if not COPY_INSTEAD_OF_MOVE:
        print("  • Set COPY_INSTEAD_OF_MOVE = True to duplicate instead")

except Exception as e:
    print(f"\n{'='*60}")
    print(f"✗ ERROR")
    print(f"{'='*60}")
    print(f"{e}")
    print("\nPossible causes:")
    print("  • Tapir Add-On is not installed or not loaded")
    print("  • Tapir Add-On version is too old (needs v1.0.2+)")
    print("  • Elements are locked")
    print("  • Movement values are invalid")
    print("\nTo install/update Tapir:")
    print("  https://github.com/ENZYME-APD/tapir-archicad-automation/releases")

# Examples
print("\n" + "="*60)
print("MOVEMENT EXAMPLES")
print("="*60)
print("Move 10m to the right:       MOVE_X = 10.0")
print("Move 5m forward:             MOVE_Y = 5.0")
print("Move 3m up (vertical):       MOVE_Z = 3.0")
print("Move back and down:          MOVE_X = -5.0, MOVE_Y = -3.0")
print("Copy 2m to the right:        MOVE_X = 2.0, COPY_INSTEAD_OF_MOVE = True")
print("\nCoordinate system:")
print("  X = Left(-) / Right(+)")
print("  Y = Back(-) / Forward(+)")
print("  Z = Down(-) / Up(+)")
