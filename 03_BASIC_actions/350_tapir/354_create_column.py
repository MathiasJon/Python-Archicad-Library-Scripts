"""
Script: Create Column
API: Tapir Add-On
Description: Creates a simple column at specified coordinates
Usage: Modify coordinates and run to create columns
Requirements:
    - archicad-api package
    - Tapir Add-On installed
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# === CONFIGURATION ===
# Define column positions (x, y, z coordinates in meters)
# Z coordinate is the base elevation

# Example: Create a 3x3 grid of columns
COLUMN_POSITIONS = []
for x in range(3):
    for y in range(3):
        COLUMN_POSITIONS.append({
            'x': x * 5.0,    # 5 meter spacing in X
            'y': y * 5.0,    # 5 meter spacing in Y
            'z': 0.0         # Ground level
        })

# Alternatively, define specific positions:
# COLUMN_POSITIONS = [
#     {'x': 0.0, 'y': 0.0, 'z': 0.0},
#     {'x': 5.0, 'y': 0.0, 'z': 0.0},
#     {'x': 10.0, 'y': 0.0, 'z': 0.0},
# ]

print(f"\n=== CREATING COLUMNS ===")
print(f"Number of columns to create: {len(COLUMN_POSITIONS)}")

try:
    # Create columns using Tapir command
    # Format data according to API schema: columnsData array with coordinates objects
    columns_data = [{'coordinates': pos} for pos in COLUMN_POSITIONS]

    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'CreateColumns'),
        {
            'columnsData': columns_data
        }
    )

    print(f"\n✓ Successfully created {len(COLUMN_POSITIONS)} column(s)")

    print("\nColumn positions:")
    for i, pos in enumerate(COLUMN_POSITIONS):
        print(
            f"  {i+1}. X: {pos['x']:.2f}, Y: {pos['y']:.2f}, Z: {pos['z']:.2f}")

    print("\nNote: Columns are created with default settings")
    print("Use Archicad's UI to modify column properties")

except Exception as e:
    print(f"\n✗ Error creating columns: {e}")
    print("\nMake sure:")
    print("  1. Tapir Add-On is installed")
    print("  2. Coordinates are valid")

# Example: Create a single column at origin
print("\n=== EXAMPLE: Single Column at Origin ===")
print("To create one column at (0,0,0):")
print("""
response = acc.ExecuteAddOnCommand(
    act.AddOnCommandId('TapirCommand', 'CreateColumns'),
    {
        'columnsData': [
            {
                'coordinates': {'x': 0.0, 'y': 0.0, 'z': 0.0}
            }
        ]
    }
)
""")
