"""
Script: Create Zone
API: Tapir Add-On
Description: Creates a zone with specified boundary polygon
Usage: Define zone boundary coordinates and run
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
# Zone boundary polygon (counter-clockwise)
# Coordinates in meters (x, y)
ZONE_POLYGON = [
    {'x': 0.0, 'y': 0.0},     # Bottom-left
    {'x': 10.0, 'y': 0.0},    # Bottom-right
    {'x': 10.0, 'y': 8.0},    # Top-right
    {'x': 0.0, 'y': 8.0}      # Top-left
]

# Zone name (required)
ZONE_NAME = "Conference Room"

# Zone number (required)
ZONE_NUMBER = "101"

# Floor index (0 = ground floor, 1 = first floor, etc.)
FLOOR_INDEX = 0

# Position for zone stamp/label (center of your zone)
STAMP_X = 5.0
STAMP_Y = 4.0

print(f"\n=== CREATING ZONE ===")
print(f"Name: {ZONE_NAME}")
print(f"Number: {ZONE_NUMBER}")
print(f"Floor Index: {FLOOR_INDEX}")

try:
    # Create zone using Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'CreateZones'),
        {
            'zonesData': [
                {
                    'name': ZONE_NAME,
                    'numberStr': ZONE_NUMBER,
                    'floorIndex': FLOOR_INDEX,
                    'stampPosition': {
                        'x': STAMP_X,
                        'y': STAMP_Y
                    },
                    'geometry': {
                        'polygonCoordinates': ZONE_POLYGON
                    }
                }
            ]
        }
    )

    print(f"\n✓ Successfully created zone: {ZONE_NAME}")

except Exception as e:
    print(f"\n✗ Error creating zone: {e}")
    print("\nPossible causes:")
    print("  1. Tapir Add-On not installed")
    print("  2. Invalid polygon coordinates")
    print("  3. Floor index not found")

print("\n" + "="*60)
print("SIMPLE EXAMPLE")
print("="*60)
print("""
response = acc.ExecuteAddOnCommand(
    act.AddOnCommandId('TapirCommand', 'CreateZones'),
    {
        'zonesData': [
            {
                'name': 'Office',
                'numberStr': '102',
                'floorIndex': 0,
                'stampPosition': {'x': 2.5, 'y': 1.5},
                'geometry': {
                    'polygonCoordinates': [
                        {'x': 0.0, 'y': 0.0},
                        {'x': 5.0, 'y': 0.0},
                        {'x': 5.0, 'y': 3.0},
                        {'x': 0.0, 'y': 3.0}
                    ]
                }
            }
        ]
    }
)
""")
