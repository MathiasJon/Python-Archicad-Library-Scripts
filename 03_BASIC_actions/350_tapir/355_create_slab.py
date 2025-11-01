"""
Script: Create Slab
API: Tapir Add-On
Description: Creates a rectangular slab at specified location
Usage: Modify dimensions and position, then run
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
# Slab dimensions
SLAB_WIDTH = 10.0   # meters (X direction)
SLAB_DEPTH = 8.0    # meters (Y direction)
SLAB_ELEVATION = 0.0  # meters (Z coordinate)

# Slab corner position (bottom-left corner)
START_X = 0.0
START_Y = 0.0

print(f"\n=== CREATING SLAB ===")
print(f"Dimensions: {SLAB_WIDTH}m x {SLAB_DEPTH}m")
print(f"Elevation: {SLAB_ELEVATION}m")
print(f"Position: ({START_X}, {START_Y})")

# Define slab polygon (rectangle)
# Points must be in counter-clockwise order
polygon = [
    {'x': START_X,              'y': START_Y},              # Bottom-left
    {'x': START_X + SLAB_WIDTH, 'y': START_Y},              # Bottom-right
    {'x': START_X + SLAB_WIDTH, 'y': START_Y + SLAB_DEPTH},  # Top-right
    {'x': START_X,              'y': START_Y + SLAB_DEPTH}  # Top-left
]

try:
    # Create slab using Tapir command
    # Format data according to API schema: slabsData array with objects containing level and polygonCoordinates
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'CreateSlabs'),
        {
            'slabsData': [
                {
                    'level': SLAB_ELEVATION,
                    'polygonCoordinates': polygon
                }
            ]
        }
    )

    print(f"\n✓ Successfully created slab")
    print("\nSlab corners:")
    for i, point in enumerate(polygon):
        print(f"  {i+1}. X: {point['x']:.2f}, Y: {point['y']:.2f}")

    print("\nNote: Slab created with default settings")
    print("Use Archicad's UI to modify slab properties")

except Exception as e:
    print(f"\n✗ Error creating slab: {e}")
    print("\nMake sure:")
    print("  1. Tapir Add-On is installed")
    print("  2. Polygon coordinates are valid")
    print("  3. Polygon has at least 3 points")

# Example: Create multiple slabs
print("\n=== EXAMPLE: Multiple Slabs ===")
print("To create multiple slabs at once:")
print("""
response = acc.ExecuteAddOnCommand(
    act.AddOnCommandId('TapirCommand', 'CreateSlabs'),
    {
        'slabsData': [
            {
                'level': 0.0,
                'polygonCoordinates': [
                    {'x': 0, 'y': 0}, 
                    {'x': 5, 'y': 0}, 
                    {'x': 5, 'y': 5}, 
                    {'x': 0, 'y': 5}
                ]
            },
            {
                'level': 0.0,
                'polygonCoordinates': [
                    {'x': 10, 'y': 0}, 
                    {'x': 15, 'y': 0}, 
                    {'x': 15, 'y': 5}, 
                    {'x': 10, 'y': 5}
                ]
            }
        ]
    }
)
""")

# Example: Slab with a hole
print("\n=== EXAMPLE: Slab with Hole ===")
print("To create a slab with a circular opening:")
print("""
response = acc.ExecuteAddOnCommand(
    act.AddOnCommandId('TapirCommand', 'CreateSlabs'),
    {
        'slabsData': [
            {
                'level': 0.0,
                'polygonCoordinates': [
                    {'x': 0, 'y': 0}, 
                    {'x': 10, 'y': 0}, 
                    {'x': 10, 'y': 10}, 
                    {'x': 0, 'y': 10}
                ],
                'holes': [
                    {
                        'polygonOutline': [
                            {'x': 4, 'y': 5},
                            {'x': 6, 'y': 5},
                            {'x': 6, 'y': 7},
                            {'x': 4, 'y': 7}
                        ]
                    }
                ]
            }
        ]
    }
)
""")
