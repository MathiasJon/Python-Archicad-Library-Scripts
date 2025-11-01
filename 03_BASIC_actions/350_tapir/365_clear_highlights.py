"""
Script: Clear All Highlights (CORRECT FORMAT)
API: Tapir Add-On v1.0.3+
Description: Removes all element highlights from Archicad
Usage: Simply run this script to clear all highlights
Requirements:
    - archicad-api package
    - Tapir Add-On v1.0.3+ installed
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print(f"\n{'='*60}")
print(f"CLEARING ALL HIGHLIGHTS")
print(f"{'='*60}")

try:
    # According to Tapir schema: empty elements array removes all highlights
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'HighlightElements'),
        {
            'elements': [],  # Empty array clears all highlights
            'highlightedColors': [],
            'wireframe3D': False
        }
    )

    print("✓ All highlights cleared successfully")

    if response.get('success'):
        print("\nAll elements returned to normal display")

except Exception as e:
    print(f"✗ Error clearing highlights: {e}")
    print("\nPossible causes:")
    print("  • Tapir Add-On is not installed or not loaded")
    print("  • Tapir Add-On version is too old (needs v1.0.3+)")
    print("  • Archicad version incompatibility")
    print("\nTo install/update Tapir:")
    print("  https://github.com/ENZYME-APD/tapir-archicad-automation/releases")

print(f"\n{'='*60}")
