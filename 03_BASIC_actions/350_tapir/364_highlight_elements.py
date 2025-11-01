"""
Script: Highlight Elements (CORRECT FORMAT)
API: Tapir Add-On v1.0.3+
Description: Highlights elements in Archicad with the correct schema format
Usage: 
    1. Select elements you want to highlight
    2. Run script
    3. Elements will be highlighted in Archicad
Requirements:
    - archicad-api package
    - Tapir Add-On v1.0.3+ installed
"""

from archicad import ACConnection

# ============ CONFIGURATION ============
# Highlight color (RGBA: Red, Green, Blue, Alpha)
# Values range from 0 to 255

HIGHLIGHT_COLOR = [255, 100, 100, 255]  # Red (default)

# Wireframe mode for non-highlighted elements in 3D
WIREFRAME_3D = False

# Optional: Color for non-highlighted elements (uncomment to use)
NON_HIGHLIGHTED_COLOR = [200, 200, 200, 100]  # Gray semi-transparent

# =======================================

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

print(f"\n{'='*60}")
print(f"HIGHLIGHTING ELEMENTS")
print(f"{'='*60}")
print(f"Elements to highlight: {len(selected_elements)}")
print(f"Color (RGBA): {HIGHLIGHT_COLOR}")
print(f"Wireframe 3D: {WIREFRAME_3D}")

# Get element types for display
element_types = acc.GetTypesOfElements(
    [elem.elementId for elem in selected_elements])

print("\nElements:")
for i, elem_type in enumerate(element_types):
    elem_type_name = elem_type.typeOfElement.elementType
    elem_guid = str(selected_elements[i].elementId.guid)
    print(f"  {i+1}. {elem_type_name:15} | GUID: {elem_guid}")

try:
    # Build elements array with CORRECT structure according to Tapir schema
    # Each element must be: { elementId: { guid: "guid-string" } }
    elements_array = []
    for elem in selected_elements:
        elements_array.append({
            'elementId': {
                'guid': str(elem.elementId.guid)
            }
        })

    # Build the parameters with CORRECT schema
    params = {
        'elements': elements_array,
        'highlightedColors': [HIGHLIGHT_COLOR] * len(selected_elements),
        'wireframe3D': WIREFRAME_3D
    }

    # Add optional non-highlighted color if defined
    if 'NON_HIGHLIGHTED_COLOR' in globals():
        params['nonHighlightedColor'] = NON_HIGHLIGHTED_COLOR

    # Execute the Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'HighlightElements'),
        params
    )

    print(f"\n{'='*60}")
    print(f"✓ SUCCESS")
    print(f"{'='*60}")
    print(f"Highlighted {len(selected_elements)} element(s) in Archicad")

    if response.get('success'):
        print("\nElements are now highlighted!")

    print(f"\n💡 To clear highlights:")
    print("  • Run: 95_clear_highlights.py")
    print("  • Or select no elements and run this script")

except Exception as e:
    print(f"\n{'='*60}")
    print(f"✗ ERROR")
    print(f"{'='*60}")
    print(f"{e}")
    print("\nPossible causes:")
    print("  • Tapir Add-On is not installed or not loaded")
    print("  • Tapir Add-On version is too old (needs v1.0.3+)")
    print("  • Archicad version incompatibility")
    print("\nTo install/update Tapir:")
    print("  https://github.com/ENZYME-APD/tapir-archicad-automation/releases")

print(f"\n{'='*60}")
print("COLOR REFERENCE")
print(f"{'='*60}")
print("Common colors (copy to HIGHLIGHT_COLOR above):")
print("  • Red:     [255, 0, 0, 255]")
print("  • Green:   [0, 255, 0, 255]")
print("  • Blue:    [0, 0, 255, 255]")
print("  • Yellow:  [255, 255, 0, 255]")
print("  • Orange:  [255, 165, 0, 255]")
print("  • Magenta: [255, 0, 255, 255]")
print("  • Cyan:    [0, 255, 255, 255]")
print("  • Purple:  [128, 0, 128, 255]")
print("  • White:   [255, 255, 255, 255]")
