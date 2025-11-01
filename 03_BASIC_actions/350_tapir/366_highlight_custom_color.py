"""
Script: Highlight Elements - Advanced (CORRECT FORMAT)
API: Tapir Add-On v1.0.3+
Description: Advanced highlighting with per-element colors and wireframe mode
Usage: 
    1. Modify configuration below
    2. Select elements
    3. Run script
Requirements:
    - archicad-api package
    - Tapir Add-On v1.0.3+ installed
"""

from archicad import ACConnection

# ============ CONFIGURATION ============

# MODE 1: Single color for all elements
USE_SINGLE_COLOR = False
SINGLE_COLOR = [255, 100, 100, 255]  # Red

# MODE 2: Different colors per element type (only used if USE_SINGLE_COLOR = False)
COLORS_BY_TYPE = {
    'Wall': [255, 0, 0, 255],      # Red
    'Slab': [0, 255, 0, 255],      # Green
    'Column': [0, 0, 255, 255],    # Blue
    'Beam': [255, 255, 0, 255],    # Yellow
    'Window': [0, 255, 255, 255],  # Cyan
    'Door': [255, 0, 255, 255],    # Magenta
    'Object': [255, 165, 0, 255],  # Orange
    'default': [128, 128, 128, 255]  # Gray for unknown types
}

# Wireframe mode: Make non-highlighted elements wireframe in 3D
WIREFRAME_3D = False

# Optional: Color for non-highlighted elements
USE_NON_HIGHLIGHTED_COLOR = True
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

print(f"\n{'='*70}")
print(f"ADVANCED ELEMENT HIGHLIGHTING")
print(f"{'='*70}")
print(f"Elements to highlight: {len(selected_elements)}")
print(f"Mode: {'Single Color' if USE_SINGLE_COLOR else 'Color by Type'}")
print(f"Wireframe 3D: {WIREFRAME_3D}")
if USE_NON_HIGHLIGHTED_COLOR:
    print(f"Non-highlighted color: {NON_HIGHLIGHTED_COLOR}")

# Get element types
element_types = acc.GetTypesOfElements(
    [elem.elementId for elem in selected_elements])

# Build elements array and colors array
elements_array = []
colors_array = []

print("\nElements and colors:")
print("-" * 70)

for i, (elem, elem_type) in enumerate(zip(selected_elements, element_types)):
    elem_type_name = elem_type.typeOfElement.elementType
    elem_guid = str(elem.elementId.guid)

    # Add element with correct structure
    elements_array.append({
        'elementId': {
            'guid': elem_guid
        }
    })

    # Determine color for this element
    if USE_SINGLE_COLOR:
        color = SINGLE_COLOR
    else:
        color = COLORS_BY_TYPE.get(elem_type_name, COLORS_BY_TYPE['default'])

    colors_array.append(color)

    # Display info
    color_name = f"RGBA{color}"
    print(
        f"  {i+1:3}. {elem_type_name:15} | {color_name:25} | {elem_guid[:18]}...")

try:
    # Build parameters with correct schema
    params = {
        'elements': elements_array,
        'highlightedColors': colors_array,
        'wireframe3D': WIREFRAME_3D
    }

    # Add optional non-highlighted color
    if USE_NON_HIGHLIGHTED_COLOR:
        params['nonHighlightedColor'] = NON_HIGHLIGHTED_COLOR

    # Execute the Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'HighlightElements'),
        params
    )

    print(f"\n{'='*70}")
    print(f"✓ SUCCESS")
    print(f"{'='*70}")
    print(f"Highlighted {len(selected_elements)} element(s)")

    if response.get('success'):
        print("\nElements are now highlighted in Archicad!")

    if WIREFRAME_3D:
        print("\n💡 Non-highlighted elements are shown as wireframe in 3D")

    print(f"\n💡 To clear highlights, run: 96_clear_highlights_FINAL.py")

except Exception as e:
    print(f"\n{'='*70}")
    print(f"✗ ERROR")
    print(f"{'='*70}")
    print(f"{e}")
    print("\nPossible causes:")
    print("  • Tapir Add-On is not installed or not loaded")
    print("  • Tapir Add-On version is too old (needs v1.0.3+)")
    print("  • Archicad version incompatibility")
    print("\nTo install/update Tapir:")
    print("  https://github.com/ENZYME-APD/tapir-archicad-automation/releases")

print(f"\n{'='*70}")
print("TIPS")
print(f"{'='*70}")
print("\n1. To use different colors per type:")
print("   Set USE_SINGLE_COLOR = False")
print("\n2. To show non-highlighted elements as wireframe:")
print("   Set WIREFRAME_3D = True")
print("\n3. To gray out non-highlighted elements:")
print("   Set USE_NON_HIGHLIGHTED_COLOR = True")
print("\n4. Add more element types to COLORS_BY_TYPE as needed")
