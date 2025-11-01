"""
================================================================================
SCRIPT: Get Element 3D Bounding Box
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Element Information - 3D Analysis

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Gets 3D bounding box dimensions for selected elements in Archicad. This script
provides:
- 3D bounding box coordinates (min/max X, Y, Z)
- Width, depth, and height dimensions
- Element positioning in 3D space
- Bounding volume calculations
- Element type identification

The 3D bounding box is the smallest rectangular box that completely contains
the element's 3D geometry, useful for collision detection, spatial analysis,
and element positioning.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Elements must be selected in Archicad

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements
  Returns currently selected elements

- Get3DBoundingBoxes(elementIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.Get3DBoundingBoxes
  Returns 3D bounding boxes for specified elements

- GetTypesOfElements(elementIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements
  Returns element type information

[Data Types]
- ElementId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementId

- BoundingBox3D
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.BoundingBox3D
  Contains 3D bounding box coordinates

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select one or more elements
3. Run this script
4. Review 3D bounding box information

No configuration needed - script analyzes selected elements.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows for each selected element:
- Element type
- Bounding box position (X, Y, Z ranges)
- Dimensions (width, depth, height)
- Bounding volume

Example output:
  ═══════════════════════════════════════════════════════
  3D BOUNDING BOXES
  ═══════════════════════════════════════════════════════
  Analyzing 2 element(s)
  
  --- Element 1: Wall ---
    Position:
      X: 0.00 to 5.00
      Y: 0.00 to 0.30
      Z: 0.00 to 3.00
    
    Dimensions:
      Width (X):  5.00 m
      Depth (Y):  0.30 m
      Height (Z): 3.00 m
    
    Bounding Volume: 4.500 m³
  
  --- Element 2: Column ---
    Position:
      X: 5.00 to 5.40
      Y: 0.00 to 0.40
      Z: 0.00 to 3.00
    
    Dimensions:
      Width (X):  0.40 m
      Depth (Y):  0.40 m
      Height (Z): 3.00 m
    
    Bounding Volume: 0.480 m³
  
  ══════════════════════════════════════════════════════
  Analysis complete

Example output (No selection):
  ⚠  No elements selected
  Please select elements in Archicad and run again

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
3D BOUNDING BOX:
- Axis-aligned rectangular box (not rotated)
- Contains entire 3D geometry of element
- Minimum (xMin, yMin, zMin) and maximum (xMax, yMax, zMax) corners
- Coordinates in project coordinate system

COORDINATES:
- X axis: Width direction
- Y axis: Depth direction
- Z axis: Height/elevation direction
- All values in meters
- Relative to project origin

DIMENSIONS:
- Width: xMax - xMin (X-axis extent)
- Depth: yMax - yMin (Y-axis extent)
- Height: zMax - zMin (Z-axis extent)

BOUNDING VOLUME:
- Calculated as: Width × Depth × Height
- Represents occupied space, not actual element volume
- Always ≥ actual element volume
- Useful for space planning approximations

USE CASES:
- Collision detection between elements
- Spatial analysis and clearance checking
- Element positioning validation
- Rough volume calculations
- BIM coordination checks
- Export bounding information

LIMITATIONS:
- Bounding box is axis-aligned (not rotated with element)
- Volume is approximate (larger than actual)
- Does not account for element voids or cutouts
- 2D elements may have minimal Z extent

ELEMENT TYPES SUPPORTED:
- Walls, columns, beams, slabs, roofs
- Doors, windows
- Objects, lamps
- Curtain walls, shells, morphs
- Stairs, railings
- Most 3D modeling elements

TROUBLESHOOTING:
- If no bounding box: Element may lack 3D representation
- If unexpected values: Check project coordinate system
- If all zeros: Element may be at project origin

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 101_get_all_elements.py (get all elements)
- 111_get_element_full_info.py (detailed element info)
- 159_get_element_subelements.py (hierarchical elements)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# Get selected elements
selected_elements = acc.GetSelectedElements()

if len(selected_elements) == 0:
    print("\n⚠ No elements selected")
    print("Please select elements in Archicad and run again")
    exit()

print(f"\n=== 3D BOUNDING BOXES ===")
print(f"Analyzing {len(selected_elements)} element(s)\n")

# Get 3D bounding boxes
element_ids = [elem.elementId for elem in selected_elements]
bounding_boxes = acc.Get3DBoundingBoxes(element_ids)

# Get element types for context
element_types = acc.GetTypesOfElements(element_ids)

# Display results
for i, (bbox, elem_type) in enumerate(zip(bounding_boxes, element_types)):
    print(f"--- Element {i+1}: {elem_type.typeOfElement.elementType} ---")
    
    if hasattr(bbox, 'boundingBox3D'):
        box = bbox.boundingBox3D
        
        # Calculate dimensions
        width = box.xMax - box.xMin
        depth = box.yMax - box.yMin
        height = box.zMax - box.zMin
        
        print(f"  Position:")
        print(f"    X: {box.xMin:.2f} to {box.xMax:.2f}")
        print(f"    Y: {box.yMin:.2f} to {box.yMax:.2f}")
        print(f"    Z: {box.zMin:.2f} to {box.zMax:.2f}")
        
        print(f"\n  Dimensions:")
        print(f"    Width (X):  {width:.2f} m")
        print(f"    Depth (Y):  {depth:.2f} m")
        print(f"    Height (Z): {height:.2f} m")
        
        # Calculate volume
        volume = width * depth * height
        print(f"\n  Bounding Volume: {volume:.3f} m³")
        
    else:
        print("  ⚠ No 3D bounding box available")
    
    print()  # Empty line for readability

print("="*50)
print("Analysis complete")