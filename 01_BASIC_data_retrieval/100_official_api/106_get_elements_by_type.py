"""
================================================================================
SCRIPT: Get Elements by Type
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves elements of a specific type (Wall, Slab, Column, etc.) from the 
current Archicad project. Also provides a summary of element counts for common
element types.

This script is useful for filtering elements by type before performing 
operations on specific categories of elements.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- ELEMENT_TYPE              : Type of element to retrieve (default: "Wall")
- MAX_ELEMENTS_TO_DISPLAY   : Maximum number of element GUIDs to display (default: 10)

Common element types: "Wall", "Slab", "Column", "Beam", "Window", "Door", 
                      "Object", "Roof", "Curtain Wall", "Zone", "Mesh"

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetElementsByType()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetElementsByType

[Data Types]
- ElementIdArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementIdArrayItem
  
- ElementId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project with elements
2. Modify ELEMENT_TYPE to target desired element type
3. Adjust MAX_ELEMENTS_TO_DISPLAY if needed
4. Run this script
5. View filtered elements in console output

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. FILTERED ELEMENTS:
   - Count of elements matching the specified type
   - Element GUIDs (limited by MAX_ELEMENTS_TO_DISPLAY)

2. ELEMENT COUNT BY TYPE:
   - Summary of common element types with counts
   - Only displays types that exist in the project

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

If you have multiple Archicad instances open:
- The script connects to the first one opened
- Make sure the correct project is open in that instance
- Close other instances if you need to target a specific project

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Element type names are case-sensitive
- If an invalid element type is specified, result will be empty
- Common types summary only checks predefined types
- Number of displayed GUIDs can be adjusted via MAX_ELEMENTS_TO_DISPLAY

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 104_get_all_elements.py
- 105_get_selected_elements.py
- 111_get_element_full_info.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# ============================================================================
# CONFIGURATION
# ============================================================================
# Change this to get different element types
# Common types: "Wall", "Slab", "Column", "Beam", "Window", "Door",
#               "Object", "Roof", "Curtain Wall", "Zone", "Mesh"
ELEMENT_TYPE = "Wall"
MAX_ELEMENTS_TO_DISPLAY = 10  # Number of element GUIDs to display

# ============================================================================
# GET ELEMENTS BY TYPE
# ============================================================================
elements = acc.GetElementsByType(ELEMENT_TYPE)

# Display results
print(f"\n=== {ELEMENT_TYPE.upper()} ELEMENTS ===")
print(f"Total {ELEMENT_TYPE}s found: {len(elements)}")

if len(elements) == 0:
    print(f"\n⚠️  No {ELEMENT_TYPE} elements found in the project")
else:
    print(
        f"\nFirst {min(MAX_ELEMENTS_TO_DISPLAY, len(elements))} {ELEMENT_TYPE}s:")
    for i, element in enumerate(elements[:MAX_ELEMENTS_TO_DISPLAY]):
        print(f"{i+1}. GUID: {element.elementId.guid}")

    if len(elements) > MAX_ELEMENTS_TO_DISPLAY:
        print(f"\n... and {len(elements) - MAX_ELEMENTS_TO_DISPLAY} more")

# ============================================================================
# SUMMARY: ELEMENT COUNT BY TYPE
# ============================================================================
print("\n=== ELEMENT COUNT BY TYPE ===")
common_types = ["Wall", "Slab", "Column", "Beam", "Window", "Door",
                "Object", "Roof", "Zone", "Curtain Wall", "Mesh"]

for elem_type in common_types:
    try:
        elements = acc.GetElementsByType(elem_type)
        if len(elements) > 0:
            print(f"{elem_type}: {len(elements)}")
    except:
        # Skip if element type doesn't exist or causes error
        pass

print("\n💡 TIP: Modify ELEMENT_TYPE variable to filter different element types")
