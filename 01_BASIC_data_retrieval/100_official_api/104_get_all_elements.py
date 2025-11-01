"""
================================================================================
SCRIPT: Get All Elements
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves all elements from the current Archicad project.  Displays element GUIDs and counts elements by 
type.

This script demonstrates:
- Getting all elements in the project
- Getting element types for multiple elements
- Counting elements by type

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- MAX_ELEMENTS_TO_DISPLAY : Maximum number of element GUIDs to display (default: 10)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAllElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllElements
  
- GetTypesOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements

[Data Types]
- ElementIdArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementIdArrayItem
  
- ElementId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementId
  
- TypeOfElementWrapper
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac26.html#archicad.releases.ac26.b3005types.TypeOfElementWrapper
  
- TypeOfElement
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac26.html#archicad.releases.ac26.b3005types.TypeOfElement 

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project with elements
2. Adjust MAX_ELEMENTS_TO_DISPLAY if needed
3. Run this script
4. View element statistics in console output

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. TOTAL ELEMENT COUNT:
   - Number of elements found in the project

2. ELEMENT GUID SAMPLE:
   - Element GUIDs (limited by MAX_ELEMENTS_TO_DISPLAY)

3. ELEMENT COUNT BY TYPE:
   - Complete list of element types with counts
   - Alphabetically sorted

4. SUMMARY:
   - Total elements
   - Number of different types

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
- Element types are accessed via TypeOfElementWrapper.typeOfElement.elementType
- Number of displayed GUIDs can be adjusted via MAX_ELEMENTS_TO_DISPLAY

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 105_get_selected_elements.py
- 106_get_elements_by_type.py
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
MAX_ELEMENTS_TO_DISPLAY = 10  # Number of element GUIDs to display

# ============================================================================
# GET ALL ELEMENTS
# ============================================================================
# GetAllElements returns a list of all elements in the active project
elements = acc.GetAllElements()

print(f"\n=== ALL ELEMENTS ===")
print(f"Total elements found: {len(elements)}")

if len(elements) == 0:
    print("\n⚠️  No elements found in the project")
    exit()

# ============================================================================
# DISPLAY ELEMENT GUIDS
# ============================================================================
print(f"\nFirst {min(MAX_ELEMENTS_TO_DISPLAY, len(elements))} elements:")

for i, element in enumerate(elements[:MAX_ELEMENTS_TO_DISPLAY]):
    print(f"{i+1}. GUID: {element.elementId.guid}")

if len(elements) > MAX_ELEMENTS_TO_DISPLAY:
    print(f"\n... and {len(elements) - MAX_ELEMENTS_TO_DISPLAY} more elements")

# ============================================================================
# GET ELEMENT TYPES
# ============================================================================
print("\n=== CHECKING ELEMENT TYPES ===")

# Get element IDs for type checking
element_ids = [element.elementId for element in elements]

# Get types for all elements
element_types = acc.GetTypesOfElements(element_ids)

# ============================================================================
# COUNT ELEMENTS BY TYPE
# ============================================================================
# The TypeOfElementWrapper contains a 'typeOfElement' attribute
type_count = {}
for elem_type_wrapper in element_types:
    # Access the type from the wrapper object
    type_name = elem_type_wrapper.typeOfElement.elementType
    type_count[type_name] = type_count.get(type_name, 0) + 1

print("\nElement count by type:")
for elem_type, count in sorted(type_count.items()):
    print(f"  {elem_type}: {count}")

# ============================================================================
# SUMMARY
# ============================================================================
print("\n" + "="*70)
print("SUMMARY")
print("="*70)
print(f"Total elements: {len(elements)}")
print(f"Different types: {len(type_count)}")
