"""
================================================================================
SCRIPT: Get Selected Elements
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves currently selected elements in Archicad and displays their GUIDs,
types, and statistics. Useful for identifying elements before performing
operations on the selection.

This script requires elements to be selected in Archicad before execution.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)
- At least one element selected in Archicad

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements
  
- GetTypesOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements

[Data Types]
- ElementIdArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementIdArrayItem
  
- ElementId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementId
  
- TypeOfElementWrapper
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac26.html#archicad.releases.ac26.b3005types.TypeOfElementWrapper

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select one or more elements
3. Run this script
4. View element information in console output

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. SELECTED ELEMENTS COUNT:
   - Total number of selected elements

2. ELEMENT GUIDS:
   - Complete list of all selected element GUIDs

3. ELEMENT TYPES:
   - Type of each selected element
   - Numbered list corresponding to GUIDs

4. COUNT BY TYPE (if multiple elements):
   - Summary of element types with counts

5. SUMMARY:
   - Total selected elements
   - Number of different types (if applicable)
   - Suggestions for next steps

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
- If no elements are selected, script exits with a warning
- Element types are accessed via TypeOfElementWrapper.typeOfElement.elementType
- For single element selection, type count summary is not displayed

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 104_get_all_elements.py
- 106_get_elements_by_type.py
- 111_get_element_full_info.py
- 201_set_element_classification.py
- 303_list_element_property_values.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# ============================================================================
# GET SELECTED ELEMENTS
# ============================================================================
# GetSelectedElements returns a list of currently selected elements
selected_elements = acc.GetSelectedElements()

print(f"\n=== SELECTED ELEMENTS ===")
print(f"Total selected: {len(selected_elements)}")

if len(selected_elements) == 0:
    print("\n⚠️  No elements are currently selected")
    print("Please select some elements in Archicad and run again")
    exit()

# ============================================================================
# DISPLAY GUIDS
# ============================================================================
print("\nSelected element GUIDs:")
for i, element in enumerate(selected_elements):
    print(f"{i+1}. {element.elementId.guid}")

# ============================================================================
# GET AND DISPLAY ELEMENT TYPES
# ============================================================================
print("\n=== ELEMENT TYPES ===")

# Get element IDs
element_ids = [element.elementId for element in selected_elements]

# Get types for all selected elements
element_types = acc.GetTypesOfElements(element_ids)

# Display types
for i, elem_type_wrapper in enumerate(element_types):
    type_name = elem_type_wrapper.typeOfElement.elementType
    print(f"{i+1}. {type_name}")

# ============================================================================
# COUNT BY TYPE
# ============================================================================
if len(selected_elements) > 1:
    print("\n=== COUNT BY TYPE ===")

    type_count = {}
    for elem_type_wrapper in element_types:
        type_name = elem_type_wrapper.typeOfElement.elementType
        type_count[type_name] = type_count.get(type_name, 0) + 1

    for elem_type, count in sorted(type_count.items()):
        print(f"  {elem_type}: {count}")

# ============================================================================
# SUMMARY
# ============================================================================
print("\n" + "="*70)
print("SUMMARY")
print("="*70)
print(f"Total selected elements: {len(selected_elements)}")

if len(selected_elements) > 1:
    print(f"Different types: {len(type_count)}")

print("\n💡 NEXT STEPS:")
print("   - Use 111_get_element_full_info.py for complete element details")
print("   - Use 303_list_element_property_values.py for property values")
print("   - Use 201_set_element_classification.py to classify these elements")
