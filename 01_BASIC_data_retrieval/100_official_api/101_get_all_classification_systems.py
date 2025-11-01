"""
================================================================================
SCRIPT: Get All Classification Systems
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all classification systems available in the current 
Archicad project. Classification systems include standards like Uniclass, 
Omniclass, and custom systems.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAllClassificationSystems()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationSystems

[Data Types]
- ClassificationSystem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationSystem 
  
- ClassificationSystemId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationId.classificationSystemId 

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View classification systems in console output

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of classification systems
- For each system: Name, GUID, Version, Edition Date, Source, Description
- List of system GUIDs for use in other scripts

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
- If no classification systems exist, a warning message is displayed
- System GUIDs are essential for classification-related operations
- Not all systems have version, date, or description fields

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 102_get_classification_system_with_hierarchy.py
- 107_get_elements_by_classification.py
- 201_set_element_classification.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# Get all classification systems
classification_systems = acc.GetAllClassificationSystems()

# Display results
print(f"\n=== CLASSIFICATION SYSTEMS ===")
print(f"Total classification systems: {len(classification_systems)}")

if len(classification_systems) == 0:
    print("\n⚠️  No classification systems found in the project")
else:
    print("\nClassification Systems:\n")

    for i, system in enumerate(classification_systems):
        # Display system name
        print(f"{i+1}. Name: {system.name}")
        print(f"   GUID: {system.classificationSystemId.guid}")

        # Display version if available
        if hasattr(system, 'version') and system.version:
            print(f"   Version: {system.version}")

        # Display edition date if available
        if hasattr(system, 'editionDate') and system.editionDate:
            print(f"   Edition Date: {system.editionDate}")

        # Display source if available
        if hasattr(system, 'source') and system.source:
            print(f"   Source: {system.source}")

        # Display description if available
        if hasattr(system, 'description') and system.description:
            print(f"   Description: {system.description}")

        print()  # Empty line for readability

# Example: Get system IDs for later use
print("\n=== SYSTEM IDs (for use in other scripts) ===")
for system in classification_systems:
    system_info = f"{system.name}"
    if hasattr(system, 'version') and system.version:
        system_info += f" (v{system.version})"
    if hasattr(system, 'editionDate') and system.editionDate:
        system_info += f" - {system.editionDate}"
    print(f"{system_info}")
    print(f"  GUID: {system.classificationSystemId.guid}")
