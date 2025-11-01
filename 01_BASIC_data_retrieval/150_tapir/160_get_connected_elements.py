"""
================================================================================
SCRIPT: Get Connected Elements (Owner-Connected Relationships)
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Element Information - Relationships

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves elements that are in an OWNER-CONNECTED relationship with specified
elements. This script provides:
- Windows and doors in walls
- Skylights in roofs and shells
- Labels on elements
- Host elements of openings

⚠️  IMPORTANT: This API finds PARENT-CHILD relationships, NOT physical
connections between elements (like walls touching walls). For example:
- Wall → Windows, Doors (✓ Supported)
- Roof → Skylights (✓ Supported)
- Wall → Wall (✗ NOT Supported)

Based on Tapir API documentation v1.1.4 and Archicad C++ API.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)
- Elements with owner-connected relationships must be selected

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetConnectedElements
  Returns elements in owner-connected relationships

[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements

- GetTypesOfElements(elementIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Data Types]
- AddOnCommandId
- ElementId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select parent elements (walls, roofs, shells)
3. Run this script
4. View windows, doors, skylights, or labels

No configuration needed - script automatically detects element types.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Summary of selected elements
- Each element analyzed individually with clear numbering [X/Total]
- Status indicator (✓ has connections / ⊗ no connections)
- All connected elements grouped by parent
- Overall summary with statistics

Example output (Walls with windows/doors/labels):
  ════════════════════════════════════════════════════════════
  CONNECTED ELEMENTS ANALYSIS
  ════════════════════════════════════════════════════════════
  
  SELECTED ELEMENTS:
    • 4 Wall(s)
    • 1 Slab(s)
  
  ════════════════════════════════════════════════════════════
  WALLS ANALYSIS (4 total)
  ════════════════════════════════════════════════════════════
  
  [1/4] WALL
  GUID: 7aa9f0bb-eeda-c841-9482-845aae59edf5
  Status: ✓ HAS CONNECTED ELEMENTS (1 Door(s))
  
    Doors: 1
      1. FC839630-ADFB-CC4B-AB86-F2B03C98360A
  
  ────────────────────────────────────────────────────────────
  
  [2/4] WALL
  GUID: 4915244c-beea-4f4a-8e05-575dd6bb29c2
  Status: ✓ HAS CONNECTED ELEMENTS (2 Window(s))
  
    Windows: 2
      1. FC7CF43D-2A17-5448-AB7F-BFBC07047A5F
      2. 6165624D-8155-524E-B742-4FD109A2B3E4
  
  ────────────────────────────────────────────────────────────
  
  [3/4] WALL
  GUID: 824ae2d4-666b-0140-bc6c-10ff8cc53180
  Status: ✓ HAS CONNECTED ELEMENTS (1 Window(s), 1 Label(s))
  
    Windows: 1
      1. C9711F1E-F6AA-C541-B1F6-FB46278DC1E3
    
    Labels: 1
      1. 94AA867B-48CC-9E4E-A868-B17CFFA8B202
  
  ────────────────────────────────────────────────────────────
  
  [4/4] WALL
  GUID: fee8fe69-39f8-ed41-bd8e-0b364c0c4f7f
  Status: ⊗ No connected elements
  
  ────────────────────────────────────────────────────────────
  
  ════════════════════════════════════════════════════════════
  SLABS ANALYSIS (1 total)
  ════════════════════════════════════════════════════════════
  
  [1/1] SLAB
  GUID: 7dd78a15-33ae-e64e-b61f-a7ddf7e14c9c
  Status: ✓ HAS CONNECTED ELEMENTS (1 Label(s))
  
    Labels: 1
      1. 94AA867B-48CC-9E4E-A868-B17CFFA8B202
  
  ────────────────────────────────────────────────────────────
  
  ════════════════════════════════════════════════════════════
  SUMMARY
  ════════════════════════════════════════════════════════════
  
  Connections found:
    • Wall → Door: 1
    • Wall → Window: 3
    • Wall → Label: 1
    • Slab → Label: 1

Example output (No connections):
  ════════════════════════════════════════════════════════════
  SUMMARY
  ════════════════════════════════════════════════════════════
  
  ✗ No connected elements found.
  
  Note: GetConnectedElements finds OWNER-CONNECTED relationships:
    • Walls → Windows, Doors, Labels
    • Roofs/Shells → Skylights, Labels
    • All elements → Labels

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
SUPPORTED RELATIONSHIPS (Owner → Connected):
┌─────────────────────┬──────────────────────────────────┐
│ Parent Element      │ Connected Elements               │
├─────────────────────┼──────────────────────────────────┤
│ Wall                │ Window, Door, Label              │
│ Roof, Shell         │ Skylight, Label                  │
│ Opening             │ Wall, Slab, Mesh, Beam (hosts)  │
│ All labelable elem. │ Label (always checked)           │
└─────────────────────┴──────────────────────────────────┘

Note: Labels are checked for ALL element types, as any labelable element
can have attached labels.

WHAT THIS API DOES:
✓ Find windows and doors in a wall
✓ Find skylights in a roof or shell
✓ Find labels attached to elements
✓ Find host elements of an opening

WHAT THIS API DOES NOT DO:
✗ Find walls touching other walls (corner connections)
✗ Find beams connected to columns
✗ Find slabs connected to walls
✗ General spatial/physical connections

For physical connections (walls touching walls, etc.), you would need a
different approach (spatial analysis, bounding box overlap, etc.).

OWNER-CONNECTED RELATIONSHIP:
- Parent "owns" the child element
- Child element placed within/on parent
- Examples: window in wall, door in wall, skylight in roof
- Hierarchical, not spatial relationship

API PARAMETERS:
- elements: Array of parent element IDs
- connectedElementType: Type to search for ("Window", "Door", "Skylight", etc.)
- Returns: For each parent, array of connected elements of that type

ELEMENT TYPES (connectedElementType):
Must be one of: "Wall", "Column", "Beam", "Window", "Door", "Object", "Lamp",
"Slab", "Roof", "Mesh", "Skylight", "Shell", "Label", "Opening", etc.
(See Tapir documentation for complete list)

USE CASES:
- Count windows/doors per wall
- List skylights in roof
- Export window schedule from walls
- Validate element placement
- BIM quality control
- Quantity takeoff
- Generate reports

WINDOW/DOOR QUERIES:
- Select walls
- Query for "Window" or "Door"
- Get all windows/doors in those walls
- Useful for schedules and counts

SKYLIGHT QUERIES:
- Select roofs or shells
- Query for "Skylight"
- Get all skylights in those elements
- Useful for daylighting analysis

LABEL QUERIES:
- Select any labelable element
- Query for "Label"
- Get labels attached to elements
- Useful for annotation management

OPENING QUERIES:
- Select openings
- Query for host type ("Wall", "Slab", etc.)
- Get elements hosting the openings
- Reverse relationship

SCRIPT LOGIC:
1. Get selected elements
2. Group by element type
3. For each type, determine valid connected types
4. Query for each valid connection
5. Display results grouped by parent element

PERFORMANCE:
- Efficient for multiple elements
- Single API call per element type combination
- Results grouped for easy processing

TROUBLESHOOTING:
- If no results: Check parent-child relationship is supported
- If wrong type: Verify element type names (case-sensitive)
- If command fails: Ensure Tapir Add-On is installed
- If unexpected results: Review supported relationships table

LIMITATIONS:
- Requires Tapir Add-On
- Only finds owner-connected relationships
- Not for spatial/physical connections
- Type names are case-sensitive

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 159_get_element_subelements.py (hierarchical element components)
- 158_get_element_3D_bounding_box.py (spatial analysis)
- 111_get_element_full_info.py (detailed element info)
- 151_check_tapir_version.py (verify Tapir installation)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


# Supported owner-connected relationships
# Note: Labels are checked separately for ALL elements as any labelable element can have labels
SUPPORTED_CONNECTIONS = {
    'Wall': ['Window', 'Door'],
    'Roof': ['Skylight'],
    'Shell': ['Skylight'],
    'Opening': ['Wall', 'Slab', 'Mesh', 'Beam'],
}


def get_connected_elements(element_ids, connected_type):
    """
    Get connected elements of a specific type.
    
    Args:
        element_ids: List of ElementId objects
        connected_type: Type of connected elements ('Window', 'Door', 'Skylight', etc.)
        
    Returns:
        List of connected element arrays (one per input element)
    """
    if not isinstance(element_ids, list):
        element_ids = [element_ids]
    
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'GetConnectedElements')
        
        elements_param = []
        for elem_id in element_ids:
            elements_param.append({
                'elementId': {
                    'guid': str(elem_id.guid)
                }
            })
        
        parameters = {
            'elements': elements_param,
            'connectedElementType': connected_type
        }
        
        response = acc.ExecuteAddOnCommand(command_id, parameters)
        
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get('message', 'Unknown error')
            print(f"✗ Command failed: {error_msg}")
            return None
        
        if isinstance(response, dict) and 'connectedElements' in response:
            return parse_connected_elements(response['connectedElements'])
        
        return None
            
    except Exception as e:
        print(f"✗ Error: {e}")
        return None


def parse_connected_elements(connected_data):
    """Parse connected elements from response."""
    if not connected_data:
        return []
    
    results = []
    for item in connected_data:
        if isinstance(item, dict) and 'elements' in item:
            elements = item['elements']
            element_guids = []
            
            for elem in elements:
                if isinstance(elem, dict):
                    guid = elem.get('elementId', {}).get('guid')
                    if guid:
                        element_guids.append(guid)
            
            results.append(element_guids)
        else:
            results.append([])
    
    return results


def analyze_selected_elements():
    """
    Analyze selected elements for connected relationships.
    """
    selected = acc.GetSelectedElements()
    
    if not selected or len(selected) == 0:
        print("✗ No elements selected.")
        print("\nPlease select parent elements:")
        print("  • Walls (to find windows/doors)")
        print("  • Roofs or Shells (to find skylights)")
        print("  • Any element (to find labels)")
        return
    
    # Group elements by type
    element_ids = [elem.elementId for elem in selected]
    element_types_response = acc.GetTypesOfElements(element_ids)
    
    elements_by_type = {}
    for elem_id, type_response in zip(element_ids, element_types_response):
        elem_type = type_response.typeOfElement.elementType
        if elem_type not in elements_by_type:
            elements_by_type[elem_type] = []
        elements_by_type[elem_type].append(elem_id)
    
    print("\n" + "="*80)
    print("CONNECTED ELEMENTS ANALYSIS")
    print("="*80 + "\n")
    
    # Summary of selected elements
    print("SELECTED ELEMENTS:")
    for elem_type, elem_ids in elements_by_type.items():
        print(f"  • {len(elem_ids)} {elem_type}(s)")
    print()
    
    found_any = False
    overall_stats = {}
    
    # For each element type, check supported connections
    for elem_type, elem_ids in elements_by_type.items():
        print("="*80)
        print(f"{elem_type.upper()}S ANALYSIS ({len(elem_ids)} total)")
        print("="*80 + "\n")
        
        # Get supported connected types for this element type
        connected_types = []
        
        if elem_type in SUPPORTED_CONNECTIONS:
            # Add specific connections (Window, Door, Skylight, etc.)
            connected_types.extend(SUPPORTED_CONNECTIONS[elem_type])
        
        # Always try Label for all elements (all labelable elements can have labels)
        connected_types.append('Label')
        
        # Store results for all elements first
        all_results = {}
        for connected_type in connected_types:
            results = get_connected_elements(elem_ids, connected_type)
            if results:
                all_results[connected_type] = results
        
        # Now display element by element
        for elem_num, elem_id in enumerate(elem_ids, 1):
            has_connections = False
            connections_summary = []
            
            # Check what this element has
            for connected_type, results in all_results.items():
                if elem_num - 1 < len(results):
                    connected_guids = results[elem_num - 1]
                    if connected_guids and len(connected_guids) > 0:
                        has_connections = True
                        connections_summary.append(f"{len(connected_guids)} {connected_type}(s)")
            
            # Display element header
            print(f"[{elem_num}/{len(elem_ids)}] {elem_type.upper()}")
            print(f"GUID: {elem_id.guid}")
            
            if has_connections:
                found_any = True
                print(f"Status: ✓ HAS CONNECTED ELEMENTS ({', '.join(connections_summary)})")
                print()
                
                # Display detailed connections
                for connected_type, results in all_results.items():
                    if elem_num - 1 < len(results):
                        connected_guids = results[elem_num - 1]
                        if connected_guids and len(connected_guids) > 0:
                            print(f"  {connected_type}s: {len(connected_guids)}")
                            
                            # Track stats
                            stat_key = f"{elem_type} → {connected_type}"
                            overall_stats[stat_key] = overall_stats.get(stat_key, 0) + len(connected_guids)
                            
                            for j, guid in enumerate(connected_guids, 1):
                                print(f"    {j}. {guid}")
                            print()
            else:
                print("Status: ⊗ No connected elements")
                print()
            
            print("-"*80 + "\n")
    
    # Overall summary
    print("="*80)
    print("SUMMARY")
    print("="*80)
    
    if found_any:
        print("\nConnections found:")
        for stat_key, count in sorted(overall_stats.items()):
            print(f"  • {stat_key}: {count}")
    else:
        print("\n✗ No connected elements found.")
        print("\nNote: GetConnectedElements finds OWNER-CONNECTED relationships:")
        print("  • Walls → Windows, Doors, Labels")
        print("  • Roofs/Shells → Skylights, Labels")
        print("  • All elements → Labels")
        print("\nMake sure you selected parent elements (walls, roofs, shells).")
    
    print("\n" + "="*80)


def get_windows_and_doors_from_walls():
    """
    Example: Get all windows and doors from selected walls.
    """
    selected = acc.GetSelectedElements()
    
    if not selected:
        print("No elements selected.")
        return
    
    # Filter for walls only
    element_ids = [elem.elementId for elem in selected]
    element_types = acc.GetTypesOfElements(element_ids)
    
    wall_ids = []
    for elem_id, type_response in zip(element_ids, element_types):
        if type_response.typeOfElement.elementType == 'Wall':
            wall_ids.append(elem_id)
    
    if not wall_ids:
        print("No walls selected.")
        return
    
    print("\n" + "="*80)
    print(f"WINDOWS AND DOORS IN {len(wall_ids)} WALL(S)")
    print("="*80 + "\n")
    
    # Get windows
    windows_results = get_connected_elements(wall_ids, 'Window')
    
    # Get doors
    doors_results = get_connected_elements(wall_ids, 'Door')
    
    total_windows = 0
    total_doors = 0
    
    for i, wall_id in enumerate(wall_ids):
        windows = windows_results[i] if windows_results and i < len(windows_results) else []
        doors = doors_results[i] if doors_results and i < len(doors_results) else []
        
        if windows or doors:
            print(f"WALL {i+1}")
            print(f"GUID: {wall_id.guid}")
            
            if windows:
                print(f"\n  Windows: {len(windows)}")
                total_windows += len(windows)
                for j, guid in enumerate(windows, 1):
                    print(f"    {j}. {guid}")
            
            if doors:
                print(f"\n  Doors: {len(doors)}")
                total_doors += len(doors)
                for j, guid in enumerate(doors, 1):
                    print(f"    {j}. {guid}")
            
            print("\n" + "-"*80 + "\n")
    
    print("="*80)
    print("SUMMARY")
    print("="*80)
    print(f"Total walls analyzed: {len(wall_ids)}")
    print(f"Total windows found: {total_windows}")
    print(f"Total doors found: {total_doors}")
    print("="*80)


def main():
    """Main function."""
    
    # Option 1: Automatic analysis of selected elements
    analyze_selected_elements()
    
    # Option 2: Specific example for walls
    # Uncomment to use:
    # print("\n" + "="*80)
    # print("EXAMPLE: Windows and Doors in Walls")
    # print("="*80)
    # get_windows_and_doors_from_walls()


if __name__ == "__main__":
    main()