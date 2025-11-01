"""
================================================================================
SCRIPT: Get Element Subelements
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Element Information - Hierarchical Elements

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves all subelements (components) of hierarchical elements like curtain
walls, stairs, railings, beams, and columns. This script provides:
- Complete list of subelements grouped by type
- Subelement counts and breakdowns
- Parent-child relationships
- Element GUIDs for further processing

Hierarchical elements in Archicad contain multiple subelements (e.g., curtain
wall has segments, frames, panels, etc.) that can be accessed individually.

Based on TAPIR API documentation v1.0.6

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)
- Hierarchical elements must be selected

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetSubelementsOfHierarchicalElements
  Returns all subelements of hierarchical elements with type grouping

[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Data Types]
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AddOnCommandId

- ElementId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select hierarchical elements (curtain walls, stairs, railings, beams, columns)
3. Run this script
4. Review subelement breakdown

No configuration needed - script analyzes selected elements.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Total subelement count per element
- Breakdown by subelement type
- Element GUIDs
- Export capability

Example output (Curtain Wall):
  ════════════════════════════════════════════════════════════
  HIERARCHICAL ELEMENT SUBELEMENTS
  ════════════════════════════════════════════════════════════
  
  Found 1 element(s) with subelements:
  
  ────────────────────────────────────────────────────────────
  ELEMENT 1
  ────────────────────────────────────────────────────────────
  GUID: 12345678-1234-1234-1234-123456789ABC
  Total Subelements: 45
  
  Subelement Types:
    • Curtain Wall Segments: 10
    • Curtain Wall Frames: 15
    • Curtain Wall Panels: 20

Example output (No hierarchical elements):
  No subelements found in selected elements.
  
  Note: Only hierarchical elements have subelements:
    • Curtain Walls (segments, frames, panels, junctions, accessories)
    • Stairs (risers, treads, structures)
    • Railings (posts, rails, handrails, balusters, etc.)
    • Beams (segments)
    • Columns (segments)

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
HIERARCHICAL ELEMENTS:
Elements that contain multiple component parts:
- Curtain Walls: segments, frames, panels, junctions, accessories
- Stairs: risers, treads, structures
- Railings: posts, rails, handrails, balusters, panels, etc.
- Beams: segments
- Columns: segments

SUBELEMENT TYPES:
Curtain Walls:
  - cWallSegments: Vertical/horizontal divisions
  - cWallFrames: Frame profiles
  - cWallPanels: Infill panels
  - cWallJunctions: Connection points
  - cWallAccessories: Additional components

Stairs:
  - stairRisers: Vertical faces between treads
  - stairTreads: Horizontal walking surfaces
  - stairStructures: Support structures

Railings:
  - railingNodes, railingSegments
  - railingPosts, railingInnerPosts
  - railingRails, railingToprails, railingHandrails
  - railingRailEnds, railingRailConnections
  - railingHandrailEnds, railingHandrailConnections
  - railingToprailEnds, railingToprailConnections
  - railingPatterns, railingPanels
  - railingBalusterSets, railingBalusters

Beams/Columns:
  - beamSegments: Beam component segments
  - columnSegments: Column component segments

USE CASES:
- Detailed curtain wall analysis
- Stair component counting
- Railing complexity assessment
- BIM model validation
- Quantity takeoff for hierarchical elements
- Access individual components for modification

SUBELEMENT GUIDS:
- Each subelement has its own GUID
- Can be used with other API commands
- Exported for further processing
- Useful for element selection/modification

LIMITATIONS:
- Requires Tapir Add-On
- Only works with hierarchical element types
- Non-hierarchical elements return no subelements
- Subelement properties limited to structure

TROUBLESHOOTING:
- If no subelements: Select hierarchical elements
- If command fails: Ensure Tapir is installed
- If unexpected types: Check element is hierarchical
- For empty results: Element may have no components

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 160_get_connected_elements.py (element connections)
- 158_get_element_3D_bounding_box.py (3D analysis)
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


def get_element_subelements(element_ids):
    """
    Get all subelements of given hierarchical elements.

    Args:
        element_ids: List of ElementId objects or a single ElementId

    Returns:
        Dictionary with subelements grouped by type
    """
    # Ensure element_ids is a list
    if not isinstance(element_ids, list):
        element_ids = [element_ids]

    try:
        # Create proper AddOnCommandId
        command_id = act.AddOnCommandId(
            'TapirCommand', 'GetSubelementsOfHierarchicalElements')

        # Prepare input parameters according to API docs
        # Input: { elements: [{ elementId: { guid: "..." } }] }
        # Note: Must convert UUID to string for JSON serialization
        elements_param = []
        for elem_id in element_ids:
            elements_param.append({
                'elementId': {
                    'guid': str(elem_id.guid)  # Convert UUID to string
                }
            })

        parameters = {
            'elements': elements_param
        }

        print(f"Retrieving subelements for {len(element_ids)} element(s)...")
        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error response
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ Command failed: {error_msg}")
            return None

        print("✓ Successfully retrieved subelements data\n")

        # Parse the response
        if isinstance(response, dict) and 'subelements' in response:
            return parse_subelements_dict(response['subelements'])
        elif hasattr(response, 'subelements'):
            return parse_subelements_object(response.subelements)
        else:
            print("⚠ Unexpected response format")
            return None

    except Exception as e:
        print(f"✗ Error retrieving subelements: {e}")
        import traceback
        traceback.print_exc()
        return None


def parse_subelements_dict(subelements_data):
    """Parse subelements from dictionary response."""
    if not subelements_data or len(subelements_data) == 0:
        return None

    # subelements is a list where each item contains subelements grouped by type
    # We'll take the first item (for the first parent element)
    if isinstance(subelements_data, list) and len(subelements_data) > 0:
        return subelements_data[0]

    return subelements_data


def parse_subelements_object(subelements_data):
    """Parse subelements from object response."""
    # Similar logic for object-based response
    if hasattr(subelements_data, '__iter__'):
        items = list(subelements_data)
        if len(items) > 0:
            return items[0]

    return subelements_data


def count_subelements(subelements):
    """
    Count total number of subelements across all types.

    Args:
        subelements: Dictionary with subelement types

    Returns:
        Total count and breakdown by type
    """
    if not subelements:
        return 0, {}

    counts = {}
    total = 0

    # Define all possible subelement types from API docs
    subelement_types = [
        'cWallSegments', 'cWallFrames', 'cWallPanels', 'cWallJunctions', 'cWallAccessories',
        'stairRisers', 'stairTreads', 'stairStructures',
        'railingNodes', 'railingSegments', 'railingPosts', 'railingRailEnds',
        'railingRailConnections', 'railingHandrailEnds', 'railingHandrailConnections',
        'railingToprailEnds', 'railingToprailConnections', 'railingRails',
        'railingToprails', 'railingHandrails', 'railingPatterns', 'railingInnerPosts',
        'railingPanels', 'railingBalusterSets', 'railingBalusters',
        'beamSegments', 'columnSegments'
    ]

    for sub_type in subelement_types:
        if isinstance(subelements, dict):
            elements = subelements.get(sub_type, [])
        else:
            elements = getattr(subelements, sub_type, []) if hasattr(
                subelements, sub_type) else []

        if elements:
            count = len(elements) if isinstance(elements, list) else 0
            if count > 0:
                counts[sub_type] = count
                total += count

    return total, counts


def get_subelements_for_selected():
    """
    Get subelements for all currently selected elements.

    Returns:
        Dictionary with results for each selected element
    """
    # Get selected elements
    selected = acc.GetSelectedElements()

    if not selected or len(selected) == 0:
        print("No elements selected.")
        return None

    print(f"Analyzing {len(selected)} selected element(s)...\n")

    results = []

    for element in selected:
        element_id = element.elementId

        # Get element type name for display
        try:
            # Convert UUID to string for display
            elem_guid = str(element_id.guid)
        except:
            elem_guid = "Unknown"

        # Get subelements for this element
        subelements = get_element_subelements([element_id])

        if subelements:
            total_count, type_counts = count_subelements(subelements)

            if total_count > 0:
                results.append({
                    'element_id': element_id,
                    'guid': str(element_id.guid),  # Convert UUID to string
                    'subelements': subelements,
                    'total_count': total_count,
                    'type_counts': type_counts
                })

    return results if results else None


def display_subelements_info(results):
    """
    Display subelements information in a formatted way.

    Args:
        results: List of results for each element
    """
    print("\n" + "="*80)
    print("HIERARCHICAL ELEMENT SUBELEMENTS")
    print("="*80 + "\n")

    if not results:
        print("No subelements found in selected elements.")
        print("\nNote: Only hierarchical elements have subelements:")
        print("  • Curtain Walls (segments, frames, panels, junctions, accessories)")
        print("  • Stairs (risers, treads, structures)")
        print("  • Railings (posts, rails, handrails, balusters, etc.)")
        print("  • Beams (segments)")
        print("  • Columns (segments)")
        print("\nMake sure you have selected one of these element types.")
        return

    print(f"Found {len(results)} element(s) with subelements:\n")

    for idx, result in enumerate(results, 1):
        print("─"*80)
        print(f"ELEMENT {idx}")
        print("─"*80)
        print(f"GUID: {result['guid']}")
        print(f"Total Subelements: {result['total_count']}\n")

        # Display breakdown by type
        print("Subelement Types:")
        for sub_type, count in sorted(result['type_counts'].items()):
            # Make the type name more readable
            display_name = sub_type
            if sub_type.startswith('cWall'):
                display_name = f"Curtain Wall {sub_type[5:]}"
            elif sub_type.startswith('stair'):
                display_name = f"Stair {sub_type[5:]}"
            elif sub_type.startswith('railing'):
                display_name = f"Railing {sub_type[7:]}"
            elif sub_type.startswith('beam'):
                display_name = f"Beam {sub_type[4:]}"
            elif sub_type.startswith('column'):
                display_name = f"Column {sub_type[6:]}"

            print(f"  • {display_name}: {count}")

        print()

    print("="*80)


def export_subelement_guids(results):
    """
    Export all subelement GUIDs for further processing.

    Args:
        results: List of results for each element

    Returns:
        List of all subelement GUIDs
    """
    all_guids = []

    if not results:
        return all_guids

    subelement_types = [
        'cWallSegments', 'cWallFrames', 'cWallPanels', 'cWallJunctions', 'cWallAccessories',
        'stairRisers', 'stairTreads', 'stairStructures',
        'railingNodes', 'railingSegments', 'railingPosts', 'railingRailEnds',
        'railingRailConnections', 'railingHandrailEnds', 'railingHandrailConnections',
        'railingToprailEnds', 'railingToprailConnections', 'railingRails',
        'railingToprails', 'railingHandrails', 'railingPatterns', 'railingInnerPosts',
        'railingPanels', 'railingBalusterSets', 'railingBalusters',
        'beamSegments', 'columnSegments'
    ]

    for result in results:
        subelements = result['subelements']

        for sub_type in subelement_types:
            if isinstance(subelements, dict):
                elements = subelements.get(sub_type, [])
            else:
                elements = getattr(subelements, sub_type, []) if hasattr(
                    subelements, sub_type) else []

            if elements:
                for elem in elements:
                    if isinstance(elem, dict):
                        guid = elem.get('elementId', {}).get('guid')
                    else:
                        elem_id = getattr(elem, 'elementId', None)
                        guid = getattr(elem_id, 'guid',
                                       None) if elem_id else None

                    if guid:
                        all_guids.append({
                            'guid': guid,
                            'type': sub_type,
                            'parent_guid': result['guid']
                        })

    return all_guids


def main():
    """Main function to demonstrate the script."""

    # Get subelements for selected elements
    results = get_subelements_for_selected()

    # Display the information
    display_subelements_info(results)

    # Example: Export subelement GUIDs
    if results:
        print("\n" + "─"*80)
        print("EXPORTED SUBELEMENT GUIDs")
        print("─"*80)

        guids = export_subelement_guids(results)
        print(f"\nTotal subelements exported: {len(guids)}")

        if len(guids) > 0:
            print("\nFirst 5 subelements:")
            for guid_info in guids[:5]:
                print(f"  • Type: {guid_info['type']}")
                print(f"    GUID: {guid_info['guid']}")
                print(f"    Parent: {guid_info['parent_guid']}")
                print()

            if len(guids) > 5:
                print(f"  ... and {len(guids) - 5} more")


if __name__ == "__main__":
    main()