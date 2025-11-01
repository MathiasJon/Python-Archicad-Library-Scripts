"""
================================================================================
SCRIPT: Check Element Intersection (Collision Detection)
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Element Analysis - Clash Detection

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Detects geometric collisions between elements for clash detection and BIM
coordination purposes. This script provides:
- Body collision detection (volume overlap)
- Clearance collision detection (proximity check)
- Collision report with detailed information
- Multiple analysis modes (all vs all, groups comparison)
- Export collision pairs for further processing

Based on TAPIR API documentation v1.2.2

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)
- At least 2 elements must be selected

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetCollisions
  Detects geometric collisions between two groups of elements

[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetSelectedElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSelectedElements

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Data Types]
- AddOnCommandId
- ElementId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select elements to check for collisions (minimum 2)
3. Run this script
4. Review collision detection report

Configuration parameters:
- volume_tolerance: Minimum volume for collision (default: 0.001 m³)
- perform_surface_check: Enable surface collision check (default: False)
- surface_tolerance: Minimum surface area (default: 0.001 m²)

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Number of collisions found
- Collision types (Body / Clearance)
- Element GUIDs for each collision
- Detailed collision report
- Export-ready collision pairs

Example output (Collisions found):
  ════════════════════════════════════════════════════════════
  COLLISION DETECTION REPORT
  ════════════════════════════════════════════════════════════
  
  ⚠  Found 3 collision(s):
  
  Collision Types:
    • Body Collisions: 2
    • Clearance Collisions: 1
  
  ────────────────────────────────────────────────────────────
  Collision Details:
  ────────────────────────────────────────────────────────────
  
  1. Collision between elements
     Element 1: 12345678-1234-1234-1234-123456789ABC
     Element 2: 23456789-2345-2345-2345-23456789ABCD
     Type: Body Collision
  
  2. Collision between elements
     Element 1: 34567890-3456-3456-3456-34567890ABCD
     Element 2: 45678901-4567-4567-4567-45678901ABCD
     Type: Body Collision, Clearance Collision

Example output (No collisions):
  ════════════════════════════════════════════════════════════
  COLLISION DETECTION REPORT
  ════════════════════════════════════════════════════════════
  
  ✓ No collisions detected!

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
COLLISION TYPES:
- Body Collision: Actual geometric overlap between element volumes
- Clearance Collision: Elements are too close (within clearance distance)
- Both types can occur simultaneously

COLLISION DETECTION PARAMETERS:
- volumeTolerance: Minimum intersection volume to report (default: 0.001 m³)
  - Smaller values = more sensitive detection
  - Filters out tiny overlaps that might be modeling tolerance
  
- performSurfaceCheck: Enable clearance/proximity checking (default: False)
  - When true, checks if elements are too close without touching
  - Useful for MEP coordination and clearance requirements
  
- surfaceTolerance: Minimum clearance distance (default: 0.001 m²)
  - Only used when performSurfaceCheck is true
  - Elements closer than this distance trigger clearance collision

DETECTION MODES:
1. All vs All: Checks every element against every other element
   - Use for comprehensive clash detection
   - Can find unexpected collisions
   
2. Group vs Group: Checks two specific groups against each other
   - Use for targeted analysis (e.g., structure vs MEP)
   - More efficient for large models

ANALYSIS STRATEGIES:
- Structure vs Structure: Find structural clashes
- Structure vs MEP: Find MEP penetrations
- Architecture vs Structure: Find coordination issues
- Room vs Equipment: Find space conflicts

PERFORMANCE CONSIDERATIONS:
- Collision detection is computationally intensive
- Large selections may take time to process
- Consider checking subsets for large models
- Group-based checking is more efficient than all-vs-all

USE CASES:
- BIM coordination and clash detection
- Quality control before construction
- MEP coordination with structure
- Space planning verification
- Detect modeling errors
- Construction sequencing analysis
- Export clash reports for team review

COLLISION REPORT:
- Shows all detected collisions
- Identifies element pairs in conflict
- Specifies collision type(s)
- Can be exported for documentation
- GUIDs allow element identification in Archicad

EXPORT FUNCTIONALITY:
- Collision pairs exported as tuples
- Format: (guid1, guid2, collision_type)
- Ready for CSV export or further processing
- Can be used to generate issue reports

WORKFLOW INTEGRATION:
1. Run collision detection
2. Review collision report
3. Export collision pairs if needed
4. Use GUIDs to locate elements in Archicad
5. Resolve clashes
6. Re-run to verify fixes

LIMITATIONS:
- Requires Tapir Add-On
- Limited to selected elements
- Cannot detect all coordination issues
- Tolerance settings affect results
- Very small overlaps might be modeling tolerance

TOLERANCE RECOMMENDATIONS:
- General BIM: volume_tolerance = 0.001 m³ (1 liter)
- Precision work: volume_tolerance = 0.0001 m³ (100 ml)
- MEP clearance: perform_surface_check = True, surface_tolerance = 0.05 m²

TROUBLESHOOTING:
- If command fails: Ensure Tapir Add-On is installed
- If no collisions found: Check elements actually overlap
- If too many results: Increase tolerance values
- If missing clashes: Decrease tolerance values
- Performance issues: Analyze smaller groups

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 158_get_element_3D_bounding_box.py (spatial analysis)
- 160_get_connected_elements.py (element relationships)
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


def get_collisions(elements_group1, elements_group2,
                   volume_tolerance=0.001,
                   perform_surface_check=False,
                   surface_tolerance=0.001):
    """
    Detect collisions between two groups of elements.

    Args:
        elements_group1: List of ElementId objects (first group)
        elements_group2: List of ElementId objects (second group)
        volume_tolerance: Minimum volume for collision detection (default: 0.001 m³)
        perform_surface_check: Enable surface collision check (default: False)
        surface_tolerance: Minimum surface area for collision (default: 0.001 m²)

    Returns:
        List of collision dictionaries
    """
    try:
        # Create proper AddOnCommandId
        command_id = act.AddOnCommandId('TapirCommand', 'GetCollisions')

        # Prepare input parameters according to API docs
        elements_param1 = []
        for elem_id in elements_group1:
            elements_param1.append({
                'elementId': {
                    'guid': str(elem_id.guid)  # Convert UUID to string
                }
            })

        elements_param2 = []
        for elem_id in elements_group2:
            elements_param2.append({
                'elementId': {
                    'guid': str(elem_id.guid)  # Convert UUID to string
                }
            })

        parameters = {
            'elementsGroup1': elements_param1,
            'elementsGroup2': elements_param2,
            'settings': {
                'volumeTolerance': volume_tolerance,
                'performSurfaceCheck': perform_surface_check,
                'surfaceTolerance': surface_tolerance
            }
        }

        print(
            f"Checking collisions between {len(elements_group1)} and {len(elements_group2)} element(s)...")
        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error response
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ Command failed: {error_msg}")
            return None

        print("✓ Successfully checked for collisions\n")

        # Parse the response
        if isinstance(response, dict) and 'collisions' in response:
            return parse_collisions_dict(response['collisions'])
        elif hasattr(response, 'collisions'):
            return parse_collisions_object(response.collisions)
        else:
            print("⚠ Unexpected response format")
            return None

    except Exception as e:
        print(f"✗ Error detecting collisions: {e}")
        import traceback
        traceback.print_exc()
        return None


def parse_collisions_dict(collisions_data):
    """Parse collisions from dictionary response."""
    if not collisions_data:
        return []

    results = []

    for collision in collisions_data:
        if isinstance(collision, dict):
            elem_id1 = collision.get('elementId1', {}).get('guid')
            elem_id2 = collision.get('elementId2', {}).get('guid')
            has_body = collision.get('hasBodyCollision', False)
            has_clearance = collision.get('hasClearenceCollision', False)

            results.append({
                'element1_guid': elem_id1,
                'element2_guid': elem_id2,
                'has_body_collision': has_body,
                'has_clearance_collision': has_clearance
            })

    return results


def parse_collisions_object(collisions_data):
    """Parse collisions from object response."""
    if not collisions_data:
        return []

    results = []

    # Convert to list if needed
    if hasattr(collisions_data, '__iter__'):
        collision_list = list(collisions_data)
    else:
        collision_list = [collisions_data]

    for collision in collision_list:
        elem_id1 = None
        elem_id2 = None

        if hasattr(collision, 'elementId1'):
            elem1 = collision.elementId1
            if hasattr(elem1, 'guid'):
                elem_id1 = str(elem1.guid)

        if hasattr(collision, 'elementId2'):
            elem2 = collision.elementId2
            if hasattr(elem2, 'guid'):
                elem_id2 = str(elem2.guid)

        has_body = getattr(collision, 'hasBodyCollision', False)
        has_clearance = getattr(collision, 'hasClearenceCollision', False)

        results.append({
            'element1_guid': elem_id1,
            'element2_guid': elem_id2,
            'has_body_collision': has_body,
            'has_clearance_collision': has_clearance
        })

    return results


def detect_collisions_in_selection(volume_tolerance=0.001, perform_surface_check=False):
    """
    Detect collisions between all selected elements (all vs all).

    Args:
        volume_tolerance: Minimum volume for collision detection
        perform_surface_check: Enable surface collision check

    Returns:
        List of collision dictionaries
    """
    # Get selected elements
    selected = acc.GetSelectedElements()

    if len(selected) < 2:
        print("Need at least 2 elements selected for collision detection.")
        return None

    print(f"Analyzing {len(selected)} selected element(s) for collisions...\n")

    # Get element IDs
    element_ids = [elem.elementId for elem in selected]

    # Check all elements against all elements
    # (This is the same as checking each pair, but more efficient)
    collisions = get_collisions(
        element_ids,
        element_ids,
        volume_tolerance=volume_tolerance,
        perform_surface_check=perform_surface_check
    )

    if collisions:
        # Filter out self-collisions (same element in both groups)
        filtered = [c for c in collisions if c['element1_guid']
                    != c['element2_guid']]
        return filtered

    return collisions


def detect_collisions_between_groups(group1_elements, group2_elements,
                                     volume_tolerance=0.001,
                                     perform_surface_check=False):
    """
    Detect collisions between two specific groups of elements.

    Args:
        group1_elements: List of ElementId objects (first group)
        group2_elements: List of ElementId objects (second group)
        volume_tolerance: Minimum volume for collision detection
        perform_surface_check: Enable surface collision check

    Returns:
        List of collision dictionaries
    """
    if not group1_elements or not group2_elements:
        print("Both groups must have at least one element.")
        return None

    return get_collisions(
        group1_elements,
        group2_elements,
        volume_tolerance=volume_tolerance,
        perform_surface_check=perform_surface_check
    )


def split_selection_into_groups():
    """
    Helper function to split selected elements into two groups.
    First half vs second half.

    Returns:
        Tuple of (group1_ids, group2_ids)
    """
    selected = acc.GetSelectedElements()

    if len(selected) < 2:
        return None, None

    midpoint = len(selected) // 2

    group1 = [elem.elementId for elem in selected[:midpoint]]
    group2 = [elem.elementId for elem in selected[midpoint:]]

    return group1, group2


def generate_collision_report(collisions):
    """
    Generate a formatted collision detection report.

    Args:
        collisions: List of collision dictionaries
    """
    print("\n" + "="*80)
    print("COLLISION DETECTION REPORT")
    print("="*80)

    if not collisions or len(collisions) == 0:
        print("\n✓ No collisions detected!")
        return

    print(f"\n⚠ Found {len(collisions)} collision(s):\n")

    # Count collision types
    body_collisions = sum(1 for c in collisions if c.get(
        'has_body_collision', False))
    clearance_collisions = sum(1 for c in collisions if c.get(
        'has_clearance_collision', False))

    print(f"Collision Types:")
    print(f"  • Body Collisions: {body_collisions}")
    print(f"  • Clearance Collisions: {clearance_collisions}")

    print("\n" + "─"*80)
    print("Collision Details:")
    print("─"*80 + "\n")

    for idx, collision in enumerate(collisions, 1):
        print(f"{idx}. Collision between elements")
        print(f"   Element 1: {collision['element1_guid']}")
        print(f"   Element 2: {collision['element2_guid']}")

        collision_types = []
        if collision.get('has_body_collision', False):
            collision_types.append("Body Collision")
        if collision.get('has_clearance_collision', False):
            collision_types.append("Clearance Collision")

        print(
            f"   Type: {', '.join(collision_types) if collision_types else 'Unknown'}")
        print()

    print("="*80)


def export_collision_pairs(collisions):
    """
    Export collision pairs for further processing.

    Args:
        collisions: List of collision dictionaries

    Returns:
        List of tuples (element1_guid, element2_guid, collision_type)
    """
    pairs = []

    if not collisions:
        return pairs

    for collision in collisions:
        guid1 = collision.get('element1_guid')
        guid2 = collision.get('element2_guid')

        collision_type = []
        if collision.get('has_body_collision', False):
            collision_type.append('body')
        if collision.get('has_clearance_collision', False):
            collision_type.append('clearance')

        if guid1 and guid2:
            pairs.append((guid1, guid2, ','.join(collision_type)))

    return pairs


def main():
    """Main function to demonstrate the script."""

    # Example 1: Detect all collisions in selection
    print("\n" + "="*80)
    print("EXAMPLE 1: Detect All Collisions in Selection")
    print("="*80 + "\n")

    collisions = detect_collisions_in_selection(
        volume_tolerance=0.001,  # 0.001 m³ minimum volume
        perform_surface_check=False  # Don't check surface collisions
    )

    if collisions is not None:
        generate_collision_report(collisions)
    else:
        print("No collisions detected or command failed.")

    # Example 2: Detect collisions between two groups
    print("\n" + "="*80)
    print("EXAMPLE 2: Detect Collisions Between Two Groups")
    print("="*80 + "\n")
    print("(Splitting selection into two groups: first half vs second half)\n")

    group1, group2 = split_selection_into_groups()

    if group1 and group2:
        print(f"Group 1: {len(group1)} element(s)")
        print(f"Group 2: {len(group2)} element(s)\n")

        group_collisions = detect_collisions_between_groups(
            group1,
            group2,
            volume_tolerance=0.001,
            perform_surface_check=True
        )

        if group_collisions is not None:
            generate_collision_report(group_collisions)
    else:
        print("Need at least 2 elements selected for group comparison.")

    # Example 3: Export collision pairs
    if collisions:
        print("\n" + "─"*80)
        print("EXAMPLE 3: Export Collision Pairs")
        print("─"*80 + "\n")

        pairs = export_collision_pairs(collisions)
        print(f"Exported {len(pairs)} collision pair(s):\n")

        for idx, (guid1, guid2, ctype) in enumerate(pairs[:5], 1):
            print(f"{idx}. {guid1} ←→ {guid2}")
            print(f"   Type: {ctype}\n")

        if len(pairs) > 5:
            print(f"   ... and {len(pairs) - 5} more")


if __name__ == "__main__":
    main()