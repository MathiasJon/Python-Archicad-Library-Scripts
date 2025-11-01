"""
================================================================================
SCRIPT: Get Hotlink Info
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Project Information

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves comprehensive information about all hotlink modules (external
references) placed in the current Archicad project. This script provides:
- List of all hotlink modules
- Hotlink file locations/paths
- Hierarchical structure (hotlinks with nested hotlinks)
- Parent-child relationships
- Total count of hotlinks

Hotlinks are external Archicad files referenced in the current project, useful
for modular design and collaboration workflows.

WORKING VERSION - Correctly handles TAPIR GetHotlinks response format.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)
- An Archicad project must be open

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetHotlinks
  Returns all hotlink modules with their hierarchical structure

[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Data Types]
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AddOnCommandId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project containing hotlinks
2. Run this script
3. Review hotlink modules and their hierarchy

No configuration needed - script reads all hotlink information.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Total number of hotlink modules
- Each hotlink's location/path
- Hierarchical level (indented display)
- Nested hotlink indicators

Example output (No hotlinks):
  ════════════════════════════════════════════════════════════
  HOTLINK MODULE INFORMATION
  ════════════════════════════════════════════════════════════
  
  No hotlink modules found in the project.

Example output (With hotlinks):
  ════════════════════════════════════════════════════════════
  HOTLINK MODULE INFORMATION
  ════════════════════════════════════════════════════════════
  
  Found 3 hotlink module(s):
  
  • Hotlink 1
    Location: /Projects/Structure.pln
  
    └─ Hotlink 2
       Location: /Projects/Foundation.pln
       Has sub-hotlinks: Yes
  
      └─ Hotlink 3
         Location: /Projects/Details.pln

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
HOTLINK MODULES:
- External Archicad files referenced in the current project
- Can be nested (hotlinks within hotlinks)
- Location shows file path or server location
- Used for modular design and team collaboration

HIERARCHICAL STRUCTURE:
- Parent hotlinks can contain child hotlinks
- Script recursively parses all levels
- Indentation shows nesting depth
- Has sub-hotlinks indicator when applicable

HOTLINK WORKFLOW:
- Design reusable components in separate files
- Reference them as hotlinks in master project
- Updates to hotlink file reflect in all projects
- Common for: structure, MEP, landscaping, etc.

USE CASES:
- Verify hotlink placement
- Audit external references
- Document project structure
- Check for broken hotlinks
- Understand project dependencies

TROUBLESHOOTING:
- If command fails: Ensure Tapir is installed
- If no hotlinks shown: Project may not use hotlinks
- Missing locations: Hotlink file may be moved/missing
- Parse errors: Check Tapir version compatibility

LIMITATIONS:
- Requires Tapir Add-On
- Shows location but not detailed hotlink settings
- No validation of hotlink file existence

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 154_get_project_info.py (project information)
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


def get_hotlink_info():
    """
    Retrieve information about all hotlink modules in the project.

    Returns:
        List of dictionaries containing hotlink module information
    """
    hotlinks = []

    try:
        # Execute TAPIR command to get hotlinks
        command_id = act.AddOnCommandId('TapirCommand', 'GetHotlinks')
        response = acc.ExecuteAddOnCommand(command_id, {})

        # The GetHotlinks command returns data directly without a 'succeeded' field
        # when successful. If there's an error, it will have 'succeeded': False

        # Check if this is an error response
        if isinstance(response, dict):
            # Error case: {'succeeded': False, 'error': {...}}
            if 'succeeded' in response and not response['succeeded']:
                error_msg = response.get('error', {}).get(
                    'message', 'Unknown error')
                print(f"✗ Command failed: {error_msg}")
                return hotlinks

            # Success case: {'hotlinks': [...]}
            if 'hotlinks' in response:
                print("✓ Successfully retrieved hotlinks data")
                hotlink_data = response['hotlinks']

                if hotlink_data:
                    hotlinks = parse_hotlinks(hotlink_data)
                else:
                    print("  (No hotlinks found in the project)")
            else:
                print("⚠ Unexpected response format")
                print(f"  Response keys: {list(response.keys())}")

        elif hasattr(response, 'hotlinks'):
            # Handle object-type response
            print("✓ Successfully retrieved hotlinks data")
            hotlink_data = response.hotlinks

            if hotlink_data:
                hotlinks = parse_hotlinks(hotlink_data)
            else:
                print("  (No hotlinks found in the project)")

        else:
            print(f"⚠ Unexpected response type: {type(response)}")

    except Exception as e:
        print(f"✗ Error executing GetHotlinks command: {e}")
        import traceback
        traceback.print_exc()

    return hotlinks


def parse_hotlinks(hotlink_nodes, level=0):
    """
    Recursively parse hotlink nodes with their children.

    Args:
        hotlink_nodes: List of hotlink nodes
        level: Current hierarchy level

    Returns:
        List of parsed hotlink dictionaries
    """
    result = []

    if not hotlink_nodes:
        return result

    # Handle if hotlink_nodes is not a list
    if not isinstance(hotlink_nodes, (list, tuple)):
        try:
            hotlink_nodes = list(hotlink_nodes)
        except TypeError:
            print(
                f"Warning: hotlink_nodes is not iterable: {type(hotlink_nodes)}")
            return result

    for node in hotlink_nodes:
        # Handle both dict and object types
        try:
            if isinstance(node, dict):
                location = node.get('location', 'Unknown')
                children_data = node.get('children', [])
            else:
                location = getattr(node, 'location', 'Unknown')
                children_data = getattr(node, 'children', [])
        except (AttributeError, TypeError) as e:
            print(f"Warning: Could not parse node: {e}")
            location = 'Unknown'
            children_data = []

        hotlink_info = {
            'location': location,
            'level': level,
            'has_children': bool(children_data) and len(children_data) > 0
        }

        result.append(hotlink_info)

        # Parse children recursively
        if children_data:
            children = parse_hotlinks(children_data, level + 1)
            result.extend(children)

    return result


def format_hotlink_display(hotlinks):
    """Format hotlinks information for display."""
    if not hotlinks:
        print("\nNo hotlink modules found in the project.")
        return

    print(f"\nFound {len(hotlinks)} hotlink module(s):\n")

    for idx, hotlink in enumerate(hotlinks, 1):
        indent = "  " * hotlink['level']
        marker = "└─" if hotlink['level'] > 0 else "•"

        print(f"{indent}{marker} Hotlink {idx}")
        print(f"{indent}  Location: {hotlink['location']}")

        if hotlink['has_children']:
            print(f"{indent}  Has sub-hotlinks: Yes")

        print()


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("HOTLINK MODULE INFORMATION")
    print("="*80 + "\n")

    hotlinks = get_hotlink_info()

    if hotlinks:
        format_hotlink_display(hotlinks)
    else:
        print("\nNo hotlink modules found in the project.")

    print("="*80)


if __name__ == "__main__":
    main()