"""
Get information about hotlink modules in the project.

This script retrieves details about all hotlink modules (external references)
placed in the current Archicad project.
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
        response = acc.ExecuteAddOnCommand(
            act.AddOnCommandId('TapirCommand'),
            {'command': 'GetHotlinks'}
        )

        if response and 'hotlinks' in response:
            hotlinks = parse_hotlinks(response['hotlinks'])

    except Exception as e:
        print(f"Error retrieving hotlink information: {e}")
        print("This feature requires TAPIR add-on support.")

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

    for node in hotlink_nodes:
        hotlink_info = {
            'location': node.get('location', 'Unknown'),
            'level': level,
            'has_children': 'children' in node and len(node.get('children', [])) > 0
        }

        result.append(hotlink_info)

        # Parse children recursively
        if 'children' in node and node['children']:
            children = parse_hotlinks(node['children'], level + 1)
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
    print("="*80)

    hotlinks = get_hotlink_info()

    if not hotlinks:
        print("\nNo hotlink modules found in the project.")
        print("(Or TAPIR add-on is not available)")
    else:
        format_hotlink_display(hotlinks)

    print("="*80)


if __name__ == "__main__":
    main()
