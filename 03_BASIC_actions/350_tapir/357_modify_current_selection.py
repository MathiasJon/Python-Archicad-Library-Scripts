"""
Modify the current element selection.

This script provides functions to add, remove, or modify the current selection
in Archicad using TAPIR commands:
- GetSelectedElements (v0.1.0): Get current selection
- ChangeSelectionOfElements (v1.0.7): Add/remove elements from selection

Based on TAPIR API documentation.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def get_selected_elements():
    """
    Get currently selected elements.

    Returns:
        List of element IDs
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'GetSelectedElements')
        response = acc.ExecuteAddOnCommand(command_id, {})

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ GetSelectedElements failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'elements' in response:
            elements = response['elements']
            # Extract element IDs
            element_ids = []
            for elem in elements:
                if isinstance(elem, dict) and 'elementId' in elem:
                    guid = elem['elementId'].get('guid')
                    if guid:
                        element_ids.append(guid)
            return element_ids

        return []

    except Exception as e:
        print(f"✗ Error getting selected elements: {e}")
        return []


def change_selection(add_guids=None, remove_guids=None):
    """
    Add and/or remove elements from the current selection.

    Args:
        add_guids: List of GUIDs to add to selection (optional)
        remove_guids: List of GUIDs to remove from selection (optional)

    Returns:
        Dictionary with results
    """
    try:
        command_id = act.AddOnCommandId(
            'TapirCommand', 'ChangeSelectionOfElements')

        # Prepare parameters
        parameters = {}

        if add_guids:
            parameters['addElementsToSelection'] = [
                {'elementId': {'guid': str(guid)}} for guid in add_guids
            ]
        else:
            parameters['addElementsToSelection'] = []

        if remove_guids:
            parameters['removeElementsFromSelection'] = [
                {'elementId': {'guid': str(guid)}} for guid in remove_guids
            ]
        else:
            parameters['removeElementsFromSelection'] = []

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check response
        if isinstance(response, dict):
            add_results = response.get('executionResultsOfAddToSelection', [])
            remove_results = response.get(
                'executionResultsOfRemoveFromSelection', [])

            add_success = sum(
                1 for r in add_results if r.get('success', False))
            remove_success = sum(
                1 for r in remove_results if r.get('success', False))

            return {
                'added': add_success,
                'removed': remove_success,
                'add_total': len(add_guids) if add_guids else 0,
                'remove_total': len(remove_guids) if remove_guids else 0
            }

        return {'added': 0, 'removed': 0, 'add_total': 0, 'remove_total': 0}

    except Exception as e:
        print(f"✗ Error changing selection: {e}")
        return {'added': 0, 'removed': 0, 'add_total': 0, 'remove_total': 0}


def add_to_selection(element_guids):
    """
    Add elements to the current selection.

    Args:
        element_guids: List of element GUIDs to add

    Returns:
        Number of elements added
    """
    result = change_selection(add_guids=element_guids)

    print(
        f"✓ Added {result['added']}/{result['add_total']} element(s) to selection")

    return result['added']


def remove_from_selection(element_guids):
    """
    Remove elements from the current selection.

    Args:
        element_guids: List of element GUIDs to remove

    Returns:
        Number of elements removed
    """
    result = change_selection(remove_guids=element_guids)

    print(
        f"✓ Removed {result['removed']}/{result['remove_total']} element(s) from selection")

    return result['removed']


def select_all_of_type(element_type):
    """
    Select all elements of a specific type.

    Args:
        element_type: Type of element (e.g., 'Wall', 'Slab')

    Returns:
        Number of elements selected
    """
    try:
        # Get all elements of type
        elements = acc.GetElementsByType(element_type)

        if not elements:
            print(f"⚠ No elements of type '{element_type}' found")
            return 0

        # Get GUIDs
        guids = [elem.elementId.guid for elem in elements]

        # First, clear selection by removing all currently selected
        current = get_selected_elements()
        if current:
            change_selection(remove_guids=current)

        # Then add new selection
        result = change_selection(add_guids=guids)

        print(f"✓ Selected all {result['added']} {element_type}(s)")
        return result['added']

    except Exception as e:
        print(f"✗ Error selecting by type: {e}")
        return 0


def clear_selection():
    """
    Clear the current selection (deselect all).

    Returns:
        Number of elements deselected
    """
    current_guids = get_selected_elements()

    if not current_guids:
        print("✓ Selection already empty")
        return 0

    result = change_selection(remove_guids=current_guids)

    print(f"✓ Cleared selection ({result['removed']} element(s) deselected)")
    return result['removed']


def invert_selection():
    """
    Invert the current selection.

    Returns:
        Number of elements now selected
    """
    try:
        # Get current selection
        current_guids = set(get_selected_elements())

        # Get all elements
        all_elements = acc.GetAllElements()
        all_guids = [elem.elementId.guid for elem in all_elements]

        # Calculate inverted selection
        inverted_guids = [guid for guid in all_guids if str(
            guid) not in current_guids]

        # Clear current and add inverted
        if current_guids:
            change_selection(remove_guids=list(current_guids))

        result = change_selection(add_guids=inverted_guids)

        print(f"✓ Selection inverted")
        print(f"  Previously selected: {len(current_guids)}")
        print(f"  Now selected: {result['added']}")

        return result['added']

    except Exception as e:
        print(f"✗ Error inverting selection: {e}")
        return 0


def select_by_criteria(criteria_function):
    """
    Select elements based on a custom criteria function.

    Args:
        criteria_function: Function that takes an element and returns True/False

    Returns:
        Number of elements selected
    """
    try:
        # Get all elements
        all_elements = acc.GetAllElements()

        # Filter by criteria
        matching_elements = [
            elem for elem in all_elements if criteria_function(elem)]

        if not matching_elements:
            print("⚠ No elements match the criteria")
            return 0

        # Get GUIDs
        guids = [elem.elementId.guid for elem in matching_elements]

        # Clear and select
        current = get_selected_elements()
        if current:
            change_selection(remove_guids=current)

        result = change_selection(add_guids=guids)

        print(f"✓ Selected {result['added']} element(s) matching criteria")
        return result['added']

    except Exception as e:
        print(f"✗ Error selecting by criteria: {e}")
        return 0


def display_selection_info():
    """Display information about current selection."""
    current = get_selected_elements()

    print(f"\nCurrent selection: {len(current)} element(s)")

    if current and len(current) <= 10:
        print("\nSelected element GUIDs:")
        for idx, guid in enumerate(current, 1):
            print(f"  {idx}. {guid}")
    elif len(current) > 10:
        print(f"\nFirst 10 selected element GUIDs:")
        for idx, guid in enumerate(current[:10], 1):
            print(f"  {idx}. {guid}")
        print(f"  ... and {len(current) - 10} more")


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("MODIFY CURRENT SELECTION")
    print("="*80)

    # Show current selection
    display_selection_info()

    # Available operations
    print("\n" + "─"*80)
    print("Available Selection Operations:")
    print("─"*80)
    print("1. add_to_selection(guids) - Add elements to selection")
    print("2. remove_from_selection(guids) - Remove elements from selection")
    print("3. select_all_of_type('Wall') - Select all of a type")
    print("4. invert_selection() - Invert current selection")
    print("5. clear_selection() - Clear all selection")
    print("6. select_by_criteria(func) - Select by custom function")

    # Example 1: Select all walls
    print("\n" + "─"*80)
    print("EXAMPLE 1: Select All Walls")
    print("─"*80 + "\n")

    wall_count = select_all_of_type('Wall')

    # Example 2: Clear selection
    print("\n" + "─"*80)
    print("EXAMPLE 2: Clear Selection")
    print("─"*80 + "\n")

    clear_selection()

    # Example 3: Select by criteria (walls with specific property)
    print("\n" + "─"*80)
    print("EXAMPLE 3: Select Walls (using criteria)")
    print("─"*80 + "\n")

    # Select only walls
    # Note: Element type is accessed via the element's classification info
    def is_wall(element):
        # Get element type
        try:
            element_type = acc.GetTypesOfElements([element.elementId])[0]
            return element_type.typeOfElement.elementType == 'Wall'
        except:
            return False

    select_by_criteria(is_wall)

    print("\n" + "="*80)
    print("\n💡 TIP: Use add_to_selection() and remove_from_selection()")
    print("   to build complex selections programmatically")
    print("="*80)


if __name__ == "__main__":
    main()
