"""
================================================================================
SCRIPT: Get Element GDL Parameters
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Element Information - Library Parts

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves GDL (Geometric Description Language) parameters for selected library
part elements. This script provides:
- Complete list of GDL parameters per element
- Parameter names, values, and types
- Grouped display (numeric, text, boolean)
- Support for all library part types (objects, doors, windows, lamps)
- Detailed parameter information

GDL parameters control the behavior and appearance of library parts.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)
- Library part elements must be selected (Objects, Doors, Windows, Lamps)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetGDLParametersOfElements
  Returns GDL parameters for library part elements

[Official API]
- ACConnection.connect()
- GetSelectedElements()
- GetTypesOfElements(elementIds)
- ExecuteAddOnCommand(addOnCommandId)

[Data Types]
- AddOnCommandId
- ElementId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Select library part elements (Objects, Doors, Windows, Lamps)
3. Run this script
4. Review GDL parameters for each element

No configuration needed - script automatically extracts all parameters.

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 163_get_favorites_by_type.py (favorites management)
- 164_get_loaded_libraries.py (library information)
- 111_get_element_full_info.py (detailed element info)
- 151_check_tapir_version.py (verify Tapir installation)

================================================================================
"""

from archicad import ACConnection
import json

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def get_gdl_parameters(element_ids):
    """
    Get GDL parameters for given elements.

    Args:
        element_ids: List of ElementId objects

    Returns:
        List of GDL parameter dictionaries for each element
    """
    try:
        # Create proper AddOnCommandId
        command_id = act.AddOnCommandId(
            'TapirCommand', 'GetGDLParametersOfElements')

        # Prepare parameters - elements need to be in correct format
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

        print("Retrieving GDL parameters...")
        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error response
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ Command failed: {error_msg}")
            return None

        print("✓ Successfully retrieved GDL parameters\n")

        # Parse response
        if isinstance(response, dict):
            # The response key is 'gdlParametersOfElements' (confirmed from your output)
            if 'gdlParametersOfElements' in response:
                return response['gdlParametersOfElements']
            elif 'gdlParameters' in response:
                return response['gdlParameters']

        print("⚠ Unexpected response format")
        print(
            f"Response keys: {list(response.keys()) if isinstance(response, dict) else 'not a dict'}")
        return None

    except Exception as e:
        print(f"✗ Error retrieving GDL parameters: {e}")
        import traceback
        traceback.print_exc()
        return None


def format_parameter_value(value):
    """Format a parameter value for display."""
    if isinstance(value, (int, float)):
        return f"{value}"
    elif isinstance(value, str):
        if len(value) > 60:
            return f'"{value[:57]}..."'
        return f'"{value}"'
    elif isinstance(value, bool):
        return str(value)
    elif isinstance(value, (list, dict)):
        return json.dumps(value)
    else:
        return str(value)


def display_gdl_parameters(element_ids, gdl_params_list):
    """
    Display GDL parameters in a formatted way.

    Args:
        element_ids: List of ElementId objects
        gdl_params_list: List of GDL parameter data for each element
    """
    print("\n" + "="*80)
    print("GDL PARAMETERS REPORT")
    print("="*80)

    if not gdl_params_list:
        print("\n⚠ No GDL parameters retrieved")
        print("\nPossible reasons:")
        print("  • Elements are not library parts (walls, slabs, etc. don't have GDL)")
        print("  • Library parts have no editable parameters")
        print("  • TAPIR version doesn't support this command")
        return

    # Get element types for display
    try:
        element_types = acc.GetTypesOfElements(element_ids)
    except:
        element_types = [None] * len(element_ids)

    for idx, (elem_id, gdl_data) in enumerate(zip(element_ids, gdl_params_list), 1):
        print(f"\n{'─'*80}")
        print(f"ELEMENT {idx}")
        print(f"{'─'*80}")

        # Show element info
        elem_type = element_types[idx -
                                  1].typeOfElement.elementType if element_types[idx-1] else "Unknown"
        print(f"Type: {elem_type}")
        print(f"GUID: {elem_id.guid}")

        # Debug: Show data structure
        print(f"\nDEBUG - Data type: {type(gdl_data)}")
        if isinstance(gdl_data, dict):
            print(f"DEBUG - Dict keys: {list(gdl_data.keys())}")
        elif isinstance(gdl_data, list):
            print(f"DEBUG - List length: {len(gdl_data)}")
            if gdl_data and len(gdl_data) > 0:
                print(f"DEBUG - First item type: {type(gdl_data[0])}")
                if isinstance(gdl_data[0], dict):
                    print(
                        f"DEBUG - First item keys: {list(gdl_data[0].keys())}")

        # Check if we have parameters
        if not gdl_data:
            print("\n⚠ No GDL parameters available for this element")
            print("  (Element might not be a library part)")
            continue

        # Parse GDL data structure
        # The structure is: {'parameters': [{'name': '...', 'value': '...', 'type': '...'}, ...]}

        # Extract the parameters
        if isinstance(gdl_data, dict) and 'parameters' in gdl_data:
            # Get the parameters list from the dict
            params_list = gdl_data['parameters']
        elif isinstance(gdl_data, list):
            # Already a list
            params_list = gdl_data
        elif isinstance(gdl_data, dict):
            # Plain dict format (less common)
            params_list = gdl_data
        else:
            print(f"\n⚠ Unexpected parameter format: {type(gdl_data)}")
            print(f"  Data: {gdl_data}")
            continue

        # Now handle the params_list
        if isinstance(params_list, list):
            # List format: [{'name': '...', 'value': '...', 'type': '...'}, ...]
            if params_list and len(params_list) > 0:
                print(f"\n✓ Found {len(params_list)} GDL parameter(s):\n")

                # Group by type for better display
                numeric = []
                text = []
                boolean = []
                other = []

                for param in params_list:
                    if isinstance(param, dict):
                        name = param.get('name', 'unknown')
                        value = param.get('value', 'N/A')
                        param_type = param.get('type', '')

                        # Group by type
                        if isinstance(value, (int, float)) or param_type in ['Real', 'Integer', 'Length', 'Angle']:
                            numeric.append((name, value, param_type))
                        elif isinstance(value, str) or param_type == 'String':
                            text.append((name, value, param_type))
                        elif isinstance(value, bool) or param_type == 'Boolean':
                            boolean.append((name, value, param_type))
                        else:
                            other.append((name, value, param_type))

                # Display numeric
                if numeric:
                    print("  Numeric Parameters:")
                    for name, value, ptype in numeric:
                        type_str = f" [{ptype}]" if ptype else ""
                        print(f"    {name}{type_str} = {value}")
                    print()

                # Display text
                if text:
                    print("  Text Parameters:")
                    for name, value, ptype in text:
                        formatted_value = format_parameter_value(value)
                        print(f"    {name} = {formatted_value}")
                    print()

                # Display boolean
                if boolean:
                    print("  Boolean Parameters:")
                    for name, value, ptype in boolean:
                        print(f"    {name} = {value}")
                    print()

                # Display other
                if other:
                    print("  Other Parameters:")
                    for name, value, ptype in other:
                        type_str = f" [{ptype}]" if ptype else ""
                        print(
                            f"    {name}{type_str} = {format_parameter_value(value)}")
            else:
                print("\n⚠ Parameter list is empty")

        elif isinstance(params_list, dict):
            # Dictionary format (less common): {'paramName': value, ...}
            if params_list and len(params_list) > 0:
                print(f"\n✓ Found {len(params_list)} GDL parameter(s):\n")

                # Group by type
                numeric = []
                text = []
                other = []

                for name, value in sorted(params_list.items()):
                    if isinstance(value, (int, float)):
                        numeric.append((name, value))
                    elif isinstance(value, str):
                        text.append((name, value))
                    else:
                        other.append((name, value))

                # Display numeric
                if numeric:
                    print("  Numeric Parameters:")
                    for name, value in numeric:
                        print(f"    {name} = {value}")
                    print()

                # Display text
                if text:
                    print("  Text Parameters:")
                    for name, value in text:
                        print(f"    {name} = {format_parameter_value(value)}")
                    print()

                # Display other
                if other:
                    print("  Other Parameters:")
                    for name, value in other:
                        print(f"    {name} = {format_parameter_value(value)}")
            else:
                print("\n⚠ Parameter dictionary is empty")

        else:
            print(f"\n⚠ Unexpected params format: {type(params_list)}")

    print("\n" + "="*80)


def main():
    """Main function."""
    print("\n" + "="*80)
    print("GDL PARAMETERS EXTRACTION")
    print("="*80 + "\n")

    # Get selected elements
    selected = acc.GetSelectedElements()

    if len(selected) == 0:
        print("⚠ No elements selected")
        print("\nPlease select library parts in Archicad:")
        print("  • Objects")
        print("  • Doors")
        print("  • Windows")
        print("  • Lamps")
        print("  • Furniture")
        return

    print(f"Analyzing {len(selected)} selected element(s)...\n")

    # Get element IDs
    element_ids = [elem.elementId for elem in selected]

    # Get GDL parameters
    gdl_params_list = get_gdl_parameters(element_ids)

    # Display results
    display_gdl_parameters(element_ids, gdl_params_list)

    # Additional tip
    print("\n💡 TIP:")
    print("   If no parameters are shown, the element might not be a library part")
    print("   or might not have editable GDL parameters.")
    print("\n   Try selecting an Object, Door, or Window from the library.")


if __name__ == "__main__":
    main()