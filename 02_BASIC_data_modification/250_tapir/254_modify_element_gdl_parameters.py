"""
Modify GDL parameters of elements.

This script allows modification of GDL (Geometric Description Language) parameters
for library part elements such as doors, windows, and objects.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def get_element_gdl_parameters(element_guid):
    """
    Get all GDL parameters of an element.
    
    Args:
        element_guid: GUID of the element
        
    Returns:
        Dictionary of GDL parameters
    """
    try:
        # Execute TAPIR command to get GDL parameters
        response = acc.ExecuteAddOnCommand('TapirCommand', {
            'command': 'GetGDLParameters',
            'elementId': element_guid
        })
        
        if response and 'parameters' in response:
            return response['parameters']
        else:
            print("No GDL parameters found or element is not a library part.")
            return {}
            
    except Exception as e:
        print(f"Error retrieving GDL parameters: {e}")
        return {}


def modify_gdl_parameter(element_guid, parameter_name, new_value):
    """
    Modify a single GDL parameter of an element.
    
    Args:
        element_guid: GUID of the element
        parameter_name: Name of the GDL parameter
        new_value: New value for the parameter
        
    Returns:
        Boolean indicating success
    """
    try:
        # Execute TAPIR command to modify parameter
        response = acc.ExecuteAddOnCommand('TapirCommand', {
            'command': 'SetGDLParameter',
            'elementId': element_guid,
            'parameterName': parameter_name,
            'value': new_value
        })
        
        if response and response.get('success', False):
            print(f"✓ Parameter '{parameter_name}' set to '{new_value}'")
            return True
        else:
            error_msg = response.get('error', 'Unknown error') if response else 'No response'
            print(f"✗ Failed to set parameter: {error_msg}")
            return False
            
    except Exception as e:
        print(f"Error modifying GDL parameter: {e}")
        print("This feature requires TAPIR add-on support.")
        return False


def modify_multiple_gdl_parameters(element_guid, parameters):
    """
    Modify multiple GDL parameters at once.
    
    Args:
        element_guid: GUID of the element
        parameters: Dictionary of parameter_name: new_value pairs
        
    Returns:
        Dictionary with success status for each parameter
    """
    results = {}
    
    for param_name, new_value in parameters.items():
        success = modify_gdl_parameter(element_guid, param_name, new_value)
        results[param_name] = success
    
    return results


def batch_modify_gdl_parameters(element_guids, parameter_name, new_value):
    """
    Modify a GDL parameter for multiple elements.
    
    Args:
        element_guids: List of element GUIDs
        parameter_name: Name of the GDL parameter
        new_value: New value for the parameter
        
    Returns:
        Number of successfully modified elements
    """
    success_count = 0
    
    for guid in element_guids:
        if modify_gdl_parameter(guid, parameter_name, new_value):
            success_count += 1
    
    return success_count


def modify_selected_elements_gdl(parameter_name, new_value):
    """
    Modify a GDL parameter for all currently selected elements.
    
    Args:
        parameter_name: Name of the GDL parameter
        new_value: New value for the parameter
        
    Returns:
        Number of successfully modified elements
    """
    # Get selected elements
    selected = acc.GetSelectedElements()
    
    if not selected:
        print("No elements selected.")
        return 0
    
    element_guids = [elem.elementId.guid for elem in selected]
    
    print(f"Modifying '{parameter_name}' for {len(element_guids)} selected element(s)...")
    
    success_count = batch_modify_gdl_parameters(element_guids, parameter_name, new_value)
    
    print(f"\n✓ Successfully modified {success_count}/{len(element_guids)} element(s)")
    
    return success_count


def display_gdl_parameters(element_guid):
    """
    Display all GDL parameters of an element in a formatted way.
    
    Args:
        element_guid: GUID of the element
    """
    parameters = get_element_gdl_parameters(element_guid)
    
    if not parameters:
        print("No GDL parameters available for this element.")
        return
    
    print("\n" + "="*80)
    print("GDL PARAMETERS")
    print("="*80 + "\n")
    
    # Group parameters by type
    by_type = {}
    for param_name, param_data in parameters.items():
        param_type = param_data.get('type', 'Unknown')
        if param_type not in by_type:
            by_type[param_type] = []
        by_type[param_type].append((param_name, param_data))
    
    for param_type, params in by_type.items():
        print(f"\n{param_type} Parameters:")
        print("─"*80)
        
        for param_name, param_data in params:
            value = param_data.get('value', 'N/A')
            description = param_data.get('description', '')
            
            print(f"  {param_name}: {value}")
            if description:
                print(f"    → {description}")


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("MODIFY GDL PARAMETERS")
    print("="*80)
    
    # Get selected elements
    selected = acc.GetSelectedElements()
    
    if not selected:
        print("\nNo elements selected.")
        print("Please select a library part element (door, window, object) and run again.")
        return
    
    first_element = selected[0]
    element_guid = first_element.elementId.guid
    element_type = first_element.elementId.elementType
    
    print(f"\nAnalyzing first selected element:")
    print(f"Type: {element_type}")
    print(f"GUID: {element_guid}")
    
    # Display current GDL parameters
    display_gdl_parameters(element_guid)
    
    # Example: Modify a common parameter
    print("\n" + "─"*80)
    print("Example: Modifying GDL parameters")
    print("─"*80 + "\n")
    
    # Common door/window parameters (examples - adjust based on your library)
    example_modifications = {
        'gs_frame_width': 50,  # Frame width
        'gs_sash_width': 45,   # Sash width
        # Add more parameters as needed
    }
    
    print("Example parameter modifications (uncomment to activate):")
    for param, value in example_modifications.items():
        print(f"  {param} = {value}")
    
    # Uncomment to actually modify:
    # results = modify_multiple_gdl_parameters(element_guid, example_modifications)
    # 
    # print("\nModification results:")
    # for param, success in results.items():
    #     status = "✓" if success else "✗"
    #     print(f"  {status} {param}")
    
    print("\n" + "="*80)
    print("Note: Uncomment the modification code to apply changes.")
    print("="*80)


if __name__ == "__main__":
    main()
