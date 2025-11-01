"""
Filter current selection by various criteria.

This script allows filtering the current selection based on element properties,
types, classifications, or other attributes.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def filter_by_type(element_guids, element_types):
    """
    Filter elements by type.
    
    Args:
        element_guids: List of element GUIDs to filter
        element_types: List of element types to keep (e.g., ['Wall', 'Slab'])
        
    Returns:
        List of filtered element GUIDs
    """
    filtered = []
    
    for guid in element_guids:
        element_id = act.ElementId(guid)
        
        if element_id.elementType in element_types:
            filtered.append(guid)
    
    return filtered


def filter_by_property(element_guids, property_id, value, comparison='equals'):
    """
    Filter elements by property value.
    
    Args:
        element_guids: List of element GUIDs to filter
        property_id: Property definition GUID
        value: Value to compare against
        comparison: Comparison operator ('equals', 'greater', 'less', 'contains')
        
    Returns:
        List of filtered element GUIDs
    """
    filtered = []
    
    # Get property values for all elements
    element_ids = [act.ElementId(guid) for guid in element_guids]
    
    property_ids = [act.PropertyId(property_id)]
    
    try:
        property_values = acc.GetPropertyValuesOfElements(element_ids, property_ids)
        
        for elem_props in property_values:
            if not elem_props.propertyValues:
                continue
                
            prop_value = elem_props.propertyValues[0]
            
            # Get actual value
            if hasattr(prop_value.propertyValue, 'value'):
                actual_value = prop_value.propertyValue.value
                
                # Apply comparison
                match = False
                if comparison == 'equals':
                    match = actual_value == value
                elif comparison == 'greater':
                    match = actual_value > value
                elif comparison == 'less':
                    match = actual_value < value
                elif comparison == 'contains':
                    match = str(value).lower() in str(actual_value).lower()
                
                if match:
                    filtered.append(elem_props.elementId.guid)
                    
    except Exception as e:
        print(f"Error filtering by property: {e}")
    
    return filtered


def filter_by_classification(element_guids, classification_id):
    """
    Filter elements by classification.
    
    Args:
        element_guids: List of element GUIDs to filter
        classification_id: Classification item ID
        
    Returns:
        List of filtered element GUIDs
    """
    filtered = []
    
    element_ids = [act.ElementId(guid) for guid in element_guids]
    
    try:
        # Get classifications for all elements
        classifications = acc.GetClassificationsOfElements(element_ids)
        
        for elem_class in classifications:
            if elem_class.classificationIds:
                for class_id in elem_class.classificationIds:
                    if class_id.classificationItemId == classification_id:
                        filtered.append(elem_class.elementId.guid)
                        break
                        
    except Exception as e:
        print(f"Error filtering by classification: {e}")
    
    return filtered


def filter_by_layer(element_guids, layer_name):
    """
    Filter elements by layer name.
    
    Args:
        element_guids: List of element GUIDs to filter
        layer_name: Name of the layer
        
    Returns:
        List of filtered element GUIDs
    """
    filtered = []
    
    try:
        # Get all layers
        all_layers = acc.GetAllLayers()
        
        # Find layer by name
        target_layer = None
        for layer in all_layers:
            if layer.name == layer_name:
                target_layer = layer
                break
        
        if not target_layer:
            print(f"Layer '{layer_name}' not found.")
            return filtered
        
        # Check each element
        element_ids = [act.ElementId(guid) for guid in element_guids]
        
        for element_id in element_ids:
            # Get element details
            details = acc.GetDetailsOfElements([element_id])
            
            if details and hasattr(details[0], 'layer'):
                if details[0].layer == target_layer.name:
                    filtered.append(element_id.guid)
                    
    except Exception as e:
        print(f"Error filtering by layer: {e}")
    
    return filtered


def apply_filtered_selection(element_guids):
    """
    Apply filtered selection in Archicad.
    
    Args:
        element_guids: List of element GUIDs to select
    """
    element_ids = [act.ElementId(guid) for guid in element_guids]
    
    try:
        acc.SetSelectedElements(element_ids)
        print(f"Selection updated: {len(element_guids)} element(s) selected.")
    except Exception as e:
        print(f"Error applying selection: {e}")


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("FILTER SELECTION BY CRITERIA")
    print("="*80)
    
    # Get current selection
    selected = acc.GetSelectedElements()
    
    if not selected:
        print("\nNo elements selected. Please select elements first.")
        return
    
    selected_guids = [elem.elementId.guid for elem in selected]
    print(f"\nInitial selection: {len(selected_guids)} element(s)")
    
    # Example 1: Filter by type
    print("\n" + "─"*80)
    print("Example 1: Filter by Element Type (Walls only)")
    print("─"*80)
    
    walls_only = filter_by_type(selected_guids, ['Wall'])
    print(f"Walls found: {len(walls_only)}")
    
    # Example 2: Filter by layer
    print("\n" + "─"*80)
    print("Example 2: Filter by Layer")
    print("─"*80)
    
    # This is an example - replace with actual layer name from your project
    layer_filtered = filter_by_layer(selected_guids, 'A-WALL')
    print(f"Elements on 'A-WALL' layer: {len(layer_filtered)}")
    
    # Apply the filtered selection (uncomment to activate)
    # apply_filtered_selection(walls_only)
    
    print("\n" + "="*80)
    print("Note: Uncomment 'apply_filtered_selection' in the code to apply filters.")
    print("="*80)


if __name__ == "__main__":
    main()
