"""
Complete Layer Management Script - 100% TAPIR Commands

This script uses ONLY TAPIR commands for all layer operations:
- GetAttributesByType (v1.1.3): Get all layers with names and indices
- CreateLayers (v1.0.3): Create new layers
- GetPropertyValuesOfAttributes (v1.1.8): Get layer properties
- SetPropertyValuesOfAttributes (v1.1.8): Set layer properties

Based on TAPIR API documentation.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


# =============================================================================
# GET LAYERS
# =============================================================================

def get_all_layers():
    """
    Get all layers in the project.
    Uses TAPIR GetAttributesByType command.

    Returns:
        List of layer dictionaries with attributeId, index, and name
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'GetAttributesByType')

        parameters = {
            'attributeType': 'Layer'
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ GetAttributesByType failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'attributes' in response:
            return response['attributes']

        return []

    except Exception as e:
        print(f"✗ Error getting layers: {e}")
        return []


def list_all_layers():
    """
    List all layers with their basic info.
    """
    layers = get_all_layers()

    if not layers:
        print("⚠ No layers found in project")
        return

    print(f"\nTotal layers: {len(layers)}")
    print("="*80)

    # Sort by index
    sorted_layers = sorted(layers, key=lambda x: x.get('index', 0))

    for layer in sorted_layers:
        index = layer.get('index', '?')
        name = layer.get('name', 'Unknown')
        guid = layer.get('attributeId', {}).get('guid', 'N/A')

        print(f"  [{index:3}] {name}")
        print(f"        GUID: {guid}")


def find_layer_by_name(layer_name):
    """
    Find a layer by its name.

    Args:
        layer_name: Name of the layer to find

    Returns:
        Layer dictionary or None
    """
    layers = get_all_layers()

    for layer in layers:
        if layer.get('name', '').lower() == layer_name.lower():
            return layer

    return None


def find_layer_by_index(layer_index):
    """
    Find a layer by its index.

    Args:
        layer_index: Index of the layer

    Returns:
        Layer dictionary or None
    """
    layers = get_all_layers()

    for layer in layers:
        if layer.get('index') == layer_index:
            return layer

    return None


# =============================================================================
# CREATE LAYERS
# =============================================================================

def create_layers(layer_data_array, overwrite_existing=False):
    """
    Create one or more layers.
    Uses TAPIR CreateLayers command.

    Args:
        layer_data_array: List of layer data dictionaries
                         Each dictionary should contain:
                         - name (string, required): Layer name
                         - isHidden (boolean, optional): Hide/Show
                         - isLocked (boolean, optional): Lock/Unlock
                         - isWireframe (boolean, optional): Force wireframe
        overwrite_existing: Overwrite if layer exists (default: False)

    Returns:
        List of created layer attribute IDs (GUIDs)
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'CreateLayers')

        parameters = {
            'layerDataArray': layer_data_array,
            'overwriteExisting': overwrite_existing
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ CreateLayers failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'attributeIds' in response:
            attr_ids = response['attributeIds']
            guids = []
            for attr in attr_ids:
                if isinstance(attr, dict) and 'attributeId' in attr:
                    guid = attr['attributeId'].get('guid')
                    if guid:
                        guids.append(guid)

            if guids:
                print(f"✓ Created {len(guids)} layer(s)")
            return guids

        return []

    except Exception as e:
        print(f"✗ Error creating layers: {e}")
        return []


def create_layer(name, hidden=False, locked=False, wireframe=False, overwrite=False):
    """
    Create a single layer.

    Args:
        name: Layer name
        hidden: Hide layer (default: False)
        locked: Lock layer (default: False)
        wireframe: Force wireframe display (default: False)
        overwrite: Overwrite if exists (default: False)

    Returns:
        Layer GUID or None
    """
    layer_data = {
        'name': name,
        'isHidden': hidden,
        'isLocked': locked,
        'isWireframe': wireframe
    }

    guids = create_layers([layer_data], overwrite_existing=overwrite)

    if guids:
        status = []
        if hidden:
            status.append("Hidden")
        if locked:
            status.append("Locked")
        if wireframe:
            status.append("Wireframe")

        status_str = f" [{', '.join(status)}]" if status else ""
        print(f"  ✓ Layer '{name}'{status_str}")

        return guids[0]

    return None


def create_multiple_layers_batch(layer_configs, overwrite=False):
    """
    Create multiple layers in one batch operation.

    Args:
        layer_configs: List of layer configuration dictionaries
        overwrite: Overwrite existing layers (default: False)

    Returns:
        List of created layer GUIDs
    """
    layer_data_array = []

    for config in layer_configs:
        layer_data = {
            'name': config.get('name', 'Unnamed Layer'),
            'isHidden': config.get('hidden', False) or config.get('isHidden', False),
            'isLocked': config.get('locked', False) or config.get('isLocked', False),
            'isWireframe': config.get('wireframe', False) or config.get('isWireframe', False)
        }
        layer_data_array.append(layer_data)

    return create_layers(layer_data_array, overwrite_existing=overwrite)


# =============================================================================
# LAYER PROPERTIES (CUSTOM PROPERTIES)
# =============================================================================

def get_layer_property_values(layer_guids, property_guids):
    """
    Get property values for layers.
    Uses TAPIR GetPropertyValuesOfAttributes command.

    Args:
        layer_guids: List of layer attribute GUIDs
        property_guids: List of property GUIDs to retrieve

    Returns:
        List of property value results
    """
    try:
        command_id = act.AddOnCommandId(
            'TapirCommand', 'GetPropertyValuesOfAttributes')

        # Prepare attribute IDs
        attribute_ids = [
            {'attributeId': {'guid': guid}} for guid in layer_guids
        ]

        # Prepare property IDs
        property_ids = [
            {'propertyId': {'guid': guid}} for guid in property_guids
        ]

        parameters = {
            'attributeIds': attribute_ids,
            'properties': property_ids
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ GetPropertyValuesOfAttributes failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'propertyValuesForAttributes' in response:
            return response['propertyValuesForAttributes']

        return []

    except Exception as e:
        print(f"✗ Error getting property values: {e}")
        return []


def set_layer_property_values(attribute_property_values):
    """
    Set property values for layers.
    Uses TAPIR SetPropertyValuesOfAttributes command.

    Args:
        attribute_property_values: List of dictionaries, each containing:
                                   - attributeId: {'guid': '...'}
                                   - propertyId: {'guid': '...'}
                                   - propertyValue: {'value': '...'}

    Returns:
        List of execution results
    """
    try:
        command_id = act.AddOnCommandId(
            'TapirCommand', 'SetPropertyValuesOfAttributes')

        parameters = {
            'attributePropertyValues': attribute_property_values
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ SetPropertyValuesOfAttributes failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'executionResults' in response:
            results = response['executionResults']

            success_count = sum(1 for r in results if isinstance(
                r, dict) and r.get('success', False))
            print(f"✓ Set {success_count}/{len(results)} property value(s)")

            return results

        return []

    except Exception as e:
        print(f"✗ Error setting property values: {e}")
        return []


def set_layer_property(layer_guid, property_guid, value):
    """
    Set a single property value for a layer.

    Args:
        layer_guid: Layer attribute GUID
        property_guid: Property GUID
        value: Property value (string)

    Returns:
        Boolean indicating success
    """
    property_value = {
        'attributeId': {'guid': layer_guid},
        'propertyId': {'guid': property_guid},
        'propertyValue': {'value': str(value)}
    }

    results = set_layer_property_values([property_value])

    if results and len(results) > 0:
        return results[0].get('success', False)

    return False


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def export_layer_list(filename='layers_list.txt'):
    """
    Export all layers to a text file.

    Args:
        filename: Output filename
    """
    try:
        import os

        layers = get_all_layers()

        if not layers:
            print("⚠ No layers to export")
            return

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD LAYERS LIST\n")
            f.write("="*80 + "\n\n")
            f.write(f"Total layers: {len(layers)}\n\n")

            # Sort by index
            sorted_layers = sorted(layers, key=lambda x: x.get('index', 0))

            for layer in sorted_layers:
                index = layer.get('index', '?')
                name = layer.get('name', 'Unknown')
                guid = layer.get('attributeId', {}).get('guid', 'N/A')

                f.write(f"[{index}] {name}\n")
                f.write(f"   GUID: {guid}\n")
                f.write("\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Layer list exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"✗ Error exporting layer list: {e}")


def export_layer_list_csv(filename='layers_list.csv'):
    """
    Export layers to CSV format.

    Args:
        filename: Output CSV filename
    """
    try:
        import os
        import csv

        layers = get_all_layers()

        if not layers:
            print("⚠ No layers to export")
            return

        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Index', 'Name', 'GUID'])

            # Sort by index
            sorted_layers = sorted(layers, key=lambda x: x.get('index', 0))

            for layer in sorted_layers:
                index = layer.get('index', '?')
                name = layer.get('name', 'Unknown')
                guid = layer.get('attributeId', {}).get('guid', 'N/A')

                writer.writerow([index, name, guid])

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Layers exported to CSV:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"✗ Error exporting to CSV: {e}")


# =============================================================================
# MAIN DEMONSTRATION
# =============================================================================

def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("LAYER MANAGEMENT - 100% TAPIR")
    print("="*80)

    # Example 1: List all existing layers
    print("\n" + "─"*80)
    print("EXAMPLE 1: List All Layers")
    print("─"*80)

    list_all_layers()

    # Example 2: Find a specific layer
    print("\n" + "─"*80)
    print("EXAMPLE 2: Find Layer by Name")
    print("─"*80 + "\n")

    found_layer = find_layer_by_name("ArchiCAD Layer")
    if found_layer:
        print(f"✓ Found layer: {found_layer.get('name')}")
        print(f"  Index: {found_layer.get('index')}")
        print(f"  GUID: {found_layer.get('attributeId', {}).get('guid')}")
    else:
        print("⚠ Layer not found")

    # Example 3: Create a single layer
    print("\n" + "─"*80)
    print("EXAMPLE 3: Create Single Layer")
    print("─"*80 + "\n")

    new_layer_guid = create_layer(
        name="TAPIR-TEST-LAYER",
        hidden=False,
        locked=False,
        wireframe=False,
        overwrite=True  # Will overwrite if exists
    )

    # Example 4: Create multiple layers in batch
    print("\n" + "─"*80)
    print("EXAMPLE 4: Create Multiple Layers (Batch)")
    print("─"*80 + "\n")

    layer_configs = [
        {
            'name': 'TAPIR-STRUCTURE',
            'hidden': False,
            'locked': False,
            'wireframe': False
        },
        {
            'name': 'TAPIR-FURNITURE',
            'hidden': False,
            'locked': False,
            'wireframe': True
        },
        {
            'name': 'TAPIR-ANNOTATION',
            'hidden': False,
            'locked': True,
            'wireframe': False
        },
        {
            'name': 'TAPIR-CONSTRUCTION',
            'hidden': True,
            'locked': False,
            'wireframe': False
        }
    ]

    created_guids = create_multiple_layers_batch(layer_configs, overwrite=True)
    print(f"\n✓ Batch created {len(created_guids)} layer(s)")

    # Example 5: List updated layers
    print("\n" + "─"*80)
    print("EXAMPLE 5: Updated Layer List")
    print("─"*80)

    list_all_layers()

    # Example 6: Export layers
    print("\n" + "─"*80)
    print("EXAMPLE 6: Export Layers")
    print("─"*80)

    export_layer_list('layers_list.txt')
    export_layer_list_csv('layers_list.csv')

    print("\n" + "="*80)
    print("✓ All operations completed using TAPIR commands only!")
    print("="*80)
    print("\n💡 TAPIR Commands used:")
    print("   • GetAttributesByType - Get all layers")
    print("   • CreateLayers - Create new layers")
    print("   • GetPropertyValuesOfAttributes - Get custom properties")
    print("   • SetPropertyValuesOfAttributes - Set custom properties")
    print("="*80)


if __name__ == "__main__":
    main()
