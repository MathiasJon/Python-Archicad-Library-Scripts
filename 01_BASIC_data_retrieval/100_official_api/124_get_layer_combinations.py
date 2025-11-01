"""
================================================================================
SCRIPT: Get Layer Combinations (Complete)
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Layer Combination attributes in the current Archicad
project with complete detailed information including the list of layers included
in each combination.

Layer Combinations are predefined layer visibility sets for different views
and workflows. This script provides comprehensive information:
- List of all layer combinations
- Number of layers in each combination
- Detailed list of layers included in each combination
- Folder hierarchy organization
- Optional export to text file with full details

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetAttributesByType('LayerCombination')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetLayerCombinationAttributes(attributeIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLayerCombinationAttributes
  Returns detailed layer combination attributes identified by their GUIDs including:
  * attributeId: The identifier of the attribute
  * name: The name of the layer combination
  * layerAttributeIds: List of layer attribute identifiers included in this combination

- GetAttributeFolderStructure('LayerCombination')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

- GetLayerAttributes(attributeIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLayerAttributes
  Used to retrieve layer names from their IDs

[Data Types]
- LayerCombinationAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.LayerCombinationAttribute
  A layer combination attribute with:
  * attributeId: The identifier of the attribute
  * name: The name of the layer combination
  * layerAttributeIds: List of attribute identifiers for included layers

- LayerCombinationAttributeOrError
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.LayerCombinationAttributeOrError

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View comprehensive layer combinations list with all included layers
4. View folder structure hierarchy

Optional:
5. Call export_to_file(combinations, 'filename.txt') to save detailed results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of layer combinations
- For each combination:
  * Name and GUID
  * Number of layers included
  * Detailed list of all layers in the combination
- Folder hierarchy structure
- Summary statistics

Example output for a combination:
  1. Construction Views (45 layers)
     GUID: ABC123...
     Included Layers:
        • Structure - Walls
        • Structure - Slabs
        • Structure - Columns
        ... (42 more)

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the
official Python palette or Tapir palette), the script will automatically
connect to the FIRST instance of Archicad that was launched on your computer.

If you have multiple Archicad instances open:
- The script connects to the first one opened
- Make sure the correct project is open in that instance
- Close other instances if you need to target a specific project

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Layer combinations are preset layer visibility configurations
- Used to quickly switch between different views (Construction, Presentation, etc.)
- Each combination includes a specific set of layers from the project
- The layerAttributeIds list contains the IDs of all layers included
- Combinations can have different purposes: construction, MEP, presentation, etc.
- Organized in folders for better management
- Essential for multi-discipline coordination and efficient workflow
- Note: The API provides layer IDs but not their visibility status in the combination

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 123_get_layers.py (list all layers with details)
- 202_set_layer_combination.py (activate a layer combination)
- 203_create_layer_combination.py (create new combinations)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


def get_all_layers():
    """
    Get all layers from the project to map IDs to names.
    
    Returns:
        Dictionary mapping layer GUIDs to layer names
    """
    try:
        layer_ids = acc.GetAttributesByType('Layer')
        if not layer_ids:
            return {}
        
        layers = acc.GetLayerAttributes(layer_ids)
        
        # Create GUID to name mapping
        layer_map = {}
        for layer_or_error in layers:
            if hasattr(layer_or_error, 'layerAttribute') and layer_or_error.layerAttribute:
                layer = layer_or_error.layerAttribute
                if hasattr(layer, 'attributeId') and hasattr(layer, 'name'):
                    guid = layer.attributeId.guid
                    # Convert UUID to string for dictionary key
                    guid_str = str(guid) if not isinstance(guid, str) else guid
                    name = layer.name
                    layer_map[guid_str] = name
        
        return layer_map
    except Exception as e:
        print(f"⚠️  Error getting layers: {e}")
        return {}


def get_all_layer_combinations(layer_map):
    """
    Get all layer combination attributes with complete details.
    
    Args:
        layer_map: Dictionary mapping layer GUIDs to names
    
    Returns:
        Tuple of (combinations list, attribute IDs list)
    """
    try:
        print("\n" + "="*80)
        print("LAYER COMBINATIONS - DETAILED INFORMATION")
        print("="*80)
        
        # Get all layer combination attribute IDs
        attribute_ids = acc.GetAttributesByType('LayerCombination')
        print(f"\n✓ Found {len(attribute_ids)} layer combination(s)")

        if not attribute_ids:
            print("\n⚠️  No layer combinations found in the project")
            return [], []

        # Get detailed layer combination attributes
        layer_combinations = acc.GetLayerCombinationAttributes(attribute_ids)

        print(f"\nLayer Combinations List:")
        print("-"*80)

        for idx, comb_or_error in enumerate(layer_combinations, 1):
            if hasattr(comb_or_error, 'layerCombinationAttribute') and comb_or_error.layerCombinationAttribute:
                comb = comb_or_error.layerCombinationAttribute
                
                # Basic information
                name = comb.name if hasattr(comb, 'name') else 'Unknown'
                guid = comb.attributeId.guid if hasattr(comb, 'attributeId') and comb.attributeId else 'N/A'
                
                # Get layer IDs
                layer_ids = []
                if hasattr(comb, 'layerAttributeIds') and comb.layerAttributeIds:
                    layer_ids = comb.layerAttributeIds
                
                num_layers = len(layer_ids)
                
                print(f"\n{idx}. {name} ({num_layers} layer(s))")
                print(f"   GUID: {guid}")
                
                # Display included layers
                if layer_ids and layer_map:
                    print(f"   Included Layers:")
                    
                    # Get layer names
                    layer_names = []
                    for layer_id_wrapper in layer_ids:
                        if hasattr(layer_id_wrapper, 'attributeId') and layer_id_wrapper.attributeId:
                            layer_guid = layer_id_wrapper.attributeId.guid
                            # Convert UUID to string for dictionary lookup
                            layer_guid_str = str(layer_guid) if not isinstance(layer_guid, str) else layer_guid
                            layer_name = layer_map.get(layer_guid_str, f"Unknown ({layer_guid_str[:8]}...)")
                            layer_names.append(layer_name)
                    
                    # Display first 10 layers
                    for i, layer_name in enumerate(layer_names[:10]):
                        print(f"      • {layer_name}")
                    
                    # Show count of remaining layers
                    if len(layer_names) > 10:
                        print(f"      ... and {len(layer_names) - 10} more layers")
            
            elif hasattr(comb_or_error, 'error'):
                print(f"\n{idx}. Error: {comb_or_error.error}")

        return layer_combinations, attribute_ids

    except Exception as e:
        print(f"\n✗ Error getting layer combinations: {e}")
        import traceback
        traceback.print_exc()
        return [], []


def get_folder_structure(attribute_type='LayerCombination', path=None, indent=0):
    """
    Display the folder structure for layer combinations.
    
    Args:
        attribute_type: Type of attribute (default: 'LayerCombination')
        path: Current folder path
        indent: Indentation level for display
    
    Returns:
        Folder structure object
    """
    try:
        structure = acc.GetAttributeFolderStructure(attribute_type, path)
        
        prefix = "  " * indent
        folder_name = structure.name if hasattr(structure, 'name') else 'Root'
        
        # Count attributes
        num_attributes = 0
        if hasattr(structure, 'attributes') and structure.attributes is not None:
            num_attributes = len(structure.attributes)
        
        # Count subfolders
        num_subfolders = 0
        if hasattr(structure, 'subfolders') and structure.subfolders is not None:
            num_subfolders = len(structure.subfolders)
        
        print(f"{prefix}📁 {folder_name} ({num_attributes} combination(s), {num_subfolders} subfolder(s))")
        
        # Display subfolders recursively
        if hasattr(structure, 'subfolders') and structure.subfolders is not None and len(structure.subfolders) > 0:
            for subfolder_item in structure.subfolders:
                if hasattr(subfolder_item, 'attributeFolder'):
                    folder_struct = subfolder_item.attributeFolder
                    if hasattr(folder_struct, 'name'):
                        subfolder_path = path.copy() if path else []
                        subfolder_path.append(folder_struct.name)
                        get_folder_structure(attribute_type, subfolder_path, indent + 1)
        
        return structure
    
    except Exception as e:
        print(f"{prefix}✗ Error: {e}")
        return None


def export_to_file(layer_combinations, layer_map, filename='layer_combinations_detailed.txt'):
    """
    Export layer combinations to a detailed text file.
    
    Args:
        layer_combinations: List of LayerCombinationAttributeOrError objects
        layer_map: Dictionary mapping layer GUIDs to names
        filename: Output filename (default: 'layer_combinations_detailed.txt')
    """
    try:
        import os
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD LAYER COMBINATIONS - DETAILED INFORMATION\n")
            f.write("="*80 + "\n\n")
            
            for idx, comb_or_error in enumerate(layer_combinations, 1):
                if hasattr(comb_or_error, 'layerCombinationAttribute') and comb_or_error.layerCombinationAttribute:
                    comb = comb_or_error.layerCombinationAttribute
                    
                    # Basic information
                    name = comb.name if hasattr(comb, 'name') else 'Unknown'
                    
                    # Get layer IDs
                    layer_ids = []
                    if hasattr(comb, 'layerAttributeIds') and comb.layerAttributeIds:
                        layer_ids = comb.layerAttributeIds
                    
                    num_layers = len(layer_ids)
                    
                    f.write(f"{idx}. {name} ({num_layers} layer(s))\n")
                    
                    if hasattr(comb, 'attributeId') and comb.attributeId:
                        f.write(f"   GUID: {comb.attributeId.guid}\n")
                    
                    # Write included layers
                    if layer_ids and layer_map:
                        f.write(f"   Included Layers:\n")
                        
                        # Get all layer names
                        for layer_id_wrapper in layer_ids:
                            if hasattr(layer_id_wrapper, 'attributeId') and layer_id_wrapper.attributeId:
                                layer_guid = layer_id_wrapper.attributeId.guid
                                # Convert UUID to string for dictionary lookup
                                layer_guid_str = str(layer_guid) if not isinstance(layer_guid, str) else layer_guid
                                layer_name = layer_map.get(layer_guid_str, f"Unknown ({layer_guid_str[:8]}...)")
                                f.write(f"      • {layer_name}\n")
                    
                    f.write("\n")
        
        abs_path = os.path.abspath(filename)
        print(f"\n✓ Layer combinations exported to:")
        print(f"  {abs_path}")
        
    except Exception as e:
        print(f"\n✗ Error exporting: {e}")


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("GET LAYER COMBINATIONS v1.0")
    print("="*80)
    
    # First, get all layers to map IDs to names
    print("\nLoading layer information...")
    layer_map = get_all_layers()
    print(f"✓ Loaded {len(layer_map)} layer(s)")
    
    # Get all layer combinations with details
    layer_combinations, attribute_ids = get_all_layer_combinations(layer_map)
    
    # Display folder structure
    if layer_combinations:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('LayerCombination')
    
    # Display summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Layer Combinations: {len(layer_combinations)}")
    
    # Calculate statistics
    if layer_combinations:
        total_layers = 0
        max_layers = 0
        min_layers = float('inf')
        
        for comb_or_error in layer_combinations:
            if hasattr(comb_or_error, 'layerCombinationAttribute') and comb_or_error.layerCombinationAttribute:
                comb = comb_or_error.layerCombinationAttribute
                if hasattr(comb, 'layerAttributeIds') and comb.layerAttributeIds:
                    num = len(comb.layerAttributeIds)
                    total_layers += num
                    max_layers = max(max_layers, num)
                    min_layers = min(min_layers, num)
        
        if len(layer_combinations) > 0:
            avg_layers = total_layers / len(layer_combinations)
            print(f"  Average Layers per Combination: {avg_layers:.1f}")
            print(f"  Max Layers in a Combination: {max_layers}")
            print(f"  Min Layers in a Combination: {min_layers if min_layers != float('inf') else 0}")
    
    print("\n" + "="*80)
    
    # Usage hints
    print("\n💡 To export detailed information to file:")
    print("   layer_map = get_all_layers()")
    print("   combinations, ids = get_all_layer_combinations(layer_map)")
    print("   export_to_file(combinations, layer_map, 'my_combinations_detailed.txt')")
    
    print("\n💡 Layer combinations control which layers are visible")
    print("   in different views and workflows")


if __name__ == "__main__":
    main()