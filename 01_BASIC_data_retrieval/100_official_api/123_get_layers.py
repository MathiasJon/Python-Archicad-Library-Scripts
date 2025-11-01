"""
================================================================================
SCRIPT: Get Layers (Complete)
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Layer attributes in the current Archicad project
with complete detailed information including visibility, lock status, 
intersection groups, and wireframe mode.

Layers are organizational tools that control visibility and editability of
elements. This script provides comprehensive layer information:
- List of all layers with complete status
- Intersection group numbers
- Wireframe mode status
- Visibility and lock states
- Folder hierarchy organization
- GUID and index mapping
- Optional export to text file

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
  
- GetAttributesByType('Layer')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetLayerAttributes(attributeIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLayerAttributes
  Returns detailed layer attributes identified by their GUIDs including:
  * attributeId: The identifier of the attribute
  * name: The name of the layer
  * intersectionGroupNr: The intersection group number (0-99)
  * isLocked: Whether the layer is locked (not editable)
  * isHidden: Whether the layer is hidden (not visible)
  * isWireframe: Whether elements display as wireframes or solid

- GetAttributeFolderStructure('Layer')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

- GetAttributesIndices(attributeIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesIndices
  Available in Archicad 28+

[Data Types]
- LayerAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.LayerAttribute
  A layer attribute with:
  * attributeId: The identifier of an attribute
  * name: The name of an attribute
  * intersectionGroupNr: The intersection group number
  * isLocked: Defines whether the layer is locked or not
  * isHidden: Defines whether the layer is hidden or not
  * isWireframe: Defines whether elements are visible as wireframes or solid

- LayerAttributeOrError
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.LayerAttributeOrError

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View comprehensive layers list with all properties
4. View folder structure hierarchy

Optional:
5. Call export_to_file(layers, 'filename.txt') to save detailed results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of layers
- For each layer:
  * Name and GUID
  * Status flags (Hidden/Locked/Wireframe)
  * Intersection group number
- Folder hierarchy structure
- GUID/Index mapping (Archicad 28+)
- Summary statistics

Example output for a layer:
  1. Architectural Walls [Locked]
     GUID: ABC123...
     Intersection Group: 1
     Display Mode: Solid

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
- Layers control visibility and editability of elements
- Hidden layers: elements are not displayed in any view
- Locked layers: elements cannot be selected or edited
- Wireframe mode: elements display without materials/textures
- Intersection groups (0-99): control building element intersections
  * Elements on same intersection group intersect automatically
  * Group 0 means no automatic intersections
- Layers can be combined in Layer Combinations for different views
- Layers are organized in folders for better project management
- GUIDs are permanent identifiers, indices are project-specific

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 114_10_get_layer_combinations.py (predefined layer visibility sets)
- 201_set_element_layer.py (move elements to different layers)
- 202_toggle_layer_visibility.py (show/hide layers)
- 203_create_layer.py (create new layers)

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
    Get all layer attributes with complete details.
    
    Returns:
        Tuple of (layers list, attribute IDs list)
    """
    try:
        print("\n" + "="*80)
        print("LAYERS - DETAILED INFORMATION")
        print("="*80)
        
        # Get all layer attribute IDs
        attribute_ids = acc.GetAttributesByType('Layer')
        print(f"\n✓ Found {len(attribute_ids)} layer(s)")

        if not attribute_ids:
            print("\n⚠️  No layers found in the project")
            return [], []

        # Get detailed layer attributes
        layers = acc.GetLayerAttributes(attribute_ids)

        print(f"\nLayers List:")
        print("-"*80)

        for idx, layer_or_error in enumerate(layers, 1):
            if hasattr(layer_or_error, 'layerAttribute') and layer_or_error.layerAttribute:
                layer = layer_or_error.layerAttribute
                
                # Basic information
                name = layer.name if hasattr(layer, 'name') else 'Unknown'
                guid = layer.attributeId.guid if hasattr(layer, 'attributeId') and layer.attributeId else 'N/A'
                
                # Build status flags
                status = []
                if hasattr(layer, 'isHidden') and layer.isHidden:
                    status.append("Hidden")
                if hasattr(layer, 'isLocked') and layer.isLocked:
                    status.append("Locked")
                if hasattr(layer, 'isWireframe') and layer.isWireframe:
                    status.append("Wireframe")
                
                status_str = f" [{', '.join(status)}]" if status else ""
                
                print(f"\n{idx}. {name}{status_str}")
                print(f"   GUID: {guid}")
                
                # Intersection group
                if hasattr(layer, 'intersectionGroupNr'):
                    group = layer.intersectionGroupNr
                    if group == 0:
                        print(f"   Intersection Group: {group} (No automatic intersections)")
                    else:
                        print(f"   Intersection Group: {group}")
                
                # Display mode
                if hasattr(layer, 'isWireframe'):
                    display_mode = "Wireframe" if layer.isWireframe else "Solid"
                    print(f"   Display Mode: {display_mode}")
            
            elif hasattr(layer_or_error, 'error'):
                print(f"\n{idx}. Error: {layer_or_error.error}")

        return layers, attribute_ids

    except Exception as e:
        print(f"\n✗ Error getting layers: {e}")
        import traceback
        traceback.print_exc()
        return [], []


def get_folder_structure(attribute_type='Layer', path=None, indent=0):
    """
    Display the folder structure for layers.
    
    Args:
        attribute_type: Type of attribute (default: 'Layer')
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
        
        print(f"{prefix}📁 {folder_name} ({num_attributes} layer(s), {num_subfolders} subfolder(s))")
        
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


def get_indices_mapping(attribute_ids):
    """
    Get GUID to Index mapping.
    
    Args:
        attribute_ids: List of attribute identifiers
    
    Returns:
        List of index mappings
    
    Note:
        GetAttributesIndices is available in Archicad 28+
    """
    try:
        if not attribute_ids:
            return []
        
        # Check if GetAttributesIndices is available
        if not hasattr(acc, 'GetAttributesIndices'):
            print("\n" + "="*80)
            print("GUID / INDEX MAPPING")
            print("="*80)
            print("\n⚠️  GetAttributesIndices API not available in this Archicad version")
            print("    This feature requires Archicad 28 or later")
            return []
        
        print("\n" + "="*80)
        print("GUID / INDEX MAPPING")
        print("="*80)
        
        indices = acc.GetAttributesIndices(attribute_ids)
        
        print(f"\nTotal mappings: {len(indices)}")
        print("-"*80)
        
        for idx, result in enumerate(indices, 1):
            if hasattr(result, 'attributeIndexAndGuid') and result.attributeIndexAndGuid:
                index_info = result.attributeIndexAndGuid
                print(f"  {idx:3}. Index: {index_info.index:4} | GUID: {index_info.guid}")
        
        return indices
    
    except Exception as e:
        print(f"\n✗ Error getting indices: {e}")
        return []


def export_to_file(layers, filename='layers_detailed.txt'):
    """
    Export layers to a detailed text file.
    
    Args:
        layers: List of LayerAttributeOrError objects
        filename: Output filename (default: 'layers_detailed.txt')
    """
    try:
        import os
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD LAYERS - DETAILED INFORMATION\n")
            f.write("="*80 + "\n\n")
            
            for idx, layer_or_error in enumerate(layers, 1):
                if hasattr(layer_or_error, 'layerAttribute') and layer_or_error.layerAttribute:
                    layer = layer_or_error.layerAttribute
                    
                    # Basic information
                    name = layer.name if hasattr(layer, 'name') else 'Unknown'
                    
                    # Build status
                    status = []
                    if hasattr(layer, 'isHidden') and layer.isHidden:
                        status.append("Hidden")
                    if hasattr(layer, 'isLocked') and layer.isLocked:
                        status.append("Locked")
                    if hasattr(layer, 'isWireframe') and layer.isWireframe:
                        status.append("Wireframe")
                    
                    status_str = f" [{', '.join(status)}]" if status else ""
                    
                    f.write(f"{idx}. {name}{status_str}\n")
                    
                    if hasattr(layer, 'attributeId') and layer.attributeId:
                        f.write(f"   GUID: {layer.attributeId.guid}\n")
                    
                    # Intersection group
                    if hasattr(layer, 'intersectionGroupNr'):
                        group = layer.intersectionGroupNr
                        if group == 0:
                            f.write(f"   Intersection Group: {group} (No automatic intersections)\n")
                        else:
                            f.write(f"   Intersection Group: {group}\n")
                    
                    # Display mode
                    if hasattr(layer, 'isWireframe'):
                        display_mode = "Wireframe" if layer.isWireframe else "Solid"
                        f.write(f"   Display Mode: {display_mode}\n")
                    
                    f.write("\n")
        
        abs_path = os.path.abspath(filename)
        print(f"\n✓ Layers exported to:")
        print(f"  {abs_path}")
        
    except Exception as e:
        print(f"\n✗ Error exporting: {e}")


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("GET LAYERS v1.0")
    print("="*80)
    
    # Get all layers with details
    layers, attribute_ids = get_all_layers()
    
    # Display folder structure
    if layers:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('Layer')
    
    # Display GUID/Index mapping
    if attribute_ids:
        get_indices_mapping(attribute_ids)
    
    # Display summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Layers: {len(layers)}")
    
    # Count status types
    hidden_count = 0
    locked_count = 0
    wireframe_count = 0
    intersection_groups = set()
    
    for layer_or_error in layers:
        if hasattr(layer_or_error, 'layerAttribute') and layer_or_error.layerAttribute:
            layer = layer_or_error.layerAttribute
            if hasattr(layer, 'isHidden') and layer.isHidden:
                hidden_count += 1
            if hasattr(layer, 'isLocked') and layer.isLocked:
                locked_count += 1
            if hasattr(layer, 'isWireframe') and layer.isWireframe:
                wireframe_count += 1
            if hasattr(layer, 'intersectionGroupNr'):
                intersection_groups.add(layer.intersectionGroupNr)
    
    print(f"  Hidden Layers: {hidden_count}")
    print(f"  Locked Layers: {locked_count}")
    print(f"  Wireframe Layers: {wireframe_count}")
    print(f"  Intersection Groups Used: {len(intersection_groups)}")
    if intersection_groups:
        groups_list = sorted(intersection_groups)
        print(f"  Groups: {', '.join(map(str, groups_list))}")
    print("\n" + "="*80)
    
    # Usage hints
    print("\n💡 To export detailed information to file:")
    print("   layers, ids = get_all_layers()")
    print("   export_to_file(layers, 'my_layers_detailed.txt')")
    
    print("\n💡 Intersection Groups:")
    print("   - Group 0: No automatic intersections")
    print("   - Groups 1-99: Elements in same group intersect automatically")


if __name__ == "__main__":
    main()