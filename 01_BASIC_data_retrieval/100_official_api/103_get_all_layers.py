"""
================================================================================
SCRIPT: Get All Layers with Folder Structure
================================================================================

Version:        1.0
Date:           2025-11-01
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves all layer attributes from the project organized by their folder 
structure and displays their properties. Shows layer name, GUID, lock status, 
hidden status, intersection groups, and folder hierarchy.

Also provides a filtered list of active layers (not locked or hidden).

This script uses a three-step process:
1. Get folder structure using GetAttributeFolderStructure()
2. Recursively traverse folders to collect all layers
3. Get detailed layer information using GetLayerAttributes()

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAttributeFolderStructure()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure
  
- GetLayerAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLayerAttributes

[Data Types]
- AttributeFolderStructure
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac27.html#archicad.releases.ac27.b3001types.AttributeFolderStructure
  
- AttributeHeader
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AttributeHeader
  
- AttributeId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AttributeId
  
- LayerAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.LayerAttribute

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View layer list organized by folders in console output

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. LAYER FOLDER STRUCTURE:
   - Hierarchical display of folders and layers
   - Indentation shows folder depth
   - Layer properties (GUID, lock, hidden status)

2. ACTIVE LAYERS LIST:
   - Filtered list of layers that are not locked or hidden
   - Sorted alphabetically

3. SUMMARY STATISTICS:
   - Total layers count
   - Active layers count
   - Locked/Hidden layers count
   - Number of folders

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
- If no layers are found, the script exits with a warning message
- Some layer properties may not be available depending on API version
- Layer GUIDs are obtained from AttributeHeader
- Active layers are defined as not locked AND not hidden
- Folders are displayed hierarchically with proper indentation

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 106_get_elements_by_type.py
- 459_create_layer.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types


def collect_layers_from_folder(folder_structure, folder_path="", indent_level=0):
    """
    Recursively collect all layers from folder structure

    Args:
        folder_structure: AttributeFolderStructure object
        folder_path: Current folder path for display
        indent_level: Current indentation level for display

    Returns:
        tuple: (list of layer headers, list of display info)
    """
    layer_headers = []
    display_info = []

    # Prepare indentation
    indent = "  " * indent_level

    # Display folder name
    if folder_path:
        display_info.append((indent_level, f"📁 {folder_structure.name}"))
    else:
        display_info.append((indent_level, f"📁 Root"))

    # Collect layers from current folder
    if folder_structure.attributes:
        for attr_item in folder_structure.attributes:
            # AttributeHeaderArrayItem has an 'attribute' property containing the AttributeHeader
            header = attr_item.attribute
            layer_headers.append(header)
            # Store layer info for later display with details
            display_info.append(
                (indent_level + 1, header, folder_structure.name))

    # Recursively process subfolders
    if folder_structure.subfolders:
        for subfolder_item in folder_structure.subfolders:
            # AttributeFolderStructureArrayItem has an 'attributeFolder' property
            subfolder = subfolder_item.attributeFolder
            sub_path = f"{folder_path}/{subfolder.name}" if folder_path else subfolder.name

            sub_headers, sub_display = collect_layers_from_folder(
                subfolder,
                sub_path,
                indent_level + 1
            )

            layer_headers.extend(sub_headers)
            display_info.extend(sub_display)

    return layer_headers, display_info


# ============================================================================
# STEP 1: Get folder structure for layers
# ============================================================================
# Get the complete folder structure starting from root
folder_structure = acc.GetAttributeFolderStructure('Layer')

print(f"\n=== PROJECT LAYERS - FOLDER STRUCTURE ===")

# ============================================================================
# STEP 2: Collect all layers from folder structure
# ============================================================================
layer_headers, display_info = collect_layers_from_folder(folder_structure)

print(f"Total layers: {len(layer_headers)}")
print(f"Organized in folder structure\n")

if len(layer_headers) == 0:
    print("\n⚠️  No layers found in the project")
    exit()

# ============================================================================
# STEP 3: Get detailed layer information
# ============================================================================
# Extract just the attribute IDs from the headers
layer_ids = [header.attributeId for header in layer_headers]

# Get detailed layer attributes
layers = acc.GetLayerAttributes(layer_ids)

# Create a mapping of GUID to layer details for easy lookup
layer_map = {}
for layer, header in zip(layers, layer_headers):
    layer_map[header.attributeId.guid] = (layer, header)

# ============================================================================
# DISPLAY LAYER FOLDER STRUCTURE WITH DETAILS
# ============================================================================
print("\nLayer folder structure:\n")

folder_count = 0
for item in display_info:
    if len(item) == 2:
        # This is a folder
        indent_level, folder_name = item
        indent = "  " * indent_level
        print(f"{indent}{folder_name}")
        folder_count += 1
    else:
        # This is a layer
        indent_level, header, folder_name = item
        indent = "  " * indent_level

        layer, layer_header = layer_map[header.attributeId.guid]

        print(f"{indent}└─ {layer.layerAttribute.name}")
        print(f"{indent}   GUID: {header.attributeId.guid}")

        # Check if layer is locked or hidden
        if hasattr(layer.layerAttribute, 'isLocked'):
            status = "🔒 Locked" if layer.layerAttribute.isLocked else "🔓 Unlocked"
            print(f"{indent}   {status}")

        if hasattr(layer.layerAttribute, 'isHidden'):
            visibility = "👁️ Hidden" if layer.layerAttribute.isHidden else "👁️ Visible"
            print(f"{indent}   {visibility}")

        # Additional properties if available
        if hasattr(layer.layerAttribute, 'intersectionGroupNr'):
            print(
                f"{indent}   Intersection Group: {layer.layerAttribute.intersectionGroupNr}")

        print()  # Empty line for readability

# ============================================================================
# DISPLAY ACTIVE LAYERS (not locked or hidden)
# ============================================================================
print("\n" + "="*70)
print("=== ACTIVE LAYERS ===")
print("(Layers that are not locked or hidden)")
print("="*70 + "\n")

active_layers = []
for layer in layers:
    # Get lock and hidden status with safe defaults
    is_locked = getattr(layer.layerAttribute, 'isLocked', False)
    is_hidden = getattr(layer.layerAttribute, 'isHidden', False)

    # Add to active list if not locked and not hidden
    if not is_locked and not is_hidden:
        active_layers.append(layer.layerAttribute.name)

if active_layers:
    for name in sorted(active_layers):
        print(f"  ✓ {name}")
    print(f"\nTotal active layers: {len(active_layers)}")
else:
    print("  (All layers are locked or hidden)")

# ============================================================================
# SUMMARY
# ============================================================================
print("\n" + "="*70)
print("SUMMARY")
print("="*70)
print(f"Total layers in project: {len(layers)}")
print(f"Total folders: {folder_count}")
print(f"Active layers: {len(active_layers)}")
print(f"Locked/Hidden layers: {len(layers) - len(active_layers)}")

print("\n💡 TIP: Use these layer names in scripts to:")
print("   - Filter elements by layer")
print("   - Modify layer visibility")
print("   - Create new layers")
print("   - Organize layers in folders")
