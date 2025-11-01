"""
================================================================================
SCRIPT: Get Building Materials with Cut Fill Details
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval


--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Building Material attributes in the current Archicad 
project with their folder structure organization and detailed Cut Fill properties.

Building Materials define physical material properties used in elements like 
walls, slabs, and roofs. This script provides:
- List of all building materials with detailed properties
- Folder hierarchy organization with materials listed under each folder
- GUID for each material
- Optional export to text file

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27/28)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAttributesByType('BuildingMaterial')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetBuildingMaterialAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetBuildingMaterialAttributes

- GetFillAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetFillAttributes

- GetPenTableAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetPenTableAttributes

- GetSurfaceAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSurfaceAttributes

- GetAttributeFolderStructure('BuildingMaterial')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

- GetAttributesIndices()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesIndices

[Data Types]
- BuildingMaterialAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.BuildingMaterialAttribute

- FillAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.FillAttribute

- SurfaceAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.SurfaceAttribute

- PenTableAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.PenTableAttribute

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View building materials list with Cut Fill details

Optional:
4. Call export_to_file('filename.txt') to save results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Folder structure with materials listed under each folder
- Total number of building materials
- For each material: Name, Material ID, GUID, Connection Priority, Cut Fill, 
  Cut Fill Pen, Cut Surface

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Material ID: Internal identifier used by Archicad
- Connection Priority: Determines which material takes precedence at junctions
  (Range: 0-999, higher = higher priority)
- Cut Fill: Hatch pattern displayed when element is cut in sections/elevations
- Cut Fill Pen: Pen color/width for the cut fill pattern
- Cut Surface: Material appearance in cut views

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


# =============================================================================
# HELPER FUNCTIONS - GET ATTRIBUTE NAMES
# =============================================================================

def get_fill_name_by_guid(fill_guid):
    """Get Fill attribute name and type from GUID"""
    try:
        # Create AttributeId from GUID
        attr_id = act.AttributeId(fill_guid)

        # Get Fill attributes
        fill_result = acc.GetFillAttributes([attr_id])

        if fill_result and len(fill_result) > 0:
            result = fill_result[0]
            if hasattr(result, 'fillAttribute') and result.fillAttribute:
                fill_attr = result.fillAttribute

                # Get fill name
                name = fill_attr.name if hasattr(
                    fill_attr, 'name') else 'Unknown'

                # Get fill type - subType is a simple string
                fill_type = fill_attr.subType if hasattr(
                    fill_attr, 'subType') else 'N/A'

                return {
                    'name': name,
                    'type': fill_type,
                    'guid': fill_guid
                }

        return {'name': 'Not Found', 'type': 'N/A', 'guid': fill_guid}

    except Exception as e:
        return {'name': f'Error: {str(e)}', 'type': 'N/A', 'guid': fill_guid}


def get_surface_name_by_guid(surface_guid):
    """Get Surface attribute name from GUID"""
    try:
        # Create AttributeId from GUID
        attr_id = act.AttributeId(surface_guid)

        # Get Surface attributes
        surface_result = acc.GetSurfaceAttributes([attr_id])

        if surface_result and len(surface_result) > 0:
            result = surface_result[0]
            if hasattr(result, 'surfaceAttribute') and result.surfaceAttribute:
                surf_attr = result.surfaceAttribute
                name = surf_attr.name if hasattr(
                    surf_attr, 'name') else 'Unknown'

                return {
                    'name': name,
                    'guid': surface_guid
                }

        return {'name': 'Not Found', 'guid': surface_guid}

    except Exception as e:
        return {'name': f'Error: {str(e)}', 'guid': surface_guid}


def get_pen_info(pen_index):
    """Get Pen color name from index"""
    try:
        # Try to get pen table to resolve pen name
        # Note: This requires getting the active pen table
        pen_tables = acc.GetAttributesByType('PenTable')

        if pen_tables and len(pen_tables) > 0:
            # Get first pen table (usually the active one)
            pen_table_result = acc.GetPenTableAttributes([pen_tables[0]])

            if pen_table_result and len(pen_table_result) > 0:
                result = pen_table_result[0]
                if hasattr(result, 'penTableAttribute') and result.penTableAttribute:
                    pen_table = result.penTableAttribute

                    if hasattr(pen_table, 'pens') and pen_table.pens:
                        # Find pen by index
                        for pen in pen_table.pens:
                            if hasattr(pen, 'index') and pen.index == pen_index:
                                pen_name = pen.name if hasattr(
                                    pen, 'name') else f'Pen {pen_index}'

                                # Get color if available
                                color_info = ''
                                if hasattr(pen, 'color'):
                                    color = pen.color
                                    if hasattr(color, 'red') and hasattr(color, 'green') and hasattr(color, 'blue'):
                                        r = int(color.red * 255)
                                        g = int(color.green * 255)
                                        b = int(color.blue * 255)
                                        color_info = f' (RGB: {r},{g},{b})'

                                return {
                                    'index': pen_index,
                                    'name': pen_name,
                                    'color': color_info
                                }

        return {'index': pen_index, 'name': f'Pen {pen_index}', 'color': ''}

    except Exception as e:
        return {'index': pen_index, 'name': f'Pen {pen_index}', 'color': ''}


# =============================================================================
# GET BUILDING MATERIALS WITH ENHANCED CUT FILL INFO
# =============================================================================

def get_all_building_materials_enhanced():
    """Get all building material attributes with enhanced Cut Fill details"""
    try:
        # Get IDs
        print("\n" + "="*80)
        print("BUILDING MATERIALS WITH CUT FILL DETAILS")
        print("="*80)

        attribute_ids = acc.GetAttributesByType('BuildingMaterial')
        print(f"\n✓ Found {len(attribute_ids)} building material(s)")

        if not attribute_ids:
            print("\n⚠️  No building materials found in the project")
            return []

        # Get details
        materials = acc.GetBuildingMaterialAttributes(attribute_ids)

        print(f"\nBuilding Materials List:")
        print("-"*80)

        # Store enhanced data
        enhanced_materials_data = []

        for idx, mat_or_error in enumerate(materials, 1):
            if hasattr(mat_or_error, 'buildingMaterialAttribute') and mat_or_error.buildingMaterialAttribute:
                mat = mat_or_error.buildingMaterialAttribute

                # Basic info
                name = mat.name if hasattr(mat, 'name') else 'Unknown'
                mat_id = mat.id if hasattr(mat, 'id') else 'N/A'

                print(f"\n  {idx:3}. {name}")
                print(f"       Material ID: {mat_id}")

                # Store data for export
                mat_data = {
                    'index': idx,
                    'name': name,
                    'id': mat_id
                }

                # GUID
                if hasattr(mat, 'attributeId') and mat.attributeId:
                    guid = mat.attributeId.guid
                    print(f"       GUID: {guid}")
                    mat_data['guid'] = guid

                # Connection Priority
                if hasattr(mat, 'connectionPriority'):
                    priority = mat.connectionPriority
                    print(f"       Connection Priority: {priority}")
                    mat_data['priority'] = priority

                # ENHANCED CUT FILL - Get name and details
                if hasattr(mat, 'cutFillId'):
                    if hasattr(mat.cutFillId, 'attributeId') and mat.cutFillId.attributeId:
                        fill_guid = mat.cutFillId.attributeId.guid

                        # Get Fill details
                        fill_info = get_fill_name_by_guid(fill_guid)

                        print(f"       Cut Fill: {fill_info['name']}")
                        print(f"          - Type: {fill_info['type']}")
                        print(f"          - GUID: {fill_guid}")

                        mat_data['cut_fill'] = fill_info

                    elif hasattr(mat.cutFillId, 'error'):
                        print(
                            f"       Cut Fill: Error - {mat.cutFillId.error}")
                        mat_data['cut_fill'] = {
                            'name': 'Error', 'type': 'N/A', 'guid': 'N/A'}

                # ENHANCED CUT FILL PEN - Get pen name and color
                if hasattr(mat, 'cutFillPenIndex'):
                    pen_index = mat.cutFillPenIndex
                    pen_info = get_pen_info(pen_index)

                    print(
                        f"       Cut Fill Pen: {pen_info['name']}{pen_info['color']}")
                    mat_data['cut_fill_pen'] = pen_info

                # ENHANCED CUT SURFACE - Get surface name
                if hasattr(mat, 'cutSurfaceId'):
                    if hasattr(mat.cutSurfaceId, 'attributeId') and mat.cutSurfaceId.attributeId:
                        surface_guid = mat.cutSurfaceId.attributeId.guid

                        # Get Surface details
                        surface_info = get_surface_name_by_guid(surface_guid)

                        print(f"       Cut Surface: {surface_info['name']}")
                        print(f"          - GUID: {surface_guid}")

                        mat_data['cut_surface'] = surface_info

                    elif hasattr(mat.cutSurfaceId, 'error'):
                        print(
                            f"       Cut Surface: Error - {mat.cutSurfaceId.error}")
                        mat_data['cut_surface'] = {
                            'name': 'Error', 'guid': 'N/A'}

                enhanced_materials_data.append(mat_data)

        return materials, attribute_ids, enhanced_materials_data

    except Exception as e:
        print(f"\n✗ Error getting building materials: {e}")
        import traceback
        traceback.print_exc()
        return [], [], []


# =============================================================================
# GET FOLDER STRUCTURE
# =============================================================================

def get_folder_structure(attribute_type='BuildingMaterial', path=None, indent=0):
    """Display the folder structure for building materials with material names"""
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

        print(
            f"{prefix}📁 {folder_name} ({num_attributes} material(s), {num_subfolders} subfolder(s))")

        # Display materials in this folder
        if hasattr(structure, 'attributes') and structure.attributes is not None and len(structure.attributes) > 0:
            # Get material names for the attributes in this folder
            attribute_ids = [item.attributeId for item in structure.attributes 
                           if hasattr(item, 'attributeId')]
            
            if attribute_ids:
                try:
                    materials_result = acc.GetBuildingMaterialAttributes(attribute_ids)
                    
                    for mat_item in materials_result:
                        if hasattr(mat_item, 'buildingMaterialAttribute') and mat_item.buildingMaterialAttribute:
                            mat = mat_item.buildingMaterialAttribute
                            mat_name = mat.name if hasattr(mat, 'name') else 'Unknown'
                            print(f"{prefix}   • {mat_name}")
                
                except Exception as e:
                    print(f"{prefix}   ✗ Error loading materials: {e}")

        # Display subfolders recursively
        if hasattr(structure, 'subfolders') and structure.subfolders is not None and len(structure.subfolders) > 0:
            for subfolder_item in structure.subfolders:
                if hasattr(subfolder_item, 'attributeFolder'):
                    folder_struct = subfolder_item.attributeFolder
                    if hasattr(folder_struct, 'name'):
                        subfolder_path = path.copy() if path else []
                        subfolder_path.append(folder_struct.name)
                        get_folder_structure(
                            attribute_type, subfolder_path, indent + 1)

        return structure

    except Exception as e:
        print(f"{prefix}✗ Error: {e}")
        return None


# =============================================================================
# EXPORT TO FILE - ENHANCED
# =============================================================================

def export_to_file_enhanced(enhanced_data, filename='building_materials_detailed.txt'):
    """Export building materials with enhanced Cut Fill info to a text file"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD BUILDING MATERIALS - DETAILED REPORT\n")
            f.write("With Cut Fill Names and Properties\n")
            f.write("="*80 + "\n\n")

            for mat_data in enhanced_data:
                f.write(f"{mat_data['index']}. {mat_data['name']}\n")
                f.write(f"   Material ID: {mat_data['id']}\n")

                if 'guid' in mat_data:
                    f.write(f"   GUID: {mat_data['guid']}\n")

                if 'priority' in mat_data:
                    f.write(
                        f"   Connection Priority: {mat_data['priority']}\n")

                # Enhanced Cut Fill info
                if 'cut_fill' in mat_data:
                    fill = mat_data['cut_fill']
                    f.write(f"   Cut Fill: {fill['name']}\n")
                    f.write(f"      - Type: {fill['type']}\n")
                    f.write(f"      - GUID: {fill['guid']}\n")

                # Enhanced Pen info
                if 'cut_fill_pen' in mat_data:
                    pen = mat_data['cut_fill_pen']
                    f.write(
                        f"   Cut Fill Pen: {pen['name']}{pen['color']}\n")

                # Enhanced Surface info
                if 'cut_surface' in mat_data:
                    surf = mat_data['cut_surface']
                    f.write(f"   Cut Surface: {surf['name']}\n")
                    f.write(f"      - GUID: {surf['guid']}\n")

                f.write("\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Building materials detailed report exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting: {e}")
        import traceback
        traceback.print_exc()


# =============================================================================
# GET ALL CUT FILLS SUMMARY
# =============================================================================

def get_cut_fills_summary(enhanced_data):
    """Display a summary of all unique Cut Fill patterns used"""
    print("\n" + "="*80)
    print("CUT FILLS SUMMARY - All Patterns Used")
    print("="*80)

    # Collect unique fills
    unique_fills = {}

    for mat_data in enhanced_data:
        if 'cut_fill' in mat_data:
            fill_name = mat_data['cut_fill']['name']
            if fill_name not in unique_fills:
                unique_fills[fill_name] = {
                    'type': mat_data['cut_fill']['type'],
                    'guid': mat_data['cut_fill']['guid'],
                    'used_by': []
                }
            unique_fills[fill_name]['used_by'].append(mat_data['name'])

    print(f"\nTotal unique Cut Fill patterns: {len(unique_fills)}")
    print("-"*80)

    for idx, (fill_name, fill_data) in enumerate(sorted(unique_fills.items()), 1):
        print(f"\n  {idx}. {fill_name}")
        print(f"     Type: {fill_data['type']}")
        print(f"     Used by {len(fill_data['used_by'])} material(s):")
        for mat_name in fill_data['used_by'][:5]:  # Show first 5
            print(f"       - {mat_name}")
        if len(fill_data['used_by']) > 5:
            print(f"       ... and {len(fill_data['used_by']) - 5} more")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function"""
    print("\n" + "="*80)
    print("GET BUILDING MATERIALS v1.0")
    print("="*80)

    # Get all materials with enhanced Cut Fill info
    materials, attribute_ids, enhanced_data = get_all_building_materials_enhanced()

    # Show folder structure
    if materials:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('BuildingMaterial')

    # Show Cut Fills Summary
    if enhanced_data:
        get_cut_fills_summary(enhanced_data)

    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Building Materials: {len(materials)}")
    print(
        f"  Materials with Cut Fill info: {sum(1 for m in enhanced_data if 'cut_fill' in m)}")
    print("\n" + "="*80)

    # Export option
    print("\n💡 To export detailed report to file:")
    print("   materials, ids, enhanced = get_all_building_materials_enhanced()")
    print("   export_to_file_enhanced(enhanced, 'my_materials_detailed.txt')")


if __name__ == "__main__":
    main()