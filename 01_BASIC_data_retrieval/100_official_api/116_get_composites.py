"""
================================================================================
SCRIPT: Get Composites
================================================================================

Version:        1.2
Date:           2025-10-29
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Composite attributes in the current Archicad project
with complete details including skins, materials, and folder organization.

Composites are multi-layer structural assemblies used in walls, slabs, and roofs.
This script provides:
- List of all composites with complete details
- Total thickness of each composite
- Detailed skin breakdown (material, thickness, function)
- Compatible element types (Wall, Slab, Roof, etc.)
- Folder hierarchy organization with composites listed
- GUID and index mapping
- Optional export to text file

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAttributesByType('Composite')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetCompositeAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetCompositeAttributes

- GetBuildingMaterialAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetBuildingMaterialAttributes

- GetAttributeFolderStructure('Composite')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

- GetAttributesIndices() - Archicad 28+
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesIndices

[Data Types]
- CompositeAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.CompositeAttribute

- CompositeSkin
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.CompositeSkin

- BuildingMaterialAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.BuildingMaterialAttribute

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View composites list with skin counts and folder structure

Optional:
4. Call export_to_file('filename.txt') to save results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of composites
- For each composite:
  * Name
  * GUID
  * Total thickness (cm)
  * Compatible element types
  * Number of skins
  * For each skin:
    - Material name
    - Thickness (cm)
    - Type (Core, Finish, or Other)
    - Frame pen index (if defined)
- Folder structure hierarchy with composites listed
- GUID/Index mapping (Archicad 28+)
- Summary statistics

Optional export creates a detailed text file with all information.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Composites are multi-layer assemblies (skins)
- Each skin has a material and thickness
- Skin type is determined by isCore and isFinish flags:
  * Core: Structural layer (isCore = true)
  * Finish: Finish layer (isFinish = true)
  * Other: Neither core nor finish
- Compatible element types indicate where the composite can be used
- Thickness is displayed in centimeters
- GUID is permanent, index is project-specific
- Composite lines (contour and separator lines) are only displayed if they are 
  enabled in Archicad. A line is enabled when linePenIndex != 0.
  When linePenIndex = 0, the line is disabled/hidden.

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_building_material_name(material_guid):
    """Get Building Material name from GUID"""
    try:
        attr_id = act.AttributeId(material_guid)
        mat_result = acc.GetBuildingMaterialAttributes([attr_id])
        
        if mat_result and len(mat_result) > 0:
            result = mat_result[0]
            if hasattr(result, 'buildingMaterialAttribute') and result.buildingMaterialAttribute:
                mat = result.buildingMaterialAttribute
                return mat.name if hasattr(mat, 'name') else 'Unknown'
        
        return 'Material not found'
    
    except Exception as e:
        return f'Error: {str(e)}'


# =============================================================================
# GET COMPOSITES
# =============================================================================

def get_all_composites():
    """Get all composite attributes with complete details"""
    try:
        # Get IDs
        print("\n" + "="*80)
        print("COMPOSITES")
        print("="*80)
        
        attribute_ids = acc.GetAttributesByType('Composite')
        print(f"\n✓ Found {len(attribute_ids)} composite(s)")

        if not attribute_ids:
            print("\n⚠️  No composites found in the project")
            return [], []

        # Get details
        composites = acc.GetCompositeAttributes(attribute_ids)

        print(f"\nComposites List:")
        print("-"*80)

        enhanced_composites_data = []

        for idx, comp_or_error in enumerate(composites, 1):
            if hasattr(comp_or_error, 'compositeAttribute') and comp_or_error.compositeAttribute:
                comp = comp_or_error.compositeAttribute
                name = comp.name if hasattr(comp, 'name') else 'Unknown'
                
                print(f"\n  {idx:3}. {name}")
                
                comp_data = {
                    'index': idx,
                    'name': name
                }
                
                # Get GUID
                if hasattr(comp, 'attributeId') and comp.attributeId:
                    guid = comp.attributeId.guid
                    print(f"       GUID: {guid}")
                    comp_data['guid'] = guid
                
                # Get total thickness
                if hasattr(comp, 'totalThickness'):
                    thickness_m = comp.totalThickness
                    thickness_cm = thickness_m * 100
                    print(f"       Total Thickness: {thickness_cm:.2f} cm")
                    comp_data['thickness_m'] = thickness_m
                    comp_data['thickness_cm'] = thickness_cm
                
                # Get useWith (compatible element types)
                if hasattr(comp, 'useWith') and comp.useWith:
                    use_with_str = ', '.join(comp.useWith)
                    print(f"       Compatible with: {use_with_str}")
                    comp_data['useWith'] = comp.useWith
                
                # Get skins details
                if hasattr(comp, 'compositeSkins') and comp.compositeSkins:
                    num_skins = len(comp.compositeSkins)
                    print(f"       Skins: {num_skins}")
                    comp_data['skins'] = []
                    
                    # Get composite lines for positioning
                    composite_lines = []
                    if hasattr(comp, 'compositeLines') and comp.compositeLines:
                        composite_lines = comp.compositeLines
                    
                    # Track which lines exist
                    comp_data['has_top_line'] = False
                    comp_data['separator_lines'] = []  # Boolean for each separator position
                    comp_data['has_bottom_line'] = False
                    
                    # Check top contour line (index 0)
                    if len(composite_lines) > 0:
                        line_item = composite_lines[0]
                        if hasattr(line_item, 'compositeLine') and line_item.compositeLine:
                            comp_line = line_item.compositeLine
                            # Check if line is enabled: linePenIndex must exist and be != 0
                            # linePenIndex = 0 means "no line" (disabled)
                            if (hasattr(comp_line, 'linePenIndex') and 
                                comp_line.linePenIndex is not None and 
                                comp_line.linePenIndex != 0):
                                comp_data['has_top_line'] = True
                                print(f"          ======== [Top Contour Line]")
                    
                    for skin_idx, skin_item in enumerate(comp.compositeSkins, 1):
                        # Access CompositeSkin from CompositeSkinListItem
                        if hasattr(skin_item, 'compositeSkin') and skin_item.compositeSkin:
                            skin = skin_item.compositeSkin
                            
                            skin_data = {'index': skin_idx}
                            
                            # Get material name
                            material_name = 'Unknown'
                            if hasattr(skin, 'buildingMaterialId'):
                                if hasattr(skin.buildingMaterialId, 'attributeId') and skin.buildingMaterialId.attributeId:
                                    material_guid = skin.buildingMaterialId.attributeId.guid
                                    material_name = get_building_material_name(material_guid)
                                    skin_data['material_guid'] = material_guid
                            
                            skin_data['material_name'] = material_name
                            
                            # Get thickness
                            skin_thickness_cm = 0.0
                            if hasattr(skin, 'thickness'):
                                skin_thickness_m = skin.thickness
                                skin_thickness_cm = skin_thickness_m * 100
                                skin_data['thickness_m'] = skin_thickness_m
                                skin_data['thickness_cm'] = skin_thickness_cm
                            
                            # Get skin type from isCore and isFinish flags
                            skin_type = 'Other'
                            try:
                                if skin.isCore:
                                    skin_type = 'Core'
                                elif skin.isFinish:
                                    skin_type = 'Finish'
                            except:
                                pass  # Keep 'Other' as default
                            
                            skin_data['skinType'] = skin_type
                            
                            # Get frame pen index
                            if hasattr(skin, 'framePenIndex') and skin.framePenIndex is not None:
                                skin_data['framePenIndex'] = skin.framePenIndex
                            
                            # Display skin
                            print(f"          {skin_idx}. {material_name} ({skin_thickness_cm:.2f} cm) [{skin_type}]")
                            
                            # Check separator line after this skin (if not the last skin)
                            has_separator = False
                            if skin_idx < num_skins:
                                line_index = skin_idx
                                if line_index < len(composite_lines):
                                    line_item = composite_lines[line_index]
                                    if hasattr(line_item, 'compositeLine') and line_item.compositeLine:
                                        comp_line = line_item.compositeLine
                                        # Check if line is enabled: linePenIndex must exist and be != 0
                                        if (hasattr(comp_line, 'linePenIndex') and 
                                            comp_line.linePenIndex is not None and 
                                            comp_line.linePenIndex != 0):
                                            has_separator = True
                                            print(f"          -------- [Separator Line]")
                            
                            comp_data['separator_lines'].append(has_separator)
                            comp_data['skins'].append(skin_data)
                    
                    # Check bottom contour line (index num_skins)
                    if len(composite_lines) > num_skins:
                        line_item = composite_lines[num_skins]
                        if hasattr(line_item, 'compositeLine') and line_item.compositeLine:
                            comp_line = line_item.compositeLine
                            # Check if line is enabled: linePenIndex must exist and be != 0
                            if (hasattr(comp_line, 'linePenIndex') and 
                                comp_line.linePenIndex is not None and 
                                comp_line.linePenIndex != 0):
                                comp_data['has_bottom_line'] = True
                                print(f"          ======== [Bottom Contour Line]")
                
                # Store number of composite lines (for reference)
                if hasattr(comp, 'compositeLines') and comp.compositeLines:
                    comp_data['num_lines'] = len(comp.compositeLines)
                
                enhanced_composites_data.append(comp_data)

        return composites, attribute_ids, enhanced_composites_data

    except Exception as e:
        print(f"\n✗ Error getting composites: {e}")
        import traceback
        traceback.print_exc()
        return [], [], []


# =============================================================================
# GET FOLDER STRUCTURE
# =============================================================================

def get_folder_structure(attribute_type='Composite', path=None, indent=0):
    """Display the folder structure for composites with composite names"""
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
        
        print(f"{prefix}📁 {folder_name} ({num_attributes} composite(s), {num_subfolders} subfolder(s))")
        
        # Display composites in this folder
        if hasattr(structure, 'attributes') and structure.attributes is not None and len(structure.attributes) > 0:
            # Get composite names for the attributes in this folder
            attribute_ids = [item.attributeId for item in structure.attributes 
                           if hasattr(item, 'attributeId')]
            
            if attribute_ids:
                try:
                    composites_result = acc.GetCompositeAttributes(attribute_ids)
                    
                    for comp_item in composites_result:
                        if hasattr(comp_item, 'compositeAttribute') and comp_item.compositeAttribute:
                            comp = comp_item.compositeAttribute
                            comp_name = comp.name if hasattr(comp, 'name') else 'Unknown'
                            
                            # Get skin count and thickness
                            num_skins = 0
                            thickness_str = ""
                            if hasattr(comp, 'compositeSkins') and comp.compositeSkins:
                                num_skins = len(comp.compositeSkins)
                            if hasattr(comp, 'totalThickness'):
                                thickness_cm = comp.totalThickness * 100
                                thickness_str = f" - {thickness_cm:.1f} cm"
                            
                            print(f"{prefix}   • {comp_name} ({num_skins} skins{thickness_str})")
                
                except Exception as e:
                    print(f"{prefix}   ✗ Error loading composites: {e}")
        
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


# =============================================================================
# GET ATTRIBUTE INDICES
# =============================================================================

def get_indices_mapping(attribute_ids):
    """Get GUID to Index mapping (if API is available)"""
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


# =============================================================================
# EXPORT TO FILE
# =============================================================================

def export_to_file_enhanced(enhanced_data, filename='composites_detailed.txt'):
    """Export composites with complete details to a text file"""
    try:
        import os
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD COMPOSITES - DETAILED REPORT\n")
            f.write("="*80 + "\n\n")
            
            for comp_data in enhanced_data:
                f.write(f"{comp_data['index']}. {comp_data['name']}\n")
                
                if 'guid' in comp_data:
                    f.write(f"   GUID: {comp_data['guid']}\n")
                
                if 'thickness_cm' in comp_data:
                    f.write(f"   Total Thickness: {comp_data['thickness_cm']:.2f} cm\n")
                
                if 'useWith' in comp_data:
                    use_with_str = ', '.join(comp_data['useWith'])
                    f.write(f"   Compatible with: {use_with_str}\n")
                
                if 'skins' in comp_data:
                    f.write(f"   Skins: {len(comp_data['skins'])}\n")
                    
                    # Display top contour line if it exists
                    if comp_data.get('has_top_line', False):
                        f.write(f"      ======== [Top Contour Line]\n")
                    
                    for idx, skin in enumerate(comp_data['skins'], 1):
                        f.write(f"      {skin['index']}. {skin['material_name']} ({skin['thickness_cm']:.2f} cm) [{skin['skinType']}]\n")
                        if 'framePenIndex' in skin:
                            f.write(f"         Frame Pen Index: {skin['framePenIndex']}\n")
                        
                        # Display separator line if it exists for this position
                        if idx - 1 < len(comp_data.get('separator_lines', [])):
                            if comp_data['separator_lines'][idx - 1]:
                                f.write(f"      -------- [Separator Line]\n")
                    
                    # Display bottom contour line if it exists
                    if comp_data.get('has_bottom_line', False):
                        f.write(f"      ======== [Bottom Contour Line]\n")
                
                f.write("\n")
        
        abs_path = os.path.abspath(filename)
        print(f"\n✓ Composites detailed report exported to:")
        print(f"  {abs_path}")
        
    except Exception as e:
        print(f"\n✗ Error exporting: {e}")
        import traceback
        traceback.print_exc()


def export_to_file(composites, filename='composites.txt'):
    """Export composites to a text file (basic version)"""
    try:
        import os
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD COMPOSITES\n")
            f.write("="*80 + "\n\n")
            
            for idx, comp_or_error in enumerate(composites, 1):
                if hasattr(comp_or_error, 'compositeAttribute') and comp_or_error.compositeAttribute:
                    comp = comp_or_error.compositeAttribute
                    name = comp.name if hasattr(comp, 'name') else 'Unknown'
                    
                    # Get number of skins
                    num_skins = 0
                    if hasattr(comp, 'compositeSkins') and comp.compositeSkins:
                        num_skins = len(comp.compositeSkins)
                    elif hasattr(comp, 'skins') and comp.skins:
                        num_skins = len(comp.skins)
                    
                    f.write(f"{idx}. {name} ({num_skins} skin(s))\n")
                    
                    if hasattr(comp, 'attributeId') and comp.attributeId:
                        f.write(f"   GUID: {comp.attributeId.guid}\n")
                    
                    f.write("\n")
        
        abs_path = os.path.abspath(filename)
        print(f"\n✓ Composites exported to:")
        print(f"  {abs_path}")
        
    except Exception as e:
        print(f"\n✗ Error exporting: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function"""
    print("\n" + "="*80)
    print("GET COMPOSITES v1.2")
    print("="*80)
    
    # Get all composites with detailed information
    composites, attribute_ids, enhanced_data = get_all_composites()
    
    # Show folder structure
    if composites:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('Composite')
    
    # Show GUID/Index mapping
    if attribute_ids:
        get_indices_mapping(attribute_ids)
    
    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Composites: {len(composites)}")
    print("\n" + "="*80)
    
    # Export option
    print("\n💡 To export detailed report to file:")
    print("   composites, ids, enhanced = get_all_composites()")
    print("   export_to_file_enhanced(enhanced, 'my_composites_detailed.txt')")


if __name__ == "__main__":
    main()