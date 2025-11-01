"""
================================================================================
SCRIPT: Get Surfaces with Complete Material Properties
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval


--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Surface attributes in the current Archicad project
with their complete material properties including colors, reflections, 
transparency, and texture information.

Surfaces define the visual appearance and rendering properties of materials.
This script provides:
- List of all surfaces with detailed rendering properties
- Folder hierarchy organization with surfaces listed under each folder
- Material type and rendering parameters
- Color information (surface, specular, emission)
- Reflection properties (ambient, diffuse, specular)
- Transparency and shine values
- Associated fill pattern information
- Texture details when available
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
  
- GetAttributesByType('Surface')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetSurfaceAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSurfaceAttributes

- GetFillAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetFillAttributes

- GetAttributeFolderStructure('Surface')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

[Data Types]
- SurfaceAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.SurfaceAttribute

- RGBColor
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.RGBColor

- Texture
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.Texture

- FillAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac28.html#archicad.releases.ac28.b3001types.FillAttribute

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View surfaces list with complete rendering properties

Optional:
4. Call export_to_file_enhanced('filename.txt') to save results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Folder structure with surfaces listed under each folder
- Total number of surfaces
- For each surface: 
  * Name and GUID
  * Material Type (General/Simple/Matte/Metal/Plastic/Glass/Glowing)
  * Colors (Surface RGB, Specular RGB, Emission RGB)
  * Reflection values (Ambient %, Diffuse %, Specular %)
  * Attenuation values (Transparency, Emission)
  * Transparency % and Shine value
  * Associated Fill pattern
  * Texture information (if present)
- Useful statistics:
  * Material types distribution (if multiple types present)
  * Transparency analysis (opaque vs transparent surfaces)
  * Reflection analysis (high specular/diffuse surfaces)
  * Average reflection values
  * Shine analysis for glossy surfaces
  * Textures and fills count

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
Material Types:
- General: Standard material with all properties
- Simple: Basic material without advanced rendering
- Matte: Non-reflective surface
- Metal: Metallic reflective surface
- Plastic: Plastic-like material
- Glass: Transparent/translucent material
- Glowing: Emissive/self-illuminated material

Reflection Properties:
- Ambient Reflection: Response to ambient lighting (0-100%)
- Diffuse Reflection: Scattered light reflection (0-100%)
- Specular Reflection: Mirror-like reflection (0-100%)

Attenuation:
- Transparency Attenuation: Controls transparency falloff
- Emission Attenuation: Controls emission intensity falloff

Other Properties:
- Transparency: Overall transparency level (0-100%)
- Shine: Surface glossiness/shininess (0-100)

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 115_get_building_materials.py (physical material properties)
- 114_03_get_fills.py (2D hatch patterns)

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

                # Get fill type
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


def format_rgb_color(rgb_color):
    """Format RGB color object to readable string"""
    try:
        if not rgb_color:
            return "N/A"

        r = rgb_color.red if hasattr(rgb_color, 'red') else 0
        g = rgb_color.green if hasattr(rgb_color, 'green') else 0
        b = rgb_color.blue if hasattr(rgb_color, 'blue') else 0

        return f"RGB({r}, {g}, {b})"
    except:
        return "N/A"


def format_texture_info(texture):
    """Format texture information"""
    try:
        if not texture:
            return None

        info = {}

        if hasattr(texture, 'name'):
            info['name'] = texture.name

        if hasattr(texture, 'xSize'):
            info['xSize'] = texture.xSize

        if hasattr(texture, 'ySize'):
            info['ySize'] = texture.ySize

        if hasattr(texture, 'status'):
            info['status'] = texture.status

        return info if info else None
    except:
        return None


# =============================================================================
# GET ALL SURFACES WITH ENHANCED DETAILS
# =============================================================================

def get_all_surfaces_enhanced():
    """Get all surface attributes with complete details"""
    try:
        print("\n" + "="*80)
        print("SURFACES - COMPLETE MATERIAL PROPERTIES")
        print("="*80)

        attribute_ids = acc.GetAttributesByType('Surface')
        print(f"\n✓ Found {len(attribute_ids)} surface(s)")

        if not attribute_ids:
            print("\n⚠️  No surfaces found in the project")
            return [], [], []

        surfaces = acc.GetSurfaceAttributes(attribute_ids)

        print(f"\nSurfaces List with Complete Details:")
        print("-"*80)

        enhanced_data = []

        for idx, surf_or_error in enumerate(surfaces, 1):
            if hasattr(surf_or_error, 'surfaceAttribute') and surf_or_error.surfaceAttribute:
                surf = surf_or_error.surfaceAttribute

                # Basic properties
                name = surf.name if hasattr(surf, 'name') else 'Unknown'
                guid = surf.attributeId.guid if hasattr(
                    surf, 'attributeId') and surf.attributeId else 'N/A'
                material_type = surf.materialType if hasattr(
                    surf, 'materialType') else 'N/A'

                # Reflection properties (0-100)
                ambient_refl = surf.ambientReflection if hasattr(
                    surf, 'ambientReflection') else 0
                diffuse_refl = surf.diffuseReflection if hasattr(
                    surf, 'diffuseReflection') else 0
                specular_refl = surf.specularReflection if hasattr(
                    surf, 'specularReflection') else 0

                # Attenuation properties
                transp_atten = surf.transparencyAttenuation if hasattr(
                    surf, 'transparencyAttenuation') else 0
                emission_atten = surf.emissionAttenuation if hasattr(
                    surf, 'emissionAttenuation') else 0

                # Color properties
                surface_color = surf.surfaceColor if hasattr(
                    surf, 'surfaceColor') else None
                specular_color = surf.specularColor if hasattr(
                    surf, 'specularColor') else None
                emission_color = surf.emissionColor if hasattr(
                    surf, 'emissionColor') else None

                # Other properties
                transparency = surf.transparency if hasattr(
                    surf, 'transparency') else 0
                shine = surf.shine if hasattr(surf, 'shine') else 0

                # Fill reference
                fill_info = None
                if hasattr(surf, 'fillId'):
                    fill_id = surf.fillId
                    if hasattr(fill_id, 'attributeId') and fill_id.attributeId:
                        fill_guid = fill_id.attributeId.guid
                        fill_info = get_fill_name_by_guid(fill_guid)

                # Texture
                texture_info = None
                if hasattr(surf, 'texture') and surf.texture:
                    texture_info = format_texture_info(surf.texture)

                # Display
                print(f"\n  {idx:3}. {name}")
                print(f"       GUID: {guid}")
                print(f"       Material Type: {material_type}")

                print(f"\n       Colors:")
                print(
                    f"         • Surface:  {format_rgb_color(surface_color)}")
                print(
                    f"         • Specular: {format_rgb_color(specular_color)}")
                print(
                    f"         • Emission: {format_rgb_color(emission_color)}")

                print(f"\n       Reflection Properties:")
                print(f"         • Ambient:  {ambient_refl}%")
                print(f"         • Diffuse:  {diffuse_refl}%")
                print(f"         • Specular: {specular_refl}%")

                print(f"\n       Attenuation:")
                print(f"         • Transparency: {transp_atten}")
                print(f"         • Emission:     {emission_atten}")

                print(f"\n       Other Properties:")
                print(f"         • Transparency: {transparency}%")
                print(f"         • Shine:        {shine}")

                if fill_info:
                    print(f"\n       Associated Fill:")
                    print(f"         • Name: {fill_info['name']}")
                    print(f"         • Type: {fill_info['type']}")

                if texture_info:
                    print(f"\n       Texture:")
                    if 'name' in texture_info:
                        print(f"         • Name: {texture_info['name']}")
                    if 'xSize' in texture_info and 'ySize' in texture_info:
                        print(
                            f"         • Size: {texture_info['xSize']} × {texture_info['ySize']}")
                    if 'status' in texture_info:
                        print(f"         • Status: {texture_info['status']}")

                # Store enhanced data
                enhanced_data.append({
                    'index': idx,
                    'name': name,
                    'guid': guid,
                    'material_type': material_type,
                    'ambient_reflection': ambient_refl,
                    'diffuse_reflection': diffuse_refl,
                    'specular_reflection': specular_refl,
                    'transparency_attenuation': transp_atten,
                    'emission_attenuation': emission_atten,
                    'surface_color': format_rgb_color(surface_color),
                    'specular_color': format_rgb_color(specular_color),
                    'emission_color': format_rgb_color(emission_color),
                    'transparency': transparency,
                    'shine': shine,
                    'fill': fill_info,
                    'texture': texture_info
                })

        return surfaces, attribute_ids, enhanced_data

    except Exception as e:
        print(f"\n✗ Error getting surfaces: {e}")
        import traceback
        traceback.print_exc()
        return [], [], []


# =============================================================================
# GET FOLDER STRUCTURE
# =============================================================================

def get_folder_structure(attribute_type='Surface', path=None, indent=0):
    """Display the folder structure for surfaces with surface names"""
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
            f"{prefix}📁 {folder_name} ({num_attributes} surface(s), {num_subfolders} subfolder(s))")

        # Display surfaces in this folder
        if hasattr(structure, 'attributes') and structure.attributes is not None and len(structure.attributes) > 0:
            # Get surface names for the attributes in this folder
            attribute_ids = [item.attributeId for item in structure.attributes
                             if hasattr(item, 'attributeId')]

            if attribute_ids:
                try:
                    surfaces_result = acc.GetSurfaceAttributes(attribute_ids)

                    for surf_item in surfaces_result:
                        if hasattr(surf_item, 'surfaceAttribute') and surf_item.surfaceAttribute:
                            surf = surf_item.surfaceAttribute
                            surf_name = surf.name if hasattr(
                                surf, 'name') else 'Unknown'
                            mat_type = surf.materialType if hasattr(
                                surf, 'materialType') else 'N/A'
                            print(f"{prefix}   • {surf_name} [{mat_type}]")

                except Exception as e:
                    print(f"{prefix}   ✗ Error loading surfaces: {e}")

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
# GET SURFACES STATISTICS
# =============================================================================

def get_surfaces_statistics(enhanced_data):
    """Display useful statistics about surfaces"""
    print("\n" + "="*80)
    print("SURFACES STATISTICS")
    print("="*80)

    # Material types
    material_types = {}
    for surf_data in enhanced_data:
        mat_type = surf_data['material_type']
        if mat_type not in material_types:
            material_types[mat_type] = 0
        material_types[mat_type] += 1

    # Only show material types if there's more than one type
    if len(material_types) > 1:
        print(f"\nMaterial Types Distribution:")
        print("-"*80)
        for mat_type, count in sorted(material_types.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / len(enhanced_data)) * 100
            print(f"  • {mat_type}: {count} ({percentage:.1f}%)")

    # Transparency analysis
    transparent_surfaces = [s for s in enhanced_data if s['transparency'] > 0]
    opaque_surfaces = [s for s in enhanced_data if s['transparency'] == 0]

    print(f"\nTransparency Analysis:")
    print("-"*80)
    print(f"  • Opaque surfaces (0%): {len(opaque_surfaces)}")
    print(f"  • Transparent surfaces (>0%): {len(transparent_surfaces)}")
    if transparent_surfaces:
        avg_transparency = sum(
            s['transparency'] for s in transparent_surfaces) / len(transparent_surfaces)
        print(f"    - Average transparency: {avg_transparency:.1f}%")
        max_transp = max(transparent_surfaces, key=lambda x: x['transparency'])
        print(
            f"    - Most transparent: {max_transp['name']} ({max_transp['transparency']}%)")

    # Reflection analysis
    print(f"\nReflection Analysis:")
    print("-"*80)
    high_specular = [s for s in enhanced_data if s['specular_reflection'] > 50]
    high_diffuse = [s for s in enhanced_data if s['diffuse_reflection'] > 70]

    if high_specular:
        print(f"  • High specular reflection (>50%): {len(high_specular)}")
        print(
            f"    Examples: {', '.join([s['name'] for s in high_specular[:3]])}")

    if high_diffuse:
        print(f"  • High diffuse reflection (>70%): {len(high_diffuse)}")

    # Average values
    avg_ambient = sum(s['ambient_reflection']
                      for s in enhanced_data) / len(enhanced_data)
    avg_diffuse = sum(s['diffuse_reflection']
                      for s in enhanced_data) / len(enhanced_data)
    avg_specular = sum(s['specular_reflection']
                       for s in enhanced_data) / len(enhanced_data)

    print(f"\n  Average reflection values:")
    print(f"    - Ambient:  {avg_ambient:.1f}%")
    print(f"    - Diffuse:  {avg_diffuse:.1f}%")
    print(f"    - Specular: {avg_specular:.1f}%")

    # Shine analysis
    high_shine = [s for s in enhanced_data if s['shine'] > 50]
    if high_shine:
        print(f"\nShine Analysis:")
        print("-"*80)
        print(f"  • Surfaces with high shine (>50): {len(high_shine)}")
        avg_shine = sum(s['shine'] for s in high_shine) / len(high_shine)
        print(f"  • Average shine for glossy surfaces: {avg_shine:.1f}")

    # Texture and Fill
    with_texture = [s for s in enhanced_data if s['texture']]
    with_fill = [s for s in enhanced_data if s['fill']]

    print(f"\nTextures & Fills:")
    print("-"*80)
    print(f"  • Surfaces with texture: {len(with_texture)}")
    print(f"  • Surfaces with associated fill: {len(with_fill)}")

    if with_texture:
        print(f"\n  Textured surfaces examples:")
        for s in with_texture[:3]:
            print(f"    - {s['name']}")


# =============================================================================
# EXPORT TO FILE - ENHANCED
# =============================================================================

def export_to_file_enhanced(enhanced_data, filename='surfaces_detailed.txt'):
    """Export surfaces with complete properties to a text file"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD SURFACES - DETAILED REPORT\n")
            f.write("Complete Material Properties\n")
            f.write("="*80 + "\n\n")

            for surf_data in enhanced_data:
                f.write(f"{surf_data['index']}. {surf_data['name']}\n")
                f.write(f"   GUID: {surf_data['guid']}\n")
                f.write(f"   Material Type: {surf_data['material_type']}\n\n")

                f.write(f"   Colors:\n")
                f.write(f"     • Surface:  {surf_data['surface_color']}\n")
                f.write(f"     • Specular: {surf_data['specular_color']}\n")
                f.write(f"     • Emission: {surf_data['emission_color']}\n\n")

                f.write(f"   Reflection Properties:\n")
                f.write(
                    f"     • Ambient:  {surf_data['ambient_reflection']}%\n")
                f.write(
                    f"     • Diffuse:  {surf_data['diffuse_reflection']}%\n")
                f.write(
                    f"     • Specular: {surf_data['specular_reflection']}%\n\n")

                f.write(f"   Attenuation:\n")
                f.write(
                    f"     • Transparency: {surf_data['transparency_attenuation']}\n")
                f.write(
                    f"     • Emission:     {surf_data['emission_attenuation']}\n\n")

                f.write(f"   Other Properties:\n")
                f.write(f"     • Transparency: {surf_data['transparency']}%\n")
                f.write(f"     • Shine:        {surf_data['shine']}\n")

                # Fill info
                if surf_data['fill']:
                    fill = surf_data['fill']
                    f.write(f"\n   Associated Fill:\n")
                    f.write(f"     • Name: {fill['name']}\n")
                    f.write(f"     • Type: {fill['type']}\n")

                # Texture info
                if surf_data['texture']:
                    texture = surf_data['texture']
                    f.write(f"\n   Texture:\n")
                    if 'name' in texture:
                        f.write(f"     • Name: {texture['name']}\n")
                    if 'xSize' in texture and 'ySize' in texture:
                        f.write(
                            f"     • Size: {texture['xSize']} × {texture['ySize']}\n")
                    if 'status' in texture:
                        f.write(f"     • Status: {texture['status']}\n")

                f.write("\n" + "-"*80 + "\n\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Surfaces detailed report exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting: {e}")
        import traceback
        traceback.print_exc()


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function"""
    print("\n" + "="*80)
    print("GET SURFACES v2.0 - Complete Material Properties")
    print("="*80)

    # Get all surfaces with enhanced details
    surfaces, attribute_ids, enhanced_data = get_all_surfaces_enhanced()

    # Show folder structure
    if surfaces:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('Surface')

    # Show Statistics
    if enhanced_data:
        get_surfaces_statistics(enhanced_data)

    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Surfaces: {len(surfaces)}")
    print(
        f"  Surfaces with Texture: {sum(1 for s in enhanced_data if s['texture'])}")
    print(
        f"  Surfaces with Fill: {sum(1 for s in enhanced_data if s['fill'])}")
    print("\n" + "="*80)

    # Export option
    print("\n💡 To export detailed report to file:")
    print("   surfaces, ids, enhanced = get_all_surfaces_enhanced()")
    print("   export_to_file_enhanced(enhanced, 'my_surfaces_detailed.txt')")


if __name__ == "__main__":
    main()
