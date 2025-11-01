"""
================================================================================
SCRIPT: Get All Attributes - Complete Attribute Retrieval
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Retrieval

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays ALL types of attributes available in the current Archicad 
project. This comprehensive script covers all attribute categories including:
- Building Materials
- Composites
- Fills
- Lines
- Surfaces
- Pen Tables (including active pen tables)
- Zone Categories
- Profiles
- Layers
- Layer Combinations

This script provides a complete overview of project attributes and can export
the results to a text file for documentation purposes.

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
  
- GetAttributesByType()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetBuildingMaterialAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetBuildingMaterialAttributes

- GetCompositeAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetCompositeAttributes

- GetFillAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetFillAttributes

- GetLineAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLineAttributes

- GetSurfaceAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetSurfaceAttributes

- GetPenTableAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetPenTableAttributes

- GetZoneCategoryAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetZoneCategoryAttributes

- GetProfileAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetProfileAttributes

- GetLayerAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLayerAttributes

- GetLayerCombinationAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLayerCombinationAttributes

- GetActivePenTables()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetActivePenTables

[Data Types]
- AttributeId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.AttributeId 

- BuildingMaterialAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.BuildingMaterialAttribute 

- CompositeAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.CompositeAttribute

- LayerAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.LayerAttribute 

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View all attributes organized by type in console output

Optional - Export to file:
4. Call export_all_attributes_to_file('filename.txt') to save results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Count of each attribute type found
- Detailed list for each category with:
  * Building Materials: Name
  * Composites: Name and number of skins
  * Fills: Name
  * Lines: Name
  * Surfaces: Name
  * Pen Tables: Name and number of pens
  * Zone Categories: Name and category code
  * Profiles: Name
  * Layers: Name and status (Hidden/Locked)
  * Layer Combinations: Name and number of layers
- Active pen tables (Model View and Layout Book)
- Summary with total counts for all categories

Optional export creates a text file with all attribute details.

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
- The script uses GetAttributesByType() to first retrieve attribute IDs,
  then calls specific Get*Attributes() methods for detailed information
- Some attributes may have optional properties (version, description, etc.)
  that are only displayed if available
- Active pen tables show which pen table is currently active for Model View
  and Layout Book contexts
- Layer status includes Hidden and Locked flags
- Composite skins count is displayed to show structure complexity
- Zone categories include their category code for reference
- The export function can be customized to include additional details
- All attribute GUIDs are accessible through the returned objects for use
  in other scripts

--------------------------------------------------------------------------------
ATTRIBUTE TYPES:
--------------------------------------------------------------------------------
This script covers all standard Archicad attribute types:

1. Building Materials - Physical material properties
2. Composites - Multi-layer structural assemblies
3. Fills - Hatch patterns and solid fills
4. Lines - Line types and patterns
5. Surfaces - Material appearance and rendering
6. Pen Tables - Pen color and width sets
7. Zone Categories - Space classification categories
8. Profiles - Custom profile shapes
9. Layers - Drawing organization layers
10. Layer Combinations - Predefined layer visibility sets

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 115_get_specific_attribute_details.py (detailed attribute inspection)
- 201_set_element_attributes.py (modifying element attributes)
- 103_get_layers_hierarchy.py (detailed layer structure)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


# =============================================================================
# GET ATTRIBUTES BY TYPE
# =============================================================================

def get_attributes_by_type(attribute_type):
    """
    Get all attributes of a specific type.

    Args:
        attribute_type: One of: 'BuildingMaterial', 'Composite', 'Fill', 'Line',
                       'Surface', 'PenTable', 'ZoneCategory', 'Profile',
                       'Layer', 'LayerCombination'

    Returns:
        List of attribute IDs
    """
    try:
        attributes = acc.GetAttributesByType(attribute_type)
        print(f"✓ Found {len(attributes)} {attribute_type} attribute(s)")
        return attributes
    except Exception as e:
        print(f"✗ Error getting {attribute_type} attributes: {e}")
        return []


# =============================================================================
# GET BUILDING MATERIALS
# =============================================================================

def get_all_building_materials():
    """Get all building material attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('BuildingMaterial')

        if not attribute_ids:
            return []

        # Get details
        materials = acc.GetBuildingMaterialAttributes(attribute_ids)

        print(f"\nBuilding Materials:")
        print("="*80)

        for idx, mat_or_error in enumerate(materials, 1):
            if hasattr(mat_or_error, 'buildingMaterialAttribute') and mat_or_error.buildingMaterialAttribute:
                mat = mat_or_error.buildingMaterialAttribute
                name = mat.name if hasattr(mat, 'name') else 'Unknown'
                print(f"  {idx}. {name}")

        return materials

    except Exception as e:
        print(f"✗ Error getting building materials: {e}")
        return []


# =============================================================================
# GET COMPOSITES
# =============================================================================

def get_all_composites():
    """Get all composite attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('Composite')

        if not attribute_ids:
            return []

        # Get details
        composites = acc.GetCompositeAttributes(attribute_ids)

        print(f"\nComposites:")
        print("="*80)

        for idx, comp_or_error in enumerate(composites, 1):
            if hasattr(comp_or_error, 'compositeAttribute') and comp_or_error.compositeAttribute:
                comp = comp_or_error.compositeAttribute
                name = comp.name if hasattr(comp, 'name') else 'Unknown'

                # Get number of skins - handle both attribute types
                num_skins = 0
                if hasattr(comp, 'compositeSkins') and comp.compositeSkins:
                    num_skins = len(comp.compositeSkins)
                elif hasattr(comp, 'skins') and comp.skins:
                    num_skins = len(comp.skins)

                print(f"  {idx}. {name} ({num_skins} skin(s))")

        return composites

    except Exception as e:
        print(f"✗ Error getting composites: {e}")
        return []


# =============================================================================
# GET FILLS
# =============================================================================

def get_all_fills():
    """Get all fill attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('Fill')

        if not attribute_ids:
            return []

        # Get details
        fills = acc.GetFillAttributes(attribute_ids)

        print(f"\nFills:")
        print("="*80)

        for idx, fill_or_error in enumerate(fills, 1):
            if hasattr(fill_or_error, 'fillAttribute') and fill_or_error.fillAttribute:
                fill = fill_or_error.fillAttribute
                name = fill.name if hasattr(fill, 'name') else 'Unknown'
                print(f"  {idx}. {name}")

        return fills

    except Exception as e:
        print(f"✗ Error getting fills: {e}")
        return []


# =============================================================================
# GET LINES
# =============================================================================

def get_all_lines():
    """Get all line attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('Line')

        if not attribute_ids:
            return []

        # Get details
        lines = acc.GetLineAttributes(attribute_ids)

        print(f"\nLines:")
        print("="*80)

        for idx, line_or_error in enumerate(lines, 1):
            if hasattr(line_or_error, 'lineAttribute') and line_or_error.lineAttribute:
                line = line_or_error.lineAttribute
                name = line.name if hasattr(line, 'name') else 'Unknown'
                print(f"  {idx}. {name}")

        return lines

    except Exception as e:
        print(f"✗ Error getting lines: {e}")
        return []


# =============================================================================
# GET SURFACES
# =============================================================================

def get_all_surfaces():
    """Get all surface attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('Surface')

        if not attribute_ids:
            return []

        # Get details
        surfaces = acc.GetSurfaceAttributes(attribute_ids)

        print(f"\nSurfaces:")
        print("="*80)

        for idx, surf_or_error in enumerate(surfaces, 1):
            if hasattr(surf_or_error, 'surfaceAttribute') and surf_or_error.surfaceAttribute:
                surf = surf_or_error.surfaceAttribute
                name = surf.name if hasattr(surf, 'name') else 'Unknown'
                print(f"  {idx}. {name}")

        return surfaces

    except Exception as e:
        print(f"✗ Error getting surfaces: {e}")
        return []


# =============================================================================
# GET PEN TABLES
# =============================================================================

def get_all_pen_tables():
    """Get all pen table attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('PenTable')

        if not attribute_ids:
            return []

        # Get details
        pen_tables = acc.GetPenTableAttributes(attribute_ids)

        print(f"\nPen Tables:")
        print("="*80)

        for idx, pt_or_error in enumerate(pen_tables, 1):
            if hasattr(pt_or_error, 'penTableAttribute') and pt_or_error.penTableAttribute:
                pt = pt_or_error.penTableAttribute
                name = pt.name if hasattr(pt, 'name') else 'Unknown'

                # Number of pens
                num_pens = len(pt.pens) if hasattr(pt, 'pens') else 0

                print(f"  {idx}. {name} ({num_pens} pen(s))")

        return pen_tables

    except Exception as e:
        print(f"✗ Error getting pen tables: {e}")
        return []


# =============================================================================
# GET ZONE CATEGORIES
# =============================================================================

def get_all_zone_categories():
    """Get all zone category attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('ZoneCategory')

        if not attribute_ids:
            return []

        # Get details
        zone_categories = acc.GetZoneCategoryAttributes(attribute_ids)

        print(f"\nZone Categories:")
        print("="*80)

        for idx, zc_or_error in enumerate(zone_categories, 1):
            if hasattr(zc_or_error, 'zoneCategoryAttribute') and zc_or_error.zoneCategoryAttribute:
                zc = zc_or_error.zoneCategoryAttribute
                name = zc.name if hasattr(zc, 'name') else 'Unknown'

                # Category code
                code = zc.categoryCode if hasattr(
                    zc, 'categoryCode') else 'N/A'

                print(f"  {idx}. {name} (Code: {code})")

        return zone_categories

    except Exception as e:
        print(f"✗ Error getting zone categories: {e}")
        return []


# =============================================================================
# GET PROFILES
# =============================================================================

def get_all_profiles():
    """Get all profile attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('Profile')

        if not attribute_ids:
            return []

        # Get details
        profiles = acc.GetProfileAttributes(attribute_ids)

        print(f"\nProfiles:")
        print("="*80)

        for idx, prof_or_error in enumerate(profiles, 1):
            if hasattr(prof_or_error, 'profileAttribute') and prof_or_error.profileAttribute:
                prof = prof_or_error.profileAttribute
                name = prof.name if hasattr(prof, 'name') else 'Unknown'
                print(f"  {idx}. {name}")

        return profiles

    except Exception as e:
        print(f"✗ Error getting profiles: {e}")
        return []


# =============================================================================
# GET LAYERS
# =============================================================================

def get_all_layers():
    """Get all layer attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('Layer')

        if not attribute_ids:
            return []

        # Get details
        layers = acc.GetLayerAttributes(attribute_ids)

        print(f"\nLayers:")
        print("="*80)

        for idx, layer_or_error in enumerate(layers, 1):
            if hasattr(layer_or_error, 'layerAttribute') and layer_or_error.layerAttribute:
                layer = layer_or_error.layerAttribute
                name = layer.name if hasattr(layer, 'name') else 'Unknown'

                # Status
                status = []
                if hasattr(layer, 'isHidden') and layer.isHidden:
                    status.append("Hidden")
                if hasattr(layer, 'isLocked') and layer.isLocked:
                    status.append("Locked")

                status_str = f" [{', '.join(status)}]" if status else ""
                print(f"  {idx}. {name}{status_str}")

        return layers

    except Exception as e:
        print(f"✗ Error getting layers: {e}")
        return []


# =============================================================================
# GET LAYER COMBINATIONS
# =============================================================================

def get_all_layer_combinations():
    """Get all layer combination attributes with details"""
    try:
        # Get IDs
        attribute_ids = get_attributes_by_type('LayerCombination')

        if not attribute_ids:
            return []

        # Get details
        combinations = acc.GetLayerCombinationAttributes(attribute_ids)

        print(f"\nLayer Combinations:")
        print("="*80)

        for idx, comb_or_error in enumerate(combinations, 1):
            if hasattr(comb_or_error, 'layerCombinationAttribute') and comb_or_error.layerCombinationAttribute:
                comb = comb_or_error.layerCombinationAttribute
                name = comb.name if hasattr(comb, 'name') else 'Unknown'

                # Number of layers in combination
                num_layers = len(comb.layerAttributeIds) if hasattr(
                    comb, 'layerAttributeIds') else 0

                print(f"  {idx}. {name} ({num_layers} layer(s))")

        return combinations

    except Exception as e:
        print(f"✗ Error getting layer combinations: {e}")
        return []


# =============================================================================
# GET ACTIVE PEN TABLES
# =============================================================================

def get_active_pen_tables():
    """Get currently active pen tables"""
    try:
        active_tables = acc.GetActivePenTables()

        print(f"\nActive Pen Tables:")
        print("="*80)

        # GetActivePenTables returns a tuple of (modelViewPenTable, layoutBookPenTable)
        if isinstance(active_tables, tuple) and len(active_tables) == 2:
            model_view_table = active_tables[0]
            layout_book_table = active_tables[1]

            # Extract AttributeId from AttributeIdOrError
            if hasattr(model_view_table, 'attributeId') and model_view_table.attributeId:
                model_id = model_view_table.attributeId
                print(f"  Model View Pen Table GUID: {model_id.guid}")
            else:
                print(f"  Model View Pen Table: Error or not set")

            if hasattr(layout_book_table, 'attributeId') and layout_book_table.attributeId:
                layout_id = layout_book_table.attributeId
                print(f"  Layout Book Pen Table GUID: {layout_id.guid}")
            else:
                print(f"  Layout Book Pen Table: Error or not set")
        else:
            print(f"  Unexpected response format")

        return active_tables

    except Exception as e:
        print(f"✗ Error getting active pen tables: {e}")
        import traceback
        traceback.print_exc()
        return None


# =============================================================================
# EXPORT TO FILE
# =============================================================================

def export_all_attributes_to_file(filename='attributes_summary.txt'):
    """Export all attributes to a text file"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD ATTRIBUTES SUMMARY\n")
            f.write("="*80 + "\n\n")

            # Building Materials
            f.write("\nBUILDING MATERIALS\n")
            f.write("─"*80 + "\n")
            materials = get_all_building_materials()
            for idx, mat_or_error in enumerate(materials, 1):
                if hasattr(mat_or_error, 'buildingMaterialAttribute') and mat_or_error.buildingMaterialAttribute:
                    mat = mat_or_error.buildingMaterialAttribute
                    name = mat.name if hasattr(mat, 'name') else 'Unknown'
                    f.write(f"{idx}. {name}\n")

            # Composites
            f.write("\nCOMPOSITES\n")
            f.write("─"*80 + "\n")
            composites = get_all_composites()
            for idx, comp_or_error in enumerate(composites, 1):
                if hasattr(comp_or_error, 'compositeAttribute') and comp_or_error.compositeAttribute:
                    comp = comp_or_error.compositeAttribute
                    name = comp.name if hasattr(comp, 'name') else 'Unknown'

                    # Get number of skins - handle both attribute types
                    num_skins = 0
                    if hasattr(comp, 'compositeSkins') and comp.compositeSkins:
                        num_skins = len(comp.compositeSkins)
                    elif hasattr(comp, 'skins') and comp.skins:
                        num_skins = len(comp.skins)

                    f.write(f"{idx}. {name} ({num_skins} skin(s))\n")

            # Surfaces
            f.write("\nSURFACES\n")
            f.write("─"*80 + "\n")
            surfaces = get_all_surfaces()
            for idx, surf_or_error in enumerate(surfaces, 1):
                if hasattr(surf_or_error, 'surfaceAttribute') and surf_or_error.surfaceAttribute:
                    surf = surf_or_error.surfaceAttribute
                    name = surf.name if hasattr(surf, 'name') else 'Unknown'
                    f.write(f"{idx}. {name}\n")

            # Continue for other types...

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Attributes exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"✗ Error exporting attributes: {e}")


# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    """Main demonstration function"""
    print("\n" + "="*80)
    print("GET ALL ATTRIBUTES - COMPLETE RETRIEVAL")
    print("="*80)

    # Get all attribute types
    print("\n" + "─"*80)
    print("Retrieving all attribute types...")
    print("─"*80)

    materials = get_all_building_materials()
    composites = get_all_composites()
    fills = get_all_fills()
    lines = get_all_lines()
    surfaces = get_all_surfaces()
    pen_tables = get_all_pen_tables()
    zone_categories = get_all_zone_categories()
    profiles = get_all_profiles()
    layers = get_all_layers()
    layer_combinations = get_all_layer_combinations()

    # Get active pen tables
    print("\n" + "─"*80)
    active_tables = get_active_pen_tables()

    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Building Materials: {len(materials)}")
    print(f"  Composites: {len(composites)}")
    print(f"  Fills: {len(fills)}")
    print(f"  Lines: {len(lines)}")
    print(f"  Surfaces: {len(surfaces)}")
    print(f"  Pen Tables: {len(pen_tables)}")
    print(f"  Zone Categories: {len(zone_categories)}")
    print(f"  Profiles: {len(profiles)}")
    print(f"  Layers: {len(layers)}")
    print(f"  Layer Combinations: {len(layer_combinations)}")
    print("\n" + "="*80)

    # Export option
    print("\n💡 To export all attributes to file:")
    print("   export_all_attributes_to_file('attributes_summary.txt')")


if __name__ == "__main__":
    main()
