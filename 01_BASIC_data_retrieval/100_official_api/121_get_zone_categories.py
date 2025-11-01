"""
================================================================================
SCRIPT: Get Zone Categories with Complete Details
================================================================================

Version:        1.0
Date:           2025-10-29
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes


--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Zone Category attributes in the current Archicad
project with complete details including category codes, stamps, colors, and
folder organization.

Zone Categories classify spaces for area calculations, scheduling, and space
programming. This script provides:
- List of all zone categories with complete properties
- Category codes for sorting and filtering
- Zone stamp information (name and GUIDs)
- Category colors (RGB values)
- Folder hierarchy organization
- Statistics on category usage
- Optional export to detailed text file

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
  
- GetAttributesByType('ZoneCategory')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetZoneCategoryAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetZoneCategoryAttributes

- GetAttributeFolderStructure('ZoneCategory')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

[Data Types]
- ZoneCategoryAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.Types.ZoneCategoryAttribute

- RGBColor
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.RGBColor

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View zone categories list with complete details

Optional:
4. Call export_to_file_enhanced('filename.txt') to save complete results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of zone categories
- For each zone category:
  * Name and GUID
  * Category Code (for sorting/filtering)
  * Stamp Name
  * Stamp Main GUID
  * Stamp Revision GUID
  * Category Color (RGB values)
- Folder structure with categories listed
- Statistics on stamps and colors used

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
Zone Categories:
- Classify spaces for different purposes (Residential, Office, Circulation, etc.)
- Used in area calculations and space programming
- Category codes are used for sorting in schedules
- Each category has an associated zone stamp (graphical marker)

Stamp Information:
- Stamp Name: The name of the zone stamp symbol
- Stamp Main GUID: Main identifier of the stamp
- Stamp Revision GUID: Revision identifier for version tracking

Category Colors:
- Used for visual differentiation in plans and schedules
- RGB values: 0-255 for each color channel (or 0.0-1.0 as float)

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 201_assign_zone_category.py (assign categories to zones)
- 301_generate_area_schedule.py (create area reports)
- 119_get_surfaces.py (surface material properties)

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

def format_rgb_color(rgb_color):
    """Format RGB color object to readable string"""
    try:
        if rgb_color is None:
            return "N/A"

        # Get RGB values
        r = 0
        g = 0
        b = 0

        if hasattr(rgb_color, 'red'):
            r = rgb_color.red
        if hasattr(rgb_color, 'green'):
            g = rgb_color.green
        if hasattr(rgb_color, 'blue'):
            b = rgb_color.blue

        # Convert from float (0.0-1.0) to int (0-255) if needed
        if isinstance(r, float):
            r = int(round(r * 255))
        if isinstance(g, float):
            g = int(round(g * 255))
        if isinstance(b, float):
            b = int(round(b * 255))

        return f"RGB({r:3d},{g:3d},{b:3d})"
    except Exception as e:
        return f"Error: {e}"


# =============================================================================
# GET ALL ZONE CATEGORIES WITH ENHANCED DETAILS
# =============================================================================

def get_all_zone_categories_enhanced():
    """Get all zone category attributes with complete details"""
    try:
        print("\n" + "="*80)
        print("ZONE CATEGORIES - COMPLETE DETAILS")
        print("="*80)

        attribute_ids = acc.GetAttributesByType('ZoneCategory')
        print(f"\n✓ Found {len(attribute_ids)} zone category(ies)")

        if not attribute_ids:
            print("\n⚠️  No zone categories found in the project")
            return [], [], []

        zone_categories = acc.GetZoneCategoryAttributes(attribute_ids)

        print(f"\nZone Categories List with Complete Details:")
        print("-"*80)

        enhanced_data = []

        for idx, zc_or_error in enumerate(zone_categories, 1):
            if hasattr(zc_or_error, 'zoneCategoryAttribute') and zc_or_error.zoneCategoryAttribute:
                zc = zc_or_error.zoneCategoryAttribute

                # Basic properties
                name = zc.name if hasattr(zc, 'name') else 'Unknown'
                guid = zc.attributeId.guid if hasattr(
                    zc, 'attributeId') and zc.attributeId else 'N/A'

                # Category code
                category_code = zc.categoryCode if hasattr(
                    zc, 'categoryCode') else 'N/A'

                # Stamp information
                stamp_name = zc.stampName if hasattr(
                    zc, 'stampName') else 'N/A'
                stamp_main_guid = str(zc.stampMainGuid) if hasattr(
                    zc, 'stampMainGuid') else 'N/A'
                stamp_revision_guid = str(zc.stampRevisionGuid) if hasattr(
                    zc, 'stampRevisionGuid') else 'N/A'

                # Color
                color = zc.color if hasattr(zc, 'color') else None
                color_str = format_rgb_color(color)

                # Display
                print(f"\n  {idx:3}. {name}")
                print(f"       GUID: {guid}")
                print(f"       Category Code: {category_code}")
                print(f"       Color: {color_str}")
                print(f"\n       Stamp Information:")
                print(f"         • Name: {stamp_name}")
                print(f"         • Main GUID: {stamp_main_guid}")
                print(f"         • Revision GUID: {stamp_revision_guid}")

                # Store enhanced data
                enhanced_data.append({
                    'index': idx,
                    'name': name,
                    'guid': guid,
                    'category_code': category_code,
                    'stamp_name': stamp_name,
                    'stamp_main_guid': stamp_main_guid,
                    'stamp_revision_guid': stamp_revision_guid,
                    'color': color_str
                })

        return zone_categories, attribute_ids, enhanced_data

    except Exception as e:
        print(f"\n✗ Error getting zone categories: {e}")
        import traceback
        traceback.print_exc()
        return [], [], []


# =============================================================================
# GET FOLDER STRUCTURE
# =============================================================================

def get_folder_structure(attribute_type='ZoneCategory', path=None, indent=0):
    """Display the folder structure for zone categories with category names"""
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
            f"{prefix}📁 {folder_name} ({num_attributes} category(ies), {num_subfolders} subfolder(s))")

        # Display categories in this folder
        if hasattr(structure, 'attributes') and structure.attributes is not None and len(structure.attributes) > 0:
            # Get category names for the attributes in this folder
            attribute_ids = [item.attributeId for item in structure.attributes
                             if hasattr(item, 'attributeId')]

            if attribute_ids:
                try:
                    categories_result = acc.GetZoneCategoryAttributes(
                        attribute_ids)

                    for cat_item in categories_result:
                        if hasattr(cat_item, 'zoneCategoryAttribute') and cat_item.zoneCategoryAttribute:
                            cat = cat_item.zoneCategoryAttribute
                            cat_name = cat.name if hasattr(
                                cat, 'name') else 'Unknown'
                            cat_code = cat.categoryCode if hasattr(
                                cat, 'categoryCode') else ''
                            print(f"{prefix}   • {cat_name} (Code: {cat_code})")

                except Exception as e:
                    print(f"{prefix}   ✗ Error loading categories: {e}")

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
# STATISTICS
# =============================================================================

def display_statistics(enhanced_data):
    """Display useful statistics about zone categories"""
    print("\n" + "="*80)
    print("ZONE CATEGORIES STATISTICS")
    print("="*80)

    # Unique stamps
    unique_stamps = {}
    for cat in enhanced_data:
        stamp_name = cat['stamp_name']
        if stamp_name not in unique_stamps:
            unique_stamps[stamp_name] = []
        unique_stamps[stamp_name].append(cat['name'])

    print(f"\nStamp Usage:")
    print("-"*80)
    print(f"  • Total unique stamps: {len(unique_stamps)}")

    if len(unique_stamps) > 0:
        print(f"\n  Stamp distribution:")
        for stamp_name, categories in sorted(unique_stamps.items(), key=lambda x: len(x[1]), reverse=True):
            print(f"    • {stamp_name}: {len(categories)} category(ies)")
            for cat_name in categories[:3]:
                print(f"      - {cat_name}")
            if len(categories) > 3:
                print(f"      ... and {len(categories) - 3} more")

    # Category codes analysis
    codes_with_values = [
        cat for cat in enhanced_data if cat['category_code'] and cat['category_code'] != 'N/A']

    print(f"\nCategory Codes:")
    print("-"*80)
    print(f"  • Categories with codes: {len(codes_with_values)}")
    print(
        f"  • Categories without codes: {len(enhanced_data) - len(codes_with_values)}")

    if codes_with_values:
        print(f"\n  Sample codes:")
        for cat in codes_with_values[:5]:
            print(f"    • {cat['category_code']}: {cat['name']}")
        if len(codes_with_values) > 5:
            print(f"    ... and {len(codes_with_values) - 5} more")


# =============================================================================
# EXPORT TO FILE - ENHANCED
# =============================================================================

def export_to_file_enhanced(enhanced_data, filename='zone_categories_detailed.txt'):
    """Export zone categories with complete details to a text file"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD ZONE CATEGORIES - DETAILED REPORT\n")
            f.write("Complete Category Properties\n")
            f.write("="*80 + "\n\n")

            for cat_data in enhanced_data:
                f.write(f"{cat_data['index']}. {cat_data['name']}\n")
                f.write(f"   GUID: {cat_data['guid']}\n")
                f.write(f"   Category Code: {cat_data['category_code']}\n")
                f.write(f"   Color: {cat_data['color']}\n")
                f.write(f"\n   Stamp Information:\n")
                f.write(f"     • Name: {cat_data['stamp_name']}\n")
                f.write(f"     • Main GUID: {cat_data['stamp_main_guid']}\n")
                f.write(
                    f"     • Revision GUID: {cat_data['stamp_revision_guid']}\n")
                f.write("\n" + "-"*80 + "\n\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Zone categories detailed report exported to:")
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
    print("GET ZONE CATEGORIES v1.0 - Complete Details")
    print("="*80)

    # Get all zone categories with enhanced details
    zone_categories, attribute_ids, enhanced_data = get_all_zone_categories_enhanced()

    # Show folder structure
    if zone_categories:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('ZoneCategory')

    # Show statistics
    if enhanced_data:
        display_statistics(enhanced_data)

    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Zone Categories: {len(zone_categories)}")
    print("\n" + "="*80)

    # Export option
    print("\n💡 To export detailed report to file:")
    print("   zone_categories, ids, enhanced = get_all_zone_categories_enhanced()")
    print("   export_to_file_enhanced(enhanced, 'my_zone_categories_detailed.txt')")


if __name__ == "__main__":
    main()
