"""
================================================================================
SCRIPT: Get Fills - Enhanced Version
================================================================================

Version:        1.0
Date:           2025-10-29
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Fill (hatch pattern) attributes in the current 
Archicad project with complete details and folder structure organization.

Fills are hatch patterns and solid fills used to represent materials in 2D views.
This enhanced script provides:
- List of all fill patterns with detailed properties
- Fill subtype (filling type)
- Pattern matrix (8x8 pattern as 64-bit integer)
- Appearance type
- Folder hierarchy organization with fills listed
- GUID and index mapping
- Optional export to text file with all details

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
  
- GetAttributesByType('Fill')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetFillAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetFillAttributes

- GetAttributeFolderStructure('Fill')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

- GetAttributesIndices() - Archicad 28+
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesIndices

[Data Types]
- FillAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.FillAttribute

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View fills list with complete details and folder structure

Optional:
4. Call export_to_file_enhanced(enhanced_data, 'filename.txt') to save results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of fills
- For each fill:
  * Name
  * GUID
  * SubType (filling type)
  * Pattern (64-bit integer representing 8x8 matrix)
  * Pattern visualization (binary matrix)
  * Appearance Type
- Folder structure hierarchy with fills listed
- GUID/Index mapping (Archicad 28+)
- Summary statistics

Optional export creates a detailed text file with all information.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Fills represent materials and patterns in 2D views
- Used in sections, elevations, and floor plans
- Organized in folders for better management
- FillAttribute properties (from official API):
  * attributeId: The identifier of an attribute
  * name: The name of an attribute
  * subType: The filling type of a fill attribute
  * pattern: 64-bit unsigned integer representing an 8x8 matrix
             Each byte is a row, bits are columns (1=filled, 0=empty)
  * appearanceType: The appearance type of the fill attribute
- Pattern interpretation:
  * The pattern is stored as a 64-bit integer
  * Each byte (8 bits) represents one row of the 8x8 matrix
  * Bit value 1 = filled pixel, 0 = empty pixel
  * Example: 0xFF = 11111111 (fully filled row)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


# =============================================================================
# GET FILLS WITH ENHANCED DETAILS
# =============================================================================

def get_all_fills_enhanced():
    """Get all fill attributes with complete details"""
    try:
        # Get IDs
        print("\n" + "="*80)
        print("FILLS - ENHANCED DETAILS")
        print("="*80)

        attribute_ids = acc.GetAttributesByType('Fill')
        print(f"\n✓ Found {len(attribute_ids)} fill(s)")

        if not attribute_ids:
            print("\n⚠️  No fills found in the project")
            return [], [], []

        # Get details
        fills = acc.GetFillAttributes(attribute_ids)

        print(f"\nFills List:")
        print("-"*80)

        # Store enhanced data
        enhanced_fills_data = []

        for idx, fill_or_error in enumerate(fills, 1):
            if hasattr(fill_or_error, 'fillAttribute') and fill_or_error.fillAttribute:
                fill = fill_or_error.fillAttribute

                # Basic info
                name = fill.name if hasattr(fill, 'name') else 'Unknown'

                print(f"\n  {idx:3}. {name}")

                # Store data for export
                fill_data = {
                    'index': idx,
                    'name': name
                }

                # GUID
                if hasattr(fill, 'attributeId') and fill.attributeId:
                    guid = fill.attributeId.guid
                    print(f"       GUID: {guid}")
                    fill_data['guid'] = guid

                # SubType (filling type)
                if hasattr(fill, 'subType'):
                    subtype = fill.subType
                    print(f"       SubType: {subtype}")
                    fill_data['subType'] = subtype

                # Pattern (64-bit integer representing 8x8 matrix)
                if hasattr(fill, 'pattern'):
                    pattern = fill.pattern
                    print(f"       Pattern: {pattern}")
                    fill_data['pattern'] = pattern

                    # Decode pattern as 8x8 matrix for visualization
                    print(f"       Pattern Matrix (8x8):")
                    pattern_matrix = []
                    for row in range(8):
                        # Extract byte for this row (each byte is one row)
                        byte_value = (pattern >> (row * 8)) & 0xFF
                        row_bits = []
                        for col in range(8):
                            # Check if bit is set (1 = filled, 0 = empty)
                            bit = (byte_value >> col) & 1
                            row_bits.append('█' if bit else '·')
                        pattern_matrix.append(row_bits)
                        print(f"          {''.join(row_bits)}")

                    fill_data['pattern_matrix'] = pattern_matrix

                # Appearance Type
                if hasattr(fill, 'appearanceType'):
                    appearance = fill.appearanceType
                    print(f"       Appearance Type: {appearance}")
                    fill_data['appearanceType'] = appearance

                enhanced_fills_data.append(fill_data)

        return fills, attribute_ids, enhanced_fills_data

    except Exception as e:
        print(f"\n✗ Error getting fills: {e}")
        import traceback
        traceback.print_exc()
        return [], [], []


# =============================================================================
# GET FOLDER STRUCTURE
# =============================================================================

def get_folder_structure(attribute_type='Fill', path=None, indent=0):
    """Display the folder structure for fills with fill names"""
    try:
        structure = acc.GetAttributeFolderStructure(attribute_type, path)

        prefix = "  " * indent
        folder_name = structure.name if hasattr(structure, 'name') else 'Root'

        num_attributes = 0
        if hasattr(structure, 'attributes') and structure.attributes is not None:
            num_attributes = len(structure.attributes)

        num_subfolders = 0
        if hasattr(structure, 'subfolders') and structure.subfolders is not None:
            num_subfolders = len(structure.subfolders)

        print(
            f"{prefix}📁 {folder_name} ({num_attributes} fill(s), {num_subfolders} subfolder(s))")

        # Display fills in this folder
        if hasattr(structure, 'attributes') and structure.attributes is not None and len(structure.attributes) > 0:
            for attr_item in structure.attributes:
                if hasattr(attr_item, 'attributeId'):
                    try:
                        # Get fill details
                        fills_result = acc.GetFillAttributes(
                            [attr_item.attributeId])
                        if fills_result and len(fills_result) > 0:
                            result = fills_result[0]
                            if hasattr(result, 'fillAttribute') and result.fillAttribute:
                                fill = result.fillAttribute
                                fill_name = fill.name if hasattr(
                                    fill, 'name') else 'Unknown'

                                # Get subtype
                                subtype_str = ""
                                if hasattr(fill, 'subType'):
                                    subtype_str = f" - {fill.subType}"

                                print(f"{prefix}   • {fill_name}{subtype_str}")

                    except Exception as e:
                        print(f"{prefix}   ✗ Error loading fills: {e}")

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
                print(
                    f"  {idx:3}. Index: {index_info.index:4} | GUID: {index_info.guid}")

        return indices

    except Exception as e:
        print(f"\n✗ Error getting indices: {e}")
        return []


# =============================================================================
# EXPORT TO FILE
# =============================================================================

def export_to_file_enhanced(enhanced_data, filename='fills_detailed.txt'):
    """Export fills with complete details to a text file"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD FILLS - DETAILED REPORT\n")
            f.write("="*80 + "\n\n")

            for fill_data in enhanced_data:
                f.write(f"{fill_data['index']}. {fill_data['name']}\n")

                if 'guid' in fill_data:
                    f.write(f"   GUID: {fill_data['guid']}\n")

                if 'subType' in fill_data:
                    f.write(f"   SubType: {fill_data['subType']}\n")

                if 'pattern' in fill_data:
                    f.write(f"   Pattern: {fill_data['pattern']}\n")

                    # Write pattern matrix
                    if 'pattern_matrix' in fill_data:
                        f.write(f"   Pattern Matrix (8x8):\n")
                        for row in fill_data['pattern_matrix']:
                            f.write(f"      {''.join(row)}\n")

                if 'appearanceType' in fill_data:
                    f.write(
                        f"   Appearance Type: {fill_data['appearanceType']}\n")

                f.write("\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Fills detailed report exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting: {e}")
        import traceback
        traceback.print_exc()


def export_to_file(fills, filename='fills.txt'):
    """Export fills to a text file (basic version)"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD FILLS\n")
            f.write("="*80 + "\n\n")

            for idx, fill_or_error in enumerate(fills, 1):
                if hasattr(fill_or_error, 'fillAttribute') and fill_or_error.fillAttribute:
                    fill = fill_or_error.fillAttribute
                    name = fill.name if hasattr(fill, 'name') else 'Unknown'

                    f.write(f"{idx}. {name}\n")

                    if hasattr(fill, 'attributeId') and fill.attributeId:
                        f.write(f"   GUID: {fill.attributeId.guid}\n")

                    if hasattr(fill, 'subType'):
                        f.write(f"   SubType: {fill.subType}\n")

                    f.write("\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Fills exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function"""
    print("\n" + "="*80)
    print("GET FILLS v1.0 - ENHANCED")
    print("="*80)

    # Get all fills with detailed information
    fills, attribute_ids, enhanced_data = get_all_fills_enhanced()

    # Show folder structure
    if fills:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('Fill')

    # Show GUID/Index mapping
    if attribute_ids:
        get_indices_mapping(attribute_ids)

    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Fills: {len(fills)}")

    # Count by subType
    if enhanced_data:
        subtype_counts = {}
        for fill_data in enhanced_data:
            if 'subType' in fill_data:
                subtype = fill_data['subType']
                subtype_counts[subtype] = subtype_counts.get(subtype, 0) + 1

        if subtype_counts:
            print("\n  Fills by SubType:")
            for subtype, count in sorted(subtype_counts.items()):
                print(f"    - {subtype}: {count}")

    print("\n" + "="*80)

    # Export option
    print("\n💡 To export detailed report to file:")
    print("   fills, ids, enhanced = get_all_fills_enhanced()")
    print("   export_to_file_enhanced(enhanced, 'my_fills_detailed.txt')")


if __name__ == "__main__":
    main()
