"""
================================================================================
SCRIPT: Get Lines
================================================================================

Version:        1.0
Date:           2025-10-29
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Line type attributes in the current Archicad project
with complete details and folder structure organization.

Line types define the appearance of lines (solid, dashed, dotted, etc.) used
throughout the project. This enhanced script provides:
- List of all line types with detailed properties
- Line type classification (Solid, Dashed, Symbol)
- Appearance type
- Display scale and period information
- Height for symbol lines
- Detailed dash/line items breakdown
- Folder hierarchy organization with lines listed
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
  
- GetAttributesByType('Line')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetLineAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetLineAttributes

- GetAttributeFolderStructure('Line')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

- GetAttributesIndices() - Archicad 28+
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesIndices

[Data Types]
- LineAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.LineAttribute

- DashOrLineItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.DashOrLineItem

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View line types list with complete details and folder structure

Optional:
4. Call export_to_file_enhanced(enhanced_data, 'filename.txt') to save results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of line types
- For each line type:
  * Name
  * GUID
  * Line Type (Solid, Dashed, Symbol)
  * Appearance Type
  * Display Scale
  * Period (for dashed/symbol lines)
  * Height (for symbol lines)
  * Line Items details:
    - For DashItem: dash length and gap length
    - For LineItem: type (separator, center dot, centerline, circle, arc, etc.)
                    centerOffset, length, radius, angles, positions (as applicable)
- Folder structure hierarchy with line types listed
- GUID/Index mapping (Archicad 28+)
- Summary statistics by line type category

Optional export creates a detailed text file with all information.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Line types control the appearance of lines in drawings
- Used for construction lines, dimension lines, and element edges
- Organized in folders for better management
- LineAttribute properties (from official API):
  * attributeId: The identifier of an attribute
  * name: The name of an attribute
  * appearanceType: The appearance type of a line or fill attribute
  * displayScale: The original scale of the line
  * period: The length of the dashed or symbol line's period
  * height: The height of the symbol line
  * lineType: The type of a line attribute (Solid, Dashed, Symbol)
  * lineItems: A list of DashOrLineItem (segments that make up the pattern)
- Line types:
  * Solid: Continuous line (no lineItems)
  * Dashed: Line with dash pattern (DashItem segments)
  * Symbol: Line with repeating symbols (LineItem segments)
- Period: Total length of one repetition of the pattern
- DashOrLineItem structure:
  * Can contain either dashItem (DashItem) or lineItem (LineItem)
  * DashItem: Has dash length and gap length
  * LineItem: Has lineItemType, centerOffset, length, radius, angles, positions
- LineItem types include: separator, center dot, centerline, circle, arc, 
  right angle, parallel, and custom shapes with begin/end positions

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


# =============================================================================
# GET LINES WITH ENHANCED DETAILS
# =============================================================================

def get_all_lines_enhanced():
    """Get all line type attributes with complete details"""
    try:
        # Get IDs
        print("\n" + "="*80)
        print("LINE TYPES - ENHANCED DETAILS")
        print("="*80)

        attribute_ids = acc.GetAttributesByType('Line')
        print(f"\n✓ Found {len(attribute_ids)} line type(s)")

        if not attribute_ids:
            print("\n⚠️  No line types found in the project")
            return [], [], []

        # Get details
        lines = acc.GetLineAttributes(attribute_ids)

        print(f"\nLine Types List:")
        print("-"*80)

        # Store enhanced data
        enhanced_lines_data = []

        for idx, line_or_error in enumerate(lines, 1):
            if hasattr(line_or_error, 'lineAttribute') and line_or_error.lineAttribute:
                line = line_or_error.lineAttribute

                # Basic info
                name = line.name if hasattr(line, 'name') else 'Unknown'

                print(f"\n  {idx:3}. {name}")

                # Store data for export
                line_data = {
                    'index': idx,
                    'name': name
                }

                # GUID
                if hasattr(line, 'attributeId') and line.attributeId:
                    guid = line.attributeId.guid
                    print(f"       GUID: {guid}")
                    line_data['guid'] = guid

                # Line Type (Solid, Dashed, Symbol)
                if hasattr(line, 'lineType'):
                    line_type = line.lineType
                    print(f"       Line Type: {line_type}")
                    line_data['lineType'] = line_type

                # Appearance Type
                if hasattr(line, 'appearanceType'):
                    appearance = line.appearanceType
                    print(f"       Appearance Type: {appearance}")
                    line_data['appearanceType'] = appearance

                # Display Scale
                if hasattr(line, 'displayScale'):
                    scale = line.displayScale
                    print(f"       Display Scale: {scale:.4f}")
                    line_data['displayScale'] = scale

                # Period (length of one pattern repetition)
                if hasattr(line, 'period'):
                    period = line.period
                    print(f"       Period: {period:.4f} mm")
                    line_data['period'] = period

                # Height (for symbol lines)
                if hasattr(line, 'height'):
                    height = line.height
                    if height > 0:
                        print(f"       Height: {height:.4f} mm")
                        line_data['height'] = height

                # Line Items (dash/dot/symbol segments)
                if hasattr(line, 'lineItems') and line.lineItems:
                    num_items = len(line.lineItems)
                    print(f"       Line Items: {num_items} segment(s)")
                    line_data['num_items'] = num_items

                    # Store line items details
                    items_list = []
                    for item_idx, dash_or_line_item in enumerate(line.lineItems, 1):
                        item_info = {'index': item_idx}

                        # Check if it's a DashItem
                        if hasattr(dash_or_line_item, 'dashItem') and dash_or_line_item.dashItem:
                            dash = dash_or_line_item.dashItem
                            item_type = "Dash"

                            dash_length = dash.dash if hasattr(
                                dash, 'dash') else 0.0
                            gap_length = dash.gap if hasattr(
                                dash, 'gap') else 0.0

                            print(
                                f"          {item_idx}. Dash: {dash_length:.2f} mm, Gap: {gap_length:.2f} mm")

                            item_info['type'] = item_type
                            item_info['dash'] = dash_length
                            item_info['gap'] = gap_length

                        # Check if it's a LineItem
                        elif hasattr(dash_or_line_item, 'lineItem') and dash_or_line_item.lineItem:
                            line_item = dash_or_line_item.lineItem

                            # Get lineItemType
                            line_item_type = line_item.lineItemType if hasattr(
                                line_item, 'lineItemType') else 'Unknown'
                            item_info['type'] = f"Line-{line_item_type}"

                            # Build display string based on type
                            display_parts = [f"{item_idx}. {line_item_type}"]

                            # centerOffset
                            if hasattr(line_item, 'centerOffset'):
                                center_offset = line_item.centerOffset
                                item_info['centerOffset'] = center_offset
                                display_parts.append(
                                    f"offset: {center_offset:.2f} mm")

                            # length (for centerline, right angle, parallel)
                            if hasattr(line_item, 'length'):
                                length = line_item.length
                                item_info['length'] = length
                                if length > 0:
                                    display_parts.append(
                                        f"length: {length:.2f} mm")

                            # radius (for circle, arc)
                            if hasattr(line_item, 'radius'):
                                radius = line_item.radius
                                item_info['radius'] = radius
                                if radius > 0:
                                    display_parts.append(
                                        f"radius: {radius:.2f} mm")

                            # angles (for arc)
                            if hasattr(line_item, 'begAngle') and hasattr(line_item, 'endAngle'):
                                beg_angle = line_item.begAngle
                                end_angle = line_item.endAngle
                                item_info['begAngle'] = beg_angle
                                item_info['endAngle'] = end_angle
                                if beg_angle != 0 or end_angle != 0:
                                    display_parts.append(
                                        f"angles: {beg_angle:.2f}° to {end_angle:.2f}°")

                            # positions (for custom shapes) - ALWAYS SHOW if present
                            has_positions = False
                            if hasattr(line_item, 'begPosition'):
                                beg_pos = line_item.begPosition
                                if hasattr(beg_pos, 'x') and hasattr(beg_pos, 'y'):
                                    item_info['begPosition'] = {
                                        'x': beg_pos.x, 'y': beg_pos.y}
                                    # Only show if not at origin or if LineItemType is generic
                                    if beg_pos.x != 0 or beg_pos.y != 0 or line_item_type == 'LineItemType':
                                        display_parts.append(
                                            f"from: ({beg_pos.x:.2f}, {beg_pos.y:.2f})")
                                        has_positions = True

                            if hasattr(line_item, 'endPosition'):
                                end_pos = line_item.endPosition
                                if hasattr(end_pos, 'x') and hasattr(end_pos, 'y'):
                                    item_info['endPosition'] = {
                                        'x': end_pos.x, 'y': end_pos.y}
                                    # Only show if not at origin or if LineItemType is generic
                                    if end_pos.x != 0 or end_pos.y != 0 or line_item_type == 'LineItemType':
                                        display_parts.append(
                                            f"to: ({end_pos.x:.2f}, {end_pos.y:.2f})")
                                        has_positions = True

                            print(f"          {', '.join(display_parts)}")

                        else:
                            # Fallback if neither dashItem nor lineItem
                            item_info['type'] = 'Unknown'
                            print(
                                f"          {item_idx}. Unknown item structure")

                        items_list.append(item_info)

                    line_data['items'] = items_list

                enhanced_lines_data.append(line_data)

        return lines, attribute_ids, enhanced_lines_data

    except Exception as e:
        print(f"\n✗ Error getting line types: {e}")
        import traceback
        traceback.print_exc()
        return [], [], []


# =============================================================================
# GET FOLDER STRUCTURE
# =============================================================================

def get_folder_structure(attribute_type='Line', path=None, indent=0):
    """Display the folder structure for line types with line names"""
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
            f"{prefix}📁 {folder_name} ({num_attributes} line type(s), {num_subfolders} subfolder(s))")

        # Display line types in this folder
        if hasattr(structure, 'attributes') and structure.attributes is not None and len(structure.attributes) > 0:
            for attr_item in structure.attributes:
                if hasattr(attr_item, 'attributeId'):
                    try:
                        # Get line details
                        lines_result = acc.GetLineAttributes(
                            [attr_item.attributeId])
                        if lines_result and len(lines_result) > 0:
                            result = lines_result[0]
                            if hasattr(result, 'lineAttribute') and result.lineAttribute:
                                line = result.lineAttribute
                                line_name = line.name if hasattr(
                                    line, 'name') else 'Unknown'

                                # Get line type
                                line_type_str = ""
                                if hasattr(line, 'lineType'):
                                    line_type_str = f" - {line.lineType}"

                                print(
                                    f"{prefix}   • {line_name}{line_type_str}")

                    except Exception as e:
                        print(f"{prefix}   ✗ Error loading line types: {e}")

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

def export_to_file_enhanced(enhanced_data, filename='lines_detailed.txt'):
    """Export line types with complete details to a text file"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD LINE TYPES - DETAILED REPORT\n")
            f.write("="*80 + "\n\n")

            for line_data in enhanced_data:
                f.write(f"{line_data['index']}. {line_data['name']}\n")

                if 'guid' in line_data:
                    f.write(f"   GUID: {line_data['guid']}\n")

                if 'lineType' in line_data:
                    f.write(f"   Line Type: {line_data['lineType']}\n")

                if 'appearanceType' in line_data:
                    f.write(
                        f"   Appearance Type: {line_data['appearanceType']}\n")

                if 'displayScale' in line_data:
                    f.write(
                        f"   Display Scale: {line_data['displayScale']:.4f}\n")

                if 'period' in line_data:
                    f.write(f"   Period: {line_data['period']:.4f} mm\n")

                if 'height' in line_data:
                    f.write(f"   Height: {line_data['height']:.4f} mm\n")

                if 'num_items' in line_data:
                    f.write(
                        f"   Line Items: {line_data['num_items']} segment(s)\n")

                    if 'items' in line_data:
                        for item in line_data['items']:
                            if item['type'] == 'Dash':
                                f.write(
                                    f"      {item['index']}. Dash: {item['dash']:.2f} mm, Gap: {item['gap']:.2f} mm\n")
                            elif item['type'].startswith('Line-'):
                                # LineItem
                                line_item_type = item['type'].replace(
                                    'Line-', '')
                                parts = [f"{item['index']}. {line_item_type}"]

                                if 'centerOffset' in item:
                                    parts.append(
                                        f"offset: {item['centerOffset']:.2f} mm")
                                if 'length' in item and item['length'] > 0:
                                    parts.append(
                                        f"length: {item['length']:.2f} mm")
                                if 'radius' in item and item['radius'] > 0:
                                    parts.append(
                                        f"radius: {item['radius']:.2f} mm")
                                if 'begAngle' in item and 'endAngle' in item:
                                    if item['begAngle'] != 0 or item['endAngle'] != 0:
                                        parts.append(
                                            f"angles: {item['begAngle']:.2f}° to {item['endAngle']:.2f}°")

                                # Add positions if present
                                if 'begPosition' in item:
                                    beg = item['begPosition']
                                    parts.append(
                                        f"from: ({beg['x']:.2f}, {beg['y']:.2f})")
                                if 'endPosition' in item:
                                    end = item['endPosition']
                                    parts.append(
                                        f"to: ({end['x']:.2f}, {end['y']:.2f})")

                                f.write(f"      {', '.join(parts)}\n")

                f.write("\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Line types detailed report exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting: {e}")
        import traceback
        traceback.print_exc()


def export_to_file(lines, filename='lines.txt'):
    """Export line types to a text file (basic version)"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD LINE TYPES\n")
            f.write("="*80 + "\n\n")

            for idx, line_or_error in enumerate(lines, 1):
                if hasattr(line_or_error, 'lineAttribute') and line_or_error.lineAttribute:
                    line = line_or_error.lineAttribute
                    name = line.name if hasattr(line, 'name') else 'Unknown'

                    f.write(f"{idx}. {name}\n")

                    if hasattr(line, 'attributeId') and line.attributeId:
                        f.write(f"   GUID: {line.attributeId.guid}\n")

                    if hasattr(line, 'lineType'):
                        f.write(f"   Line Type: {line.lineType}\n")

                    f.write("\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Line types exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function"""
    print("\n" + "="*80)
    print("GET LINE TYPES v1.0 - ENHANCED")
    print("="*80)

    # Get all line types with detailed information
    lines, attribute_ids, enhanced_data = get_all_lines_enhanced()

    # Show folder structure
    if lines:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('Line')

    # Show GUID/Index mapping
    if attribute_ids:
        get_indices_mapping(attribute_ids)

    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Line Types: {len(lines)}")

    # Count by line type
    if enhanced_data:
        type_counts = {}
        for line_data in enhanced_data:
            if 'lineType' in line_data:
                line_type = line_data['lineType']
                type_counts[line_type] = type_counts.get(line_type, 0) + 1

        if type_counts:
            print("\n  Line Types by Category:")
            for line_type, count in sorted(type_counts.items()):
                print(f"    - {line_type}: {count}")

    print("\n" + "="*80)

    # Export option
    print("\n💡 To export detailed report to file:")
    print("   lines, ids, enhanced = get_all_lines_enhanced()")
    print("   export_to_file_enhanced(enhanced, 'my_lines_detailed.txt')")


if __name__ == "__main__":
    main()
