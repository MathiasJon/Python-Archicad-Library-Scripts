"""
================================================================================
SCRIPT: Get Pen Tables with Complete Pen Details
================================================================================

Version:        1.0
Date:           2025-10-29
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes


--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Pen Table attributes in the current Archicad project
with complete details about each individual pen including colors, weights, and
descriptions.

Pen Tables define sets of pen colors and weights used for printing and display.
This script provides:
- List of all pen tables (typically containing 255 pens each)
- Complete details for each pen: index, RGB color, weight (mm), description
- Currently active pen tables (Model View and Layout Book)
- Pen weight distribution analysis
- Most commonly used pen weights
- Color analysis (grayscale vs color pens)
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
  
- GetAttributesByType('PenTable')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetPenTableAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetPenTableAttributes

- GetActivePenTables()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetActivePenTables

- GetAttributeFolderStructure('PenTable')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

[Data Types]
- PenTableAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.PenTableAttribute

- Pen
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.Pen

- RGBColor
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.RGBColor

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View pen tables list with complete pen details

Optional:
4. Call export_to_file_enhanced('filename.txt') to save complete results
5. Call export_pen_table_detailed('table_name.txt', pen_table_index) for single table

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of pen tables
- For each pen table:
  * Name and GUID
  * Number of pens (typically 255)
  * Active status (Model View / Layout Book)
  * Sample pens with details (first 10, last 5)
  * Each pen: Index, RGB Color, Weight (mm), Description
- Pen weight analysis:
  * Distribution of pen weights across all tables
  * Most common weights
  * Weight range (min to max)
- Pen color analysis:
  * Grayscale vs colored pens
  * Black pens count

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
Pen Properties:
- Index: Pen number (1-255 in standard tables)
- Color: RGB color values (0-255 for each channel)
- Weight: Line thickness in millimeters for printing
- Description: Text description of the pen's purpose

Pen Table Contexts:
- Model View: Used for 3D views and working drawings
- Layout Book: Used for published layouts and prints

Common Pen Weights:
- 0.13mm: Very fine lines (details, dimensions)
- 0.18mm: Fine lines (secondary elements)
- 0.25mm: Standard lines (walls, objects)
- 0.35mm: Medium lines (important elements)
- 0.50mm: Bold lines (main elements, cuts)
- 0.70mm: Very bold lines (emphasis)

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 114_04_get_lines.py (line types)
- 202_set_active_pen_table.py (change active pen table)
- 115_get_building_materials.py (materials using pens)

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


def is_grayscale(rgb_color):
    """Check if color is grayscale (R=G=B)"""
    try:
        if rgb_color is None:
            return False

        r = rgb_color.red if hasattr(rgb_color, 'red') else 0
        g = rgb_color.green if hasattr(rgb_color, 'green') else 0
        b = rgb_color.blue if hasattr(rgb_color, 'blue') else 0

        # Handle both float and int values
        if isinstance(r, float):
            r = round(r * 255)
        if isinstance(g, float):
            g = round(g * 255)
        if isinstance(b, float):
            b = round(b * 255)

        return r == g == b
    except:
        return False


def is_black(rgb_color):
    """Check if color is black (R=G=B=0)"""
    try:
        if rgb_color is None:
            return False

        r = rgb_color.red if hasattr(rgb_color, 'red') else 0
        g = rgb_color.green if hasattr(rgb_color, 'green') else 0
        b = rgb_color.blue if hasattr(rgb_color, 'blue') else 0

        # Handle both float and int values
        if isinstance(r, float):
            r = round(r * 255)
        if isinstance(g, float):
            g = round(g * 255)
        if isinstance(b, float):
            b = round(b * 255)

        return r == 0 and g == 0 and b == 0
    except:
        return False


def get_pen_details(pen):
    """Extract all details from a pen object"""
    try:
        details = {
            'index': pen.index if hasattr(pen, 'index') else 0,
            'color': pen.color if hasattr(pen, 'color') else None,
            'weight': pen.weight if hasattr(pen, 'weight') else 0.0,
            'description': pen.description if hasattr(pen, 'description') else ''
        }
        return details
    except:
        return None


# =============================================================================
# GET ACTIVE PEN TABLES
# =============================================================================

def get_active_pen_tables():
    """Get currently active pen tables and return their GUIDs"""
    try:
        active_tables = acc.GetActivePenTables()

        active_guids = {
            'model_view': None,
            'layout_book': None
        }

        # GetActivePenTables returns a tuple of (modelViewPenTable, layoutBookPenTable)
        if isinstance(active_tables, tuple) and len(active_tables) == 2:
            model_view_table = active_tables[0]
            layout_book_table = active_tables[1]

            # Extract AttributeId from AttributeIdOrError
            if hasattr(model_view_table, 'attributeId') and model_view_table.attributeId:
                active_guids['model_view'] = model_view_table.attributeId.guid

            if hasattr(layout_book_table, 'attributeId') and layout_book_table.attributeId:
                active_guids['layout_book'] = layout_book_table.attributeId.guid

        return active_guids

    except Exception as e:
        print(f"Warning: Could not get active pen tables: {e}")
        return {'model_view': None, 'layout_book': None}


# =============================================================================
# GET ALL PEN TABLES WITH ENHANCED DETAILS
# =============================================================================

def get_all_pen_tables_enhanced():
    """Get all pen table attributes with complete pen details"""
    try:
        print("\n" + "="*80)
        print("PEN TABLES - COMPLETE PEN DETAILS")
        print("="*80)

        attribute_ids = acc.GetAttributesByType('PenTable')
        print(f"\n✓ Found {len(attribute_ids)} pen table(s)")

        if not attribute_ids:
            print("\n⚠️  No pen tables found in the project")
            return [], [], []

        pen_tables = acc.GetPenTableAttributes(attribute_ids)

        # Get active pen tables
        active_guids = get_active_pen_tables()

        print(f"\nPen Tables with Sample Pens:")
        print("-"*80)

        enhanced_data = []

        for idx, pt_or_error in enumerate(pen_tables, 1):
            if hasattr(pt_or_error, 'penTableAttribute') and pt_or_error.penTableAttribute:
                pt = pt_or_error.penTableAttribute

                # Basic properties
                name = pt.name if hasattr(pt, 'name') else 'Unknown'
                guid = pt.attributeId.guid if hasattr(
                    pt, 'attributeId') and pt.attributeId else 'N/A'

                # Check if this table is active
                is_model_view_active = (guid == active_guids['model_view'])
                is_layout_book_active = (guid == active_guids['layout_book'])
                active_status = []
                if is_model_view_active:
                    active_status.append("Model View")
                if is_layout_book_active:
                    active_status.append("Layout Book")

                # Get pens
                pens = []
                if hasattr(pt, 'pens') and pt.pens:
                    for pen_item in pt.pens:
                        if hasattr(pen_item, 'pen'):
                            pen_details = get_pen_details(pen_item.pen)
                            if pen_details:
                                pens.append(pen_details)

                num_pens = len(pens)

                # Display pen table header
                print(f"\n  {idx}. {name}")
                print(f"     GUID: {guid}")
                print(f"     Pens: {num_pens}")
                if active_status:
                    print(f"     🔹 ACTIVE for: {', '.join(active_status)}")

                # Show sample pens (first 10 and last 5)
                if pens:
                    print(f"\n     Sample Pens (first 10):")
                    print(
                        f"     {'Index':>5} | {'Color':>15} | {'Weight':>8} | Description")
                    print(f"     {'-'*5}-+-{'-'*15}-+-{'-'*8}-+-{'-'*30}")

                    # First 10 pens
                    for pen in pens[:10]:
                        color_str = format_rgb_color(pen['color'])
                        weight_str = f"{pen['weight']:.3f}mm"
                        desc = pen['description'][:30] if pen['description'] else ''
                        print(
                            f"     {pen['index']:>5} | {color_str:>15} | {weight_str:>8} | {desc}")

                    if num_pens > 15:
                        print(
                            f"     {'...':>5} | {'...':>15} | {'...':>8} | ...")

                        # Last 5 pens
                        print(f"\n     Sample Pens (last 5):")
                        print(
                            f"     {'Index':>5} | {'Color':>15} | {'Weight':>8} | Description")
                        print(f"     {'-'*5}-+-{'-'*15}-+-{'-'*8}-+-{'-'*30}")
                        for pen in pens[-5:]:
                            color_str = format_rgb_color(pen['color'])
                            weight_str = f"{pen['weight']:.3f}mm"
                            desc = pen['description'][:30] if pen['description'] else ''
                            print(
                                f"     {pen['index']:>5} | {color_str:>15} | {weight_str:>8} | {desc}")

                # Store enhanced data
                enhanced_data.append({
                    'index': idx,
                    'name': name,
                    'guid': guid,
                    'num_pens': num_pens,
                    'pens': pens,
                    'is_model_view_active': is_model_view_active,
                    'is_layout_book_active': is_layout_book_active
                })

        return pen_tables, attribute_ids, enhanced_data

    except Exception as e:
        print(f"\n✗ Error getting pen tables: {e}")
        import traceback
        traceback.print_exc()
        return [], [], []


# =============================================================================
# PEN STATISTICS
# =============================================================================

def get_pen_statistics(enhanced_data):
    """Display useful statistics about pens across all tables"""
    print("\n" + "="*80)
    print("PEN STATISTICS - ANALYSIS ACROSS ALL TABLES")
    print("="*80)

    # Collect all pens from all tables
    all_pens = []
    for table_data in enhanced_data:
        all_pens.extend(table_data['pens'])

    if not all_pens:
        print("\nNo pens found to analyze")
        return

    # Weight analysis
    weights = [pen['weight'] for pen in all_pens]
    unique_weights = sorted(set(weights))
    weight_counts = {w: weights.count(w) for w in unique_weights}

    print(f"\nPen Weight Distribution:")
    print("-"*80)
    print(f"  • Total unique weights: {len(unique_weights)}")
    print(f"  • Weight range: {min(weights):.3f}mm to {max(weights):.3f}mm")
    print(f"\n  Most common pen weights:")

    # Sort by frequency
    sorted_weights = sorted(weight_counts.items(),
                            key=lambda x: x[1], reverse=True)
    for weight, count in sorted_weights[:10]:
        percentage = (count / len(all_pens)) * 100
        print(f"    • {weight:.3f}mm: {count} pens ({percentage:.1f}%)")

    # Color analysis
    grayscale_pens = [pen for pen in all_pens if is_grayscale(pen['color'])]
    black_pens = [pen for pen in all_pens if is_black(pen['color'])]
    colored_pens = [pen for pen in all_pens if not is_grayscale(pen['color'])]

    print(f"\nPen Color Analysis:")
    print("-"*80)
    print(f"  • Total pens analyzed: {len(all_pens)}")
    print(
        f"  • Grayscale pens: {len(grayscale_pens)} ({(len(grayscale_pens)/len(all_pens)*100):.1f}%)")
    print(f"    - Black pens (RGB 0,0,0): {len(black_pens)}")
    print(
        f"  • Colored pens: {len(colored_pens)} ({(len(colored_pens)/len(all_pens)*100):.1f}%)")

    # Pen descriptions analysis
    described_pens = [pen for pen in all_pens if pen['description']]

    print(f"\nPen Descriptions:")
    print("-"*80)
    print(
        f"  • Pens with descriptions: {len(described_pens)} ({(len(described_pens)/len(all_pens)*100):.1f}%)")
    print(
        f"  • Pens without descriptions: {len(all_pens) - len(described_pens)}")


# =============================================================================
# GET FOLDER STRUCTURE (simplified, typically no subfolders)
# =============================================================================

def get_folder_structure(attribute_type='PenTable', path=None):
    """Display the folder structure for pen tables (typically flat)"""
    try:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)

        structure = acc.GetAttributeFolderStructure(attribute_type, path)

        folder_name = structure.name if hasattr(structure, 'name') else 'Root'

        num_attributes = 0
        if hasattr(structure, 'attributes') and structure.attributes is not None:
            num_attributes = len(structure.attributes)

        num_subfolders = 0
        if hasattr(structure, 'subfolders') and structure.subfolders is not None:
            num_subfolders = len(structure.subfolders)

        print(f"\n📁 {folder_name}")
        print(f"   • Pen Tables: {num_attributes}")
        print(f"   • Subfolders: {num_subfolders}")

        if num_subfolders == 0:
            print(f"   ℹ️  Note: Pen tables typically have no subfolder organization")

        return structure

    except Exception as e:
        print(f"✗ Error: {e}")
        return None


# =============================================================================
# EXPORT TO FILE - ENHANCED
# =============================================================================

def export_to_file_enhanced(enhanced_data, filename='pen_tables_detailed.txt'):
    """Export pen tables with complete pen details to a text file"""
    try:
        import os

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD PEN TABLES - DETAILED REPORT\n")
            f.write("Complete Pen Details for All Tables\n")
            f.write("="*80 + "\n\n")

            for table_data in enhanced_data:
                f.write(f"{table_data['index']}. {table_data['name']}\n")
                f.write(f"   GUID: {table_data['guid']}\n")
                f.write(f"   Number of Pens: {table_data['num_pens']}\n")

                if table_data['is_model_view_active']:
                    f.write(f"   🔹 ACTIVE for Model View\n")
                if table_data['is_layout_book_active']:
                    f.write(f"   🔹 ACTIVE for Layout Book\n")

                f.write(f"\n   Complete Pen List:\n")
                f.write(
                    f"   {'Index':>5} | {'Color':>15} | {'Weight':>8} | Description\n")
                f.write(f"   {'-'*5}-+-{'-'*15}-+-{'-'*8}-+-{'-'*40}\n")

                for pen in table_data['pens']:
                    color_str = format_rgb_color(pen['color'])
                    weight_str = f"{pen['weight']:.3f}mm"
                    desc = pen['description'] if pen['description'] else ''
                    f.write(
                        f"   {pen['index']:>5} | {color_str:>15} | {weight_str:>8} | {desc}\n")

                f.write("\n" + "="*80 + "\n\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Pen tables detailed report exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting: {e}")
        import traceback
        traceback.print_exc()


def export_pen_table_detailed(table_name, pen_table_data, filename=None):
    """Export a single pen table with all pen details"""
    try:
        import os

        if filename is None:
            safe_name = "".join(c if c.isalnum() or c in (
                ' ', '-', '_') else '_' for c in table_name)
            filename = f"pen_table_{safe_name}.txt"

        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"ARCHICAD PEN TABLE: {table_name}\n")
            f.write("="*80 + "\n\n")
            f.write(f"GUID: {pen_table_data['guid']}\n")
            f.write(f"Total Pens: {pen_table_data['num_pens']}\n\n")

            f.write(
                f"{'Index':>5} | {'Red':>3} {'Green':>3} {'Blue':>3} | {'Weight':>8} | Description\n")
            f.write(f"{'-'*5}-+-{'-'*13}-+-{'-'*8}-+-{'-'*40}\n")

            for pen in pen_table_data['pens']:
                r = pen['color'].red if pen['color'] else 0
                g = pen['color'].green if pen['color'] else 0
                b = pen['color'].blue if pen['color'] else 0

                # Convert from float (0.0-1.0) to int (0-255) if needed
                if isinstance(r, float):
                    r = int(round(r * 255))
                if isinstance(g, float):
                    g = int(round(g * 255))
                if isinstance(b, float):
                    b = int(round(b * 255))

                weight_str = f"{pen['weight']:.3f}mm"
                desc = pen['description'] if pen['description'] else ''
                f.write(
                    f"{pen['index']:>5} | {r:>3} {g:>3} {b:>3} | {weight_str:>8} | {desc}\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Pen table exported to: {abs_path}")

    except Exception as e:
        print(f"\n✗ Error exporting pen table: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function"""
    print("\n" + "="*80)
    print("GET PEN TABLES v1.0 - Complete Pen Details")
    print("="*80)

    # Get all pen tables with enhanced details
    pen_tables, attribute_ids, enhanced_data = get_all_pen_tables_enhanced()

    # Show statistics
    if enhanced_data:
        get_pen_statistics(enhanced_data)

    # Show folder structure
    if pen_tables:
        get_folder_structure('PenTable')

    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Pen Tables: {len(pen_tables)}")
    if enhanced_data:
        total_pens = sum(table['num_pens'] for table in enhanced_data)
        active_model = sum(
            1 for t in enhanced_data if t['is_model_view_active'])
        active_layout = sum(
            1 for t in enhanced_data if t['is_layout_book_active'])
        print(f"  Total Pens (all tables): {total_pens}")
        print(f"  Active for Model View: {active_model} table(s)")
        print(f"  Active for Layout Book: {active_layout} table(s)")
    print("\n" + "="*80)

    # Export options
    print("\n💡 To export complete report to file:")
    print("   pen_tables, ids, enhanced = get_all_pen_tables_enhanced()")
    print("   export_to_file_enhanced(enhanced, 'my_pen_tables_detailed.txt')")
    print("\n💡 To export a single pen table:")
    print(
        "   export_pen_table_detailed('Table Name', enhanced[0], 'single_table.txt')")


if __name__ == "__main__":
    main()
