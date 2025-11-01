"""
================================================================================
SCRIPT: Visualize Line Patterns
================================================================================

Version:        1.0
Date:           2025-10-29
Purpose:        Visualize Archicad line type patterns using matplotlib

This script creates visual representations of line patterns showing:
- Dash patterns (dash/gap sequences)
- Symbol patterns (LineItem shapes)
- Multiple repetitions to show the pattern effect

Usage: 
    python visualize_lines.py "Line Name"
    or modify the script to visualize specific lines

================================================================================
"""

from archicad import ACConnection
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import Arc
import numpy as np

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


def visualize_line_pattern(line_name, num_repetitions=3, output_file=None):
    """
    Visualize a line pattern
    
    Args:
        line_name: Name of the line to visualize (case insensitive search)
        num_repetitions: Number of times to repeat the pattern
        output_file: Optional output file path (PNG)
    """
    # Get all lines
    attribute_ids = acc.GetAttributesByType('Line')
    lines = acc.GetLineAttributes(attribute_ids)
    
    # Find the requested line
    target_line = None
    for line_or_error in lines:
        if hasattr(line_or_error, 'lineAttribute') and line_or_error.lineAttribute:
            line = line_or_error.lineAttribute
            name = line.name if hasattr(line, 'name') else ''
            if line_name.lower() in name.lower():
                target_line = line
                print(f"Found line: {name}")
                break
    
    if not target_line:
        print(f"Line '{line_name}' not found!")
        return
    
    # Get line properties
    line_type = target_line.lineType if hasattr(target_line, 'lineType') else 'Unknown'
    period = target_line.period if hasattr(target_line, 'period') else 10.0
    height = target_line.height if hasattr(target_line, 'height') else 2.0
    
    print(f"Line Type: {line_type}")
    print(f"Period: {period:.2f} mm")
    print(f"Height: {height:.2f} mm")
    
    # Create figure
    fig, ax = plt.subplots(figsize=(14, 4))
    
    # Set up axes
    total_length = period * num_repetitions
    ax.set_xlim(-5, total_length + 5)
    ax.set_ylim(-height * 2, height * 2)
    ax.set_aspect('equal')
    ax.grid(True, alpha=0.3)
    ax.axhline(y=0, color='gray', linestyle='--', linewidth=0.5, alpha=0.5)
    
    # Title
    ax.set_title(f"Line Pattern: {target_line.name}\n{line_type}, Period: {period:.1f} mm, Height: {height:.1f} mm", 
                 fontsize=12, fontweight='bold')
    ax.set_xlabel('Length (mm)')
    ax.set_ylabel('Height (mm)')
    
    # Draw the pattern
    if hasattr(target_line, 'lineItems') and target_line.lineItems:
        num_items = len(target_line.lineItems)
        print(f"Drawing {num_items} items per period, {num_repetitions} repetitions...")
        
        for rep in range(num_repetitions):
            x_offset = rep * period
            
            # Draw base line for this repetition
            ax.plot([x_offset, x_offset + period], [0, 0], 
                   color='lightgray', linewidth=0.5, alpha=0.5, zorder=1)
            
            for item_idx, dash_or_line_item in enumerate(target_line.lineItems):
                
                # Handle DashItem
                if hasattr(dash_or_line_item, 'dashItem') and dash_or_line_item.dashItem:
                    dash = dash_or_line_item.dashItem
                    dash_length = dash.dash if hasattr(dash, 'dash') else 0
                    gap_length = dash.gap if hasattr(dash, 'gap') else 0
                    
                    # Draw dash on the baseline
                    ax.plot([x_offset, x_offset + dash_length], [0, 0], 
                           color='black', linewidth=2, solid_capstyle='butt', zorder=2)
                    x_offset += dash_length + gap_length
                
                # Handle LineItem
                elif hasattr(dash_or_line_item, 'lineItem') and dash_or_line_item.lineItem:
                    line_item = dash_or_line_item.lineItem
                    line_item_type = line_item.lineItemType if hasattr(line_item, 'lineItemType') else 'Unknown'
                    
                    # Get common properties
                    center_offset = line_item.centerOffset if hasattr(line_item, 'centerOffset') else 0
                    length = line_item.length if hasattr(line_item, 'length') else 0
                    radius = line_item.radius if hasattr(line_item, 'radius') else 0
                    
                    # Get positions
                    beg_pos = None
                    end_pos = None
                    if hasattr(line_item, 'begPosition'):
                        bp = line_item.begPosition
                        if hasattr(bp, 'x') and hasattr(bp, 'y'):
                            beg_pos = (bp.x, bp.y)
                    if hasattr(line_item, 'endPosition'):
                        ep = line_item.endPosition
                        if hasattr(ep, 'x') and hasattr(ep, 'y'):
                            end_pos = (ep.x, ep.y)
                    
                    # Draw based on type
                    if line_item_type == 'CenterLineItemType':
                        # Vertical line at center_offset
                        y_start = -height / 2
                        y_end = height / 2
                        ax.plot([x_offset + center_offset, x_offset + center_offset], 
                               [y_start, y_end], 
                               color='blue', linewidth=1.5, zorder=2)
                    
                    elif line_item_type == 'ArcItemType':
                        # Draw arc
                        if radius > 0:
                            beg_angle = line_item.begAngle if hasattr(line_item, 'begAngle') else 0
                            end_angle = line_item.endAngle if hasattr(line_item, 'endAngle') else 0
                            
                            # Center of arc at x_offset + center_offset
                            arc = Arc((x_offset + center_offset, 0), 
                                     radius * 2, radius * 2,
                                     angle=0, 
                                     theta1=beg_angle, 
                                     theta2=end_angle,
                                     color='green', linewidth=1.5, zorder=2)
                            ax.add_patch(arc)
                    
                    elif line_item_type == 'CircleItemType':
                        # Draw circle
                        if radius > 0:
                            circle = plt.Circle((x_offset + center_offset, 0), 
                                              radius, 
                                              fill=False, 
                                              color='red', 
                                              linewidth=1.5, 
                                              zorder=2)
                            ax.add_patch(circle)
                    
                    elif line_item_type == 'LineItemType' and beg_pos and end_pos:
                        # Draw line from begPosition to endPosition
                        ax.plot([x_offset + beg_pos[0], x_offset + end_pos[0]], 
                               [beg_pos[1], end_pos[1]], 
                               color='purple', linewidth=1.5, zorder=2)
                    
                    elif line_item_type == 'SeparatorLineItemType':
                        # Short vertical line
                        y_span = height / 4
                        ax.plot([x_offset + center_offset, x_offset + center_offset], 
                               [-y_span, y_span], 
                               color='orange', linewidth=1.5, zorder=2)
                    
                    elif line_item_type == 'CenterDotItemType':
                        # Dot at center
                        ax.plot(x_offset + center_offset, 0, 
                               'ko', markersize=4, zorder=2)
    
    else:
        # Solid line or no items
        for rep in range(num_repetitions):
            x_offset = rep * period
            ax.plot([x_offset, x_offset + period], [0, 0], 
                   color='black', linewidth=2, solid_capstyle='butt', zorder=2)
    
    # Add period markers
    for rep in range(num_repetitions + 1):
        x = rep * period
        ax.axvline(x=x, color='red', linestyle=':', linewidth=1, alpha=0.3)
        if rep < num_repetitions:
            ax.text(x + period/2, -height * 1.5, f'Period {rep+1}', 
                   ha='center', va='top', fontsize=9, color='red', alpha=0.7)
    
    plt.tight_layout()
    
    # Save or show
    if output_file:
        plt.savefig(output_file, dpi=150, bbox_inches='tight')
        print(f"\nVisualization saved to: {output_file}")
    else:
        plt.savefig('/mnt/user-data/outputs/line_pattern.png', dpi=150, bbox_inches='tight')
        print(f"\nVisualization saved to: /mnt/user-data/outputs/line_pattern.png")
    
    plt.close()


def visualize_all_lines(output_dir='/mnt/user-data/outputs/line_patterns'):
    """Visualize all line types in the project"""
    import os
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Get all lines
    attribute_ids = acc.GetAttributesByType('Line')
    lines = acc.GetLineAttributes(attribute_ids)
    
    print(f"\nGenerating visualizations for all lines...")
    count = 0
    
    for line_or_error in lines:
        if hasattr(line_or_error, 'lineAttribute') and line_or_error.lineAttribute:
            line = line_or_error.lineAttribute
            name = line.name if hasattr(line, 'name') else 'Unknown'
            
            # Skip solid lines
            line_type = line.lineType if hasattr(line, 'lineType') else 'Unknown'
            if line_type == 'SolidLine':
                continue
            
            # Create safe filename
            safe_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in name)
            safe_name = safe_name.replace(' ', '_')
            output_file = os.path.join(output_dir, f"{safe_name}.png")
            
            try:
                visualize_line_pattern(name, num_repetitions=3, output_file=output_file)
                count += 1
            except Exception as e:
                print(f"Error visualizing {name}: {e}")
    
    print(f"\n✓ Generated {count} visualizations in {output_dir}")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        # Visualize specific line
        line_name = sys.argv[1]
        visualize_line_pattern(line_name, num_repetitions=3)
    else:
        # Example usage
        print("Visualizing example lines...")
        
        # Visualize specific lines
        visualize_line_pattern("Télécom", num_repetitions=3, 
                              output_file='/mnt/user-data/outputs/telecom_pattern.png')
        
        visualize_line_pattern("Terrain", num_repetitions=2, 
                              output_file='/mnt/user-data/outputs/terrain_pattern.png')
        
        # Or visualize all lines
        # visualize_all_lines()
        
        print("\n💡 To visualize a specific line:")
        print("   python visualize_lines.py 'Line Name'")
        print("\n💡 To visualize all lines:")
        print("   Uncomment the visualize_all_lines() call in the script")
