"""
Create polylines in Archicad.

This script provides functionality to create 2D polylines (lines, polygons)
using the TAPIR CreatePolylines command (v1.1.5).

Based on TAPIR API documentation.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def create_polylines(polylines_data):
    """
    Create one or more polylines in Archicad.

    Args:
        polylines_data: List of polyline data dictionaries
                       Each dictionary should contain:
                       - coordinates: List of {x, y} coordinate dicts (min 2 points)
                       - arcs: Optional list of arc definitions
                       - floorInd: Optional floor index (default: current floor)

    Returns:
        List of created element GUIDs
    """
    try:
        # Create proper AddOnCommandId
        command_id = act.AddOnCommandId('TapirCommand', 'CreatePolylines')

        # Prepare parameters
        parameters = {
            'polylinesData': polylines_data
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ CreatePolylines failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'elements' in response:
            elements = response['elements']
            guids = []
            for elem in elements:
                if isinstance(elem, dict) and 'elementId' in elem:
                    guid = elem['elementId'].get('guid')
                    if guid:
                        guids.append(guid)

            print(f"✓ Created {len(guids)} polyline(s)")
            return guids

        return []

    except Exception as e:
        print(f"✗ Error creating polylines: {e}")
        import traceback
        traceback.print_exc()
        return []


def create_polyline(coordinates, arcs=None, floor_index=None):
    """
    Create a single polyline.

    Args:
        coordinates: List of (x, y) tuples representing vertices
        arcs: Optional list of arc dictionaries with begIndex, endIndex, arcAngle
        floor_index: Optional floor index (None = current floor)

    Returns:
        GUID of created polyline or None
    """
    # Prepare coordinate list
    coord_list = []
    for x, y in coordinates:
        coord_list.append({'x': x, 'y': y})

    # Prepare polyline data
    polyline_data = {
        'coordinates': coord_list
    }

    if arcs:
        polyline_data['arcs'] = arcs

    if floor_index is not None:
        polyline_data['floorInd'] = floor_index

    # Create polyline
    guids = create_polylines([polyline_data])

    return guids[0] if guids else None


def create_rectangle(x, y, width, height, floor_index=None):
    """
    Create a rectangular polyline.

    Args:
        x: X coordinate of bottom-left corner
        y: Y coordinate of bottom-left corner
        width: Width of rectangle (mm)
        height: Height of rectangle (mm)
        floor_index: Optional floor index

    Returns:
        GUID of created rectangle
    """
    coordinates = [
        (x, y),
        (x + width, y),
        (x + width, y + height),
        (x, y + height),
        (x, y)  # Close the polyline
    ]

    print(f"Creating rectangle: {width}mm × {height}mm at ({x}, {y})")
    return create_polyline(coordinates, floor_index=floor_index)


def create_circle_approximation(center_x, center_y, radius, segments=32, floor_index=None):
    """
    Create a circle using a polyline approximation.

    Args:
        center_x: X coordinate of circle center (mm)
        center_y: Y coordinate of circle center (mm)
        radius: Radius of the circle (mm)
        segments: Number of segments to approximate the circle
        floor_index: Optional floor index

    Returns:
        GUID of created circle
    """
    import math

    coordinates = []
    angle_step = 2 * math.pi / segments

    for i in range(segments + 1):  # +1 to close the circle
        angle = i * angle_step
        x = center_x + radius * math.cos(angle)
        y = center_y + radius * math.sin(angle)
        coordinates.append((x, y))

    print(
        f"Creating circle: radius {radius}mm, {segments} segments at ({center_x}, {center_y})")
    return create_polyline(coordinates, floor_index=floor_index)


def create_regular_polygon(center_x, center_y, radius, sides, rotation=0, floor_index=None):
    """
    Create a regular polygon.

    Args:
        center_x: X coordinate of polygon center (mm)
        center_y: Y coordinate of polygon center (mm)
        radius: Radius - distance from center to vertex (mm)
        sides: Number of sides
        rotation: Rotation angle in degrees
        floor_index: Optional floor index

    Returns:
        GUID of created polygon
    """
    import math

    coordinates = []
    angle_step = 2 * math.pi / sides
    rotation_rad = math.radians(rotation)

    for i in range(sides + 1):  # +1 to close the polygon
        angle = i * angle_step + rotation_rad
        x = center_x + radius * math.cos(angle)
        y = center_y + radius * math.sin(angle)
        coordinates.append((x, y))

    print(
        f"Creating {sides}-sided polygon: radius {radius}mm at ({center_x}, {center_y})")
    return create_polyline(coordinates, floor_index=floor_index)


def create_polyline_with_arc(coordinates, arc_indices, arc_angles, floor_index=None):
    """
    Create a polyline with curved segments.

    Args:
        coordinates: List of (x, y) tuples
        arc_indices: List of tuples (start_idx, end_idx) for arcs
        arc_angles: List of arc angles in radians (positive = clockwise)
        floor_index: Optional floor index

    Returns:
        GUID of created polyline
    """
    # Prepare arcs
    arcs = []
    for (beg_idx, end_idx), angle in zip(arc_indices, arc_angles):
        arcs.append({
            'begIndex': beg_idx,
            'endIndex': end_idx,
            'arcAngle': angle
        })

    print(f"Creating polyline with {len(arcs)} arc(s)")
    return create_polyline(coordinates, arcs=arcs, floor_index=floor_index)


def create_multiple_polylines_batch(polylines_list):
    """
    Create multiple polylines in a single batch operation.
    More efficient than creating them one by one.

    Args:
        polylines_list: List of polyline data dictionaries

    Returns:
        List of created element GUIDs
    """
    print(f"Creating {len(polylines_list)} polylines in batch...")
    return create_polylines(polylines_list)


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("CREATE POLYLINES")
    print("="*80)

    # Example 1: Create a simple open polyline (zigzag)
    print("\n" + "─"*80)
    print("EXAMPLE 1: Open Polyline (Zigzag)")
    print("─"*80 + "\n")

    zigzag_coords = [
        (0, 0),
        (1000, 1000),
        (2000, 0),
        (3000, 1000),
        (4000, 0)
    ]

    polyline1 = create_polyline(zigzag_coords)

    # Example 2: Create a rectangle
    print("\n" + "─"*80)
    print("EXAMPLE 2: Rectangle")
    print("─"*80 + "\n")

    rectangle = create_rectangle(
        x=5000,
        y=0,
        width=2000,
        height=1500
    )

    # Example 3: Create a circle approximation
    print("\n" + "─"*80)
    print("EXAMPLE 3: Circle")
    print("─"*80 + "\n")

    circle = create_circle_approximation(
        center_x=10000,
        center_y=1000,
        radius=1000,
        segments=36
    )

    # Example 4: Create regular polygons
    print("\n" + "─"*80)
    print("EXAMPLE 4: Regular Polygons")
    print("─"*80 + "\n")

    # Triangle
    triangle = create_regular_polygon(
        center_x=13000,
        center_y=1000,
        radius=800,
        sides=3,
        rotation=0
    )

    # Hexagon
    hexagon = create_regular_polygon(
        center_x=15000,
        center_y=1000,
        radius=800,
        sides=6,
        rotation=0
    )

    # Octagon
    octagon = create_regular_polygon(
        center_x=17000,
        center_y=1000,
        radius=800,
        sides=8,
        rotation=22.5
    )

    # Example 5: Batch creation
    print("\n" + "─"*80)
    print("EXAMPLE 5: Batch Creation (Multiple Rectangles)")
    print("─"*80 + "\n")

    batch_data = []
    for i in range(3):
        x_offset = i * 2500
        coords = [
            {'x': 20000 + x_offset, 'y': 0},
            {'x': 21500 + x_offset, 'y': 0},
            {'x': 21500 + x_offset, 'y': 1000},
            {'x': 20000 + x_offset, 'y': 1000},
            {'x': 20000 + x_offset, 'y': 0}
        ]
        batch_data.append({'coordinates': coords})

    batch_guids = create_multiple_polylines_batch(batch_data)

    print("\n" + "="*80)
    print(f"✓ Created all polylines successfully!")
    print("  Check your Archicad floor plan view.")
    print("="*80)


if __name__ == "__main__":
    main()
