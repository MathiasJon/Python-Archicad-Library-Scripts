"""
Get all favorites by element type.

This script retrieves all saved favorites (predefined settings) for different
element types in Archicad using the TAPIR GetFavoritesByType command.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def get_favorites_by_type(element_type):
    """
    Get all favorites for a specific element type.

    Args:
        element_type: Type of element ('Wall', 'Slab', 'Door', 'Window', etc.)

    Returns:
        List of favorite names
    """
    try:
        # Create proper AddOnCommandId
        command_id = act.AddOnCommandId('TapirCommand', 'GetFavoritesByType')

        # Prepare parameters
        parameters = {
            'elementType': element_type
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error response
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            if 'does not have the registered Add-On command' not in error_msg:
                print(f"  ✗ Error for {element_type}: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'favorites' in response:
            return response['favorites']
        elif hasattr(response, 'favorites'):
            # Object response format
            return list(response.favorites) if hasattr(response.favorites, '__iter__') else [response.favorites]

        return []

    except Exception as e:
        # Don't print error for command not found - it's expected
        if 'does not have the registered Add-On command' not in str(e):
            print(f"  ✗ Error retrieving favorites for {element_type}: {e}")
        return []


def get_all_favorites():
    """
    Get all favorites from all element types.

    Returns:
        Dictionary mapping element types to their favorites
    """
    # Common element types that support favorites
    # NOTE: Use exact type names as they appear in Archicad API (no spaces)
    element_types = [
        'Wall', 'Slab', 'Roof', 'Beam', 'Column',
        'Door', 'Window', 'Object', 'Zone',
        'CurtainWall', 'Shell', 'Mesh', 'Stair',
        'Railing', 'Morph', 'Skylight'
    ]

    all_favorites = {}

    print("\nRetrieving favorites for each element type...")

    for elem_type in element_types:
        favorites = get_favorites_by_type(elem_type)
        if favorites and len(favorites) > 0:
            all_favorites[elem_type] = favorites
            print(f"  ✓ {elem_type}: {len(favorites)} favorite(s)")

    return all_favorites


def display_favorites(favorites_dict):
    """
    Display favorites in a formatted way.

    Args:
        favorites_dict: Dictionary of element types and their favorites
    """
    print("\n" + "="*80)
    print("FAVORITES LIBRARY")
    print("="*80)

    if not favorites_dict:
        print("\n⚠ No favorites found.")
        print("\nPossible reasons:")
        print("  • No favorites defined in this Archicad project")
        print("  • TAPIR command GetFavoritesByType not available")
        return

    total_favorites = sum(len(favs) for favs in favorites_dict.values())
    print(
        f"\nTotal: {total_favorites} favorite(s) across {len(favorites_dict)} element type(s)\n")

    for elem_type, favorites in sorted(favorites_dict.items()):
        print(f"\n{elem_type} ({len(favorites)} favorite(s)):")
        print("─"*80)

        for idx, fav_name in enumerate(favorites, 1):
            print(f"  {idx}. {fav_name}")


def find_favorite_by_name(element_type, favorite_name):
    """
    Check if a favorite with a specific name exists.

    Args:
        element_type: Type of element
        favorite_name: Name of the favorite to find

    Returns:
        True if found, False otherwise
    """
    favorites = get_favorites_by_type(element_type)

    for fav in favorites:
        if isinstance(fav, str):
            if fav.lower() == favorite_name.lower():
                return True

    return False


def export_favorites_list(filename='favorites_list.txt'):
    """
    Export a list of all favorites to a text file.

    Args:
        filename: Output filename
    """
    all_favorites = get_all_favorites()

    if not all_favorites:
        print("\n⚠ No favorites to export")
        return

    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD FAVORITES LIST\n")
            f.write("="*80 + "\n\n")

            total = sum(len(favs) for favs in all_favorites.values())
            f.write(
                f"Total: {total} favorite(s) across {len(all_favorites)} element type(s)\n\n")

            for elem_type, favorites in sorted(all_favorites.items()):
                f.write(f"\n{elem_type} ({len(favorites)} favorite(s)):\n")
                f.write("-"*80 + "\n")

                for idx, fav_name in enumerate(favorites, 1):
                    f.write(f"  {idx}. {fav_name}\n")

        print(f"\n✓ Favorites list exported to: {filename}")

    except Exception as e:
        print(f"✗ Error exporting favorites list: {e}")


def export_favorites_csv(filename='favorites_list.csv'):
    """
    Export favorites to CSV format.

    Args:
        filename: Output CSV filename
    """
    all_favorites = get_all_favorites()

    if not all_favorites:
        print("\n⚠ No favorites to export")
        return

    try:
        import csv

        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Element Type', 'Favorite Name'])

            for elem_type, favorites in sorted(all_favorites.items()):
                for fav_name in favorites:
                    writer.writerow([elem_type, fav_name])

        print(f"\n✓ Favorites exported to CSV: {filename}")

    except Exception as e:
        print(f"✗ Error exporting to CSV: {e}")


def count_favorites_by_type():
    """
    Get a count of favorites for each element type.

    Returns:
        Dictionary mapping element type to count
    """
    all_favorites = get_all_favorites()

    counts = {}
    for elem_type, favorites in all_favorites.items():
        counts[elem_type] = len(favorites)

    return counts


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("GET FAVORITES BY TYPE")
    print("="*80)

    # Example 1: Get all favorites
    print("\n" + "─"*80)
    print("EXAMPLE 1: Get All Favorites")
    print("─"*80)

    all_favorites = get_all_favorites()
    display_favorites(all_favorites)

    # Example 2: Get favorites for a specific type
    print("\n" + "─"*80)
    print("EXAMPLE 2: Wall Favorites")
    print("─"*80 + "\n")

    wall_favorites = get_favorites_by_type('Wall')

    if wall_favorites:
        print(f"Found {len(wall_favorites)} Wall favorite(s):")
        for idx, fav_name in enumerate(wall_favorites, 1):
            print(f"  {idx}. {fav_name}")
    else:
        print("⚠ No Wall favorites found")

    # Example 3: Count favorites by type
    if all_favorites:
        print("\n" + "─"*80)
        print("EXAMPLE 3: Favorites Count by Type")
        print("─"*80 + "\n")

        counts = count_favorites_by_type()
        for elem_type, count in sorted(counts.items(), key=lambda x: x[1], reverse=True):
            print(f"  {elem_type}: {count} favorite(s)")

    # Example 4: Export to file
    if all_favorites:
        print("\n" + "─"*80)
        print("EXAMPLE 4: Export Favorites")
        print("─"*80 + "\n")

        export_favorites_list('favorites_list.txt')
        export_favorites_csv('favorites_list.csv')

    print("\n" + "="*80)
    print("\n💡 TIP: Favorites are predefined element settings that you can")
    print("   save and reuse. They appear in the favorites palette in Archicad.")


if __name__ == "__main__":
    main()
