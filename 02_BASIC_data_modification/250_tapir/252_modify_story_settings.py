"""
Get and Modify Story Settings

This script uses TAPIR commands to retrieve and modify story (floor) settings:
- GetStories (v1.1.5): Get all story information
- SetStories (v1.1.5): Modify story structure

Based on TAPIR API documentation.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def get_stories():
    """
    Get information about all stories in the project.

    Returns:
        Dictionary with story structure information
    """
    try:
        # Create proper AddOnCommandId
        command_id = act.AddOnCommandId('TapirCommand', 'GetStories')

        # No parameters needed
        parameters = {}

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ GetStories failed: {error_msg}")
            return None

        return response

    except Exception as e:
        print(f"✗ Error getting stories: {e}")
        return None


def set_stories(stories_settings):
    """
    Set the story structure of the project.

    Args:
        stories_settings: List of story setting dictionaries
                         Each dictionary should contain:
                         - name (string): Story name
                         - level (number): Story level/elevation
                         - dispOnSections (boolean): Show on sections

    Returns:
        Boolean indicating success
    """
    try:
        # Create proper AddOnCommandId
        command_id = act.AddOnCommandId('TapirCommand', 'SetStories')

        # Prepare parameters
        parameters = {
            'stories': stories_settings
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'success' in response:
            if response['success']:
                print("✓ Stories updated successfully")
                return True
            else:
                error = response.get('error', {})
                print(
                    f"✗ SetStories failed: {error.get('message', 'Unknown error')}")
                return False

        # If no 'success' field, assume success
        print("✓ Stories updated")
        return True

    except Exception as e:
        print(f"✗ Error setting stories: {e}")
        return False


def display_story_info(story_data):
    """
    Display story information in a formatted way.

    Args:
        story_data: Story data from GetStories
    """
    if not story_data:
        print("\n⚠ No story information available")
        return

    print("\n" + "="*80)
    print("STORY STRUCTURE")
    print("="*80)

    # Basic info
    first_story = story_data.get('firstStory', 0)
    last_story = story_data.get('lastStory', 0)
    act_story = story_data.get('actStory', 0)
    skip_null = story_data.get('skipNullFloor', False)

    print(f"\nFirst Story Index: {first_story}")
    print(f"Last Story Index: {last_story}")
    print(f"Current Story Index: {act_story}")
    print(f"Skip Null Floor: {skip_null}")

    # Story list
    stories = story_data.get('stories', [])

    if stories:
        print(f"\nTotal Stories: {len(stories)}")
        print("\n" + "─"*80)
        print("Story Details:")
        print("─"*80 + "\n")

        for story in stories:
            index = story.get('index', 0)
            floor_id = story.get('floorId', 0)
            name = story.get('name', 'Unnamed')
            level = story.get('level', 0)
            disp_on_sections = story.get('dispOnSections', False)

            print(f"Index: {index} | Floor ID: {floor_id}")
            print(f"  Name: {name}")
            print(f"  Level: {level:,.0f} mm ({level/1000:.2f} m)")
            print(f"  Display on Sections: {disp_on_sections}")
            print()


def modify_story_name(stories_data, story_index, new_name):
    """
    Modify the name of a specific story.

    Args:
        stories_data: Current story data from GetStories
        story_index: Index of the story to modify
        new_name: New name for the story

    Returns:
        Boolean indicating success
    """
    stories = stories_data.get('stories', [])

    # Find the story
    story_to_modify = None
    for story in stories:
        if story.get('index') == story_index:
            story_to_modify = story
            break

    if not story_to_modify:
        print(f"✗ Story with index {story_index} not found")
        return False

    # Create new settings list with all stories
    new_settings = []
    for story in stories:
        settings = {
            'name': new_name if story.get('index') == story_index else story.get('name'),
            'level': story.get('level'),
            'dispOnSections': story.get('dispOnSections')
        }
        new_settings.append(settings)

    # Apply changes
    return set_stories(new_settings)


def modify_story_level(stories_data, story_index, new_level):
    """
    Modify the level/elevation of a specific story.

    Args:
        stories_data: Current story data from GetStories
        story_index: Index of the story to modify
        new_level: New level in millimeters

    Returns:
        Boolean indicating success
    """
    stories = stories_data.get('stories', [])

    # Create new settings list
    new_settings = []
    for story in stories:
        settings = {
            'name': story.get('name'),
            'level': new_level if story.get('index') == story_index else story.get('level'),
            'dispOnSections': story.get('dispOnSections')
        }
        new_settings.append(settings)

    return set_stories(new_settings)


def standardize_story_spacing(stories_data, base_level, typical_height):
    """
    Set all stories to equal spacing.

    Args:
        stories_data: Current story data from GetStories
        base_level: Level of the first story (mm)
        typical_height: Height between floors (mm)

    Returns:
        Boolean indicating success
    """
    stories = stories_data.get('stories', [])

    # Create new settings with equal spacing
    new_settings = []
    current_level = base_level

    for story in stories:
        settings = {
            'name': story.get('name'),
            'level': current_level,
            'dispOnSections': story.get('dispOnSections')
        }
        new_settings.append(settings)
        current_level += typical_height

    return set_stories(new_settings)


def rename_stories_systematic(stories_data, base_name, start_number=1):
    """
    Rename all stories systematically.

    Args:
        stories_data: Current story data from GetStories
        base_name: Base name (e.g., "Level", "Floor")
        start_number: Starting number

    Returns:
        Boolean indicating success
    """
    stories = stories_data.get('stories', [])

    # Create new settings with systematic names
    new_settings = []

    for idx, story in enumerate(stories):
        settings = {
            'name': f"{base_name} {start_number + idx}",
            'level': story.get('level'),
            'dispOnSections': story.get('dispOnSections')
        }
        new_settings.append(settings)

    return set_stories(new_settings)


def export_story_info(story_data, filename='story_info.txt'):
    """
    Export story information to a text file.

    Args:
        story_data: Story data from GetStories
        filename: Output filename
    """
    if not story_data:
        print("⚠ No story data to export")
        return

    try:
        import os

        stories = story_data.get('stories', [])

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD STORY STRUCTURE\n")
            f.write("="*80 + "\n\n")

            f.write(f"First Story Index: {story_data.get('firstStory', 0)}\n")
            f.write(f"Last Story Index: {story_data.get('lastStory', 0)}\n")
            f.write(f"Current Story: {story_data.get('actStory', 0)}\n")
            f.write(
                f"Skip Null Floor: {story_data.get('skipNullFloor', False)}\n\n")

            f.write("STORIES\n")
            f.write("-"*80 + "\n\n")

            for story in stories:
                f.write(f"Index: {story.get('index', 0)}\n")
                f.write(f"  Name: {story.get('name', 'Unnamed')}\n")
                f.write(
                    f"  Level: {story.get('level', 0)} mm ({story.get('level', 0)/1000:.2f} m)\n")
                f.write(f"  Floor ID: {story.get('floorId', 0)}\n")
                f.write(
                    f"  Display on Sections: {story.get('dispOnSections', False)}\n")
                f.write("\n")

        abs_path = os.path.abspath(filename)
        print(f"\n✓ Story information exported to:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"✗ Error exporting story info: {e}")


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("STORY SETTINGS MANAGER")
    print("="*80)

    # Get current story information
    print("\n" + "─"*80)
    print("Getting Current Story Structure...")
    print("─"*80)

    story_data = get_stories()

    if not story_data:
        print("\n✗ Failed to retrieve story information")
        return

    # Display current structure
    display_story_info(story_data)

    # Example 1: Modify a story name
    print("\n" + "─"*80)
    print("EXAMPLE 1: Rename a Story")
    print("─"*80 + "\n")

    print("💡 To rename story at index 0 to 'Ground Floor':")
    print("   Uncomment the line below and run again\n")

    # Uncomment to rename:
    # modify_story_name(story_data, 0, 'Ground Floor')
    # story_data = get_stories()  # Refresh data
    # display_story_info(story_data)

    # Example 2: Modify story level
    print("\n" + "─"*80)
    print("EXAMPLE 2: Modify Story Level")
    print("─"*80 + "\n")

    print("💡 To set story index 0 to level 0mm:")
    print("   modify_story_level(story_data, 0, 0)\n")

    # Example 3: Standardize spacing
    print("\n" + "─"*80)
    print("EXAMPLE 3: Standardize Story Spacing")
    print("─"*80 + "\n")

    print("💡 To set all stories to 3000mm spacing starting from 0:")
    print("   standardize_story_spacing(story_data, 0, 3000)\n")

    # Example 4: Systematic renaming
    print("\n" + "─"*80)
    print("EXAMPLE 4: Systematic Renaming")
    print("─"*80 + "\n")

    print("💡 To rename all stories as 'Level 0', 'Level 1', etc.:")
    print("   rename_stories_systematic(story_data, 'Level', 0)\n")

    # Example 5: Export
    print("\n" + "─"*80)
    print("EXAMPLE 5: Export Story Information")
    print("─"*80)

    export_story_info(story_data)

    print("\n" + "="*80)
    print("\n💡 TIP: Always call get_stories() first to get current data,")
    print("   then modify using the returned story_data object.")
    print("="*80)


if __name__ == "__main__":
    main()
