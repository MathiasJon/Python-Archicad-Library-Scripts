"""
================================================================================
SCRIPT: Get Story Info
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Project Information

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Gets comprehensive information about all stories (floors/levels) in the
Archicad project. This script provides:
- Complete list of all stories
- Story names and indices
- Elevation values for each story
- Story heights
- Floor indices
- Stories sorted by elevation

This information is essential for vertical navigation, element placement by
story, and understanding the project's vertical structure.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)
- An Archicad project must be open

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetStories
  Returns all stories in the project with their properties

[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Data Types]
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AddOnCommandId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. Review story levels and elevations

No configuration needed - script reads all project stories.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Project story range (first/last story indices)
- Active story (currently visible in 2D)
- Total number of stories
- Each story's complete details (sorted by level, highest first)
- Story statistics and summary

Example output:
  ════════════════════════════════════════════════════════════
  STORY INFORMATION
  ════════════════════════════════════════════════════════════
  
  PROJECT STORY RANGE:
    First Story Index: -1
    Last Story Index: 2
    Active Story (visible in 2D): 0
    Skip Null Floor: No
  
  TOTAL STORIES: 4
  
  ────────────────────────────────────────────────────────────
  STORY DETAILS:
  ────────────────────────────────────────────────────────────
  
  1. Toiture
     Story Index: 2
     Level: 9.100 m
     Floor ID: 3
     Display on Sections: Yes
  
  2. 1er étage
     Story Index: 1
     Level: 5.600 m
     Floor ID: 2
     Display on Sections: Yes
  
  3. Rez-de-chaussée
     Story Index: 0
     Level: 0.000 m
     Floor ID: 1
     Display on Sections: Yes
  
  4. Sous-sol
     Story Index: -1
     Level: -3.000 m
     Floor ID: 0
     Display on Sections: Yes
  
  ────────────────────────────────────────────────────────────
  SUMMARY:
  ────────────────────────────────────────────────────────────
    Lowest Level: -3.000 m
    Highest Level: 9.100 m
    Total Building Height: 12.100 m
    Stories Above Ground: 3
    Stories Below Ground: 1
  
  ════════════════════════════════════════════════════════════

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
PROJECT-LEVEL PROPERTIES (from Tapir API v1.1.5):
- firstStory: Index of the first (lowest) story
- lastStory: Index of the last (highest) story
- actStory: Currently active/visible story in 2D views
- skipNullFloor: Whether floor numbering skips 0 above ground level

STORY PROPERTIES (per story):
- index: Internal story index (can be negative for basements)
- name: Display name of the story
- level: Vertical position from project zero (meters)
- floorId: Unique identifier for the story
- dispOnSections: Whether story level lines appear on sections/elevations

STORY INDEX:
- Zero-based or can start from negative values
- Negative indices typically for below-ground stories
- Index 0 is often ground floor (but not always)
- Can be non-contiguous

LEVEL VALUES:
- Measured from project zero point (meters)
- Negative values for below-ground stories
- Positive values for above-ground stories
- Represents the floor plane elevation
- Used for element vertical positioning

FLOOR ID:
- Unique identifier for each story
- Persistent across project changes
- Used internally by Archicad
- Different from story index

DISPLAY ON SECTIONS:
- Controls visibility of story level lines
- On sections and elevations
- Boolean flag per story
- Affects drawing output

SORTING:
- Stories displayed from highest to lowest level
- Makes it easy to understand vertical structure
- Roof/top floors appear first
- Based on 'level' property, not index

SKIP NULL FLOOR:
- When true: Floor numbering jumps from -1 to 1
- When false: Floor numbering includes 0
- Affects floor numbering display
- Common in European projects (no "floor 0")

USE CASES:
- Understanding project vertical structure
- Element placement calculations by story
- Vertical navigation in automation
- Story-based element filtering
- Building height calculations
- Floor numbering logic
- Export to other systems
- IFC export preparation

ARCHICAD STORY STRUCTURE:
- Each project has at least one story
- Stories define working planes for 2D/3D
- Elements are typically assigned to stories
- Story settings affect element behavior
- Stories can have custom elevations

LEVEL vs INDEX:
- level: Physical elevation in meters
- index: Logical story number
- Not always directly related
- Use level for spatial calculations
- Use index for story identification

SUMMARY CALCULATIONS:
- Script calculates total building height
- Counts above/below ground stories
- Shows level range
- Useful for project overview

TROUBLESHOOTING:
- If command fails: Ensure Tapir Add-On is installed
- If no stories: Project may be corrupted (very unusual)
- If unexpected values: Check project story settings
- If negative indices: Normal for basement stories

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 154_get_project_info.py (project information)
- 101_get_all_elements.py (elements on stories)
- 151_check_tapir_version.py (verify Tapir installation)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print("\n" + "="*80)
print("STORY INFORMATION")
print("="*80 + "\n")

try:
    # Get story info using Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'GetStories')
    )

    # Display project-level story information
    if 'firstStory' in response or 'lastStory' in response:
        print("PROJECT STORY RANGE:")
        if 'firstStory' in response:
            print(f"  First Story Index: {response['firstStory']}")
        if 'lastStory' in response:
            print(f"  Last Story Index: {response['lastStory']}")
        if 'actStory' in response:
            print(f"  Active Story (visible in 2D): {response['actStory']}")
        if 'skipNullFloor' in response:
            skip_null = "Yes" if response['skipNullFloor'] else "No"
            print(f"  Skip Null Floor: {skip_null}")
        print()

    # Check if we have stories
    if 'stories' in response:
        stories = response['stories']
        print(f"TOTAL STORIES: {len(stories)}")
        print("\n" + "-"*80)
        print("STORY DETAILS:")
        print("-"*80 + "\n")

        # Sort stories by level (highest first)
        sorted_stories = sorted(stories,
                                key=lambda s: s.get('level', s.get('index', 0)),
                                reverse=True)

        for i, story in enumerate(sorted_stories):
            story_num = i + 1
            story_name = story.get('name', 'Unnamed Story')
            story_index = story.get('index', 'N/A')
            
            print(f"{story_num}. {story_name}")
            print(f"   Story Index: {story_index}")

            # Level (elevation from project zero)
            if 'level' in story:
                level_m = story['level']
                print(f"   Level: {level_m:.3f} m")

            # Floor ID (unique identifier)
            if 'floorId' in story:
                print(f"   Floor ID: {story['floorId']}")

            # Display on sections
            if 'dispOnSections' in story:
                disp = "Yes" if story['dispOnSections'] else "No"
                print(f"   Display on Sections: {disp}")

            print()  # Empty line for readability

        # Summary statistics
        print("-"*80)
        print("SUMMARY:")
        print("-"*80)
        
        # Calculate level range
        levels = [s.get('level', 0) for s in stories if 'level' in s]
        if levels:
            min_level = min(levels)
            max_level = max(levels)
            total_height = max_level - min_level
            
            print(f"  Lowest Level: {min_level:.3f} m")
            print(f"  Highest Level: {max_level:.3f} m")
            print(f"  Total Building Height: {total_height:.3f} m")
            
            # Count above/below ground
            above_ground = sum(1 for level in levels if level >= 0)
            below_ground = sum(1 for level in levels if level < 0)
            
            print(f"  Stories Above Ground: {above_ground}")
            print(f"  Stories Below Ground: {below_ground}")

    else:
        print("No story information available")
        print("\nRaw response:")
        for key, value in response.items():
            print(f"  {key}: {value}")

    print("\n" + "="*80)

except Exception as e:
    print(f"✗ Error getting story info: {e}")
    print("\nMake sure:")
    print("  1. Tapir Add-On is installed")
    print("  2. An Archicad project is open")
    print("\n" + "="*80)