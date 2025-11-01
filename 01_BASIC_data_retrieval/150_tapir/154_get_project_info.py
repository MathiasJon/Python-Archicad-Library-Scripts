"""
================================================================================
SCRIPT: Get Project Info
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Project Information

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Gets comprehensive information about the currently open Archicad project. This
script provides:
- Project save status (saved or untitled)
- Project file location/path
- Teamwork status (solo or BIMcloud)
- Project name (for teamwork projects)
- All available project metadata

This information is essential for automation workflows, project verification,
and conditional logic based on project state.

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
- GetProjectInfo
  Returns detailed information about the current project

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
1. Open an Archicad project (saved or unsaved)
2. Run this script
3. Review the project information

No configuration needed - script reads current project state.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Project status (saved/unsaved)
- Project location path
- Teamwork mode
- Project name (if teamwork)
- All available metadata

Example output (Local project):
  ═══════════════════════════════════════════════════════
  PROJECT INFORMATION
  ═══════════════════════════════════════════════════════
  
  Project Details:
    Status: Saved
    Location: /Users/name/Projects/House.pln
    Mode: Solo Project
  
  All available information:
    isUntitled: False
    projectPath: /Users/name/Projects/House.pln
    isTeamwork: False

Example output (Teamwork project):
  PROJECT INFORMATION
  
  Project Details:
    Status: Saved
    Mode: TEAMWORK (BIMcloud)
    Project Name: Office Building
    Location: BIMcloud://server.com/Office%20Building

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
PROJECT STATUS:
- Untitled: Project has never been saved
- Saved: Project has been saved to disk or BIMcloud
- Local: Solo project on local filesystem
- Teamwork: Project on BIMcloud/BIMserver

PROJECT PATHS:
- Local projects: Standard file system paths
- Teamwork projects: BIMcloud:// URLs
- Unsaved projects: No path available

USE CASES:
- Verify project is saved before operations
- Get project path for file operations
- Check teamwork status for collaboration workflows
- Conditional logic based on project state
- Project documentation and logging

TEAMWORK INFORMATION:
- isTeamwork flag indicates BIMcloud projects
- Project name available for teamwork projects
- Server URL may be in project location
- Use for teamwork-specific workflows

TROUBLESHOOTING:
- If command fails: Ensure Tapir is installed
- If no project open: Open a project first
- For untitled projects: Some info may be unavailable

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 152_check_project_location_type.py (location type detection)
- 155_get_story_info.py (story information)
- 151_check_tapir_version.py (verify Tapir installation)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print("\n=== PROJECT INFORMATION ===")

try:
    # Get project info using Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'GetProjectInfo')
    )

    # Display project information
    print("\nProject Details:")

    # Check if project is saved
    if response.get('isUntitled', True):
        print("  Status: UNSAVED (Untitled)")
    else:
        print("  Status: Saved")

        # Project location
        if 'projectPath' in response:
            print(f"  Location: {response['projectPath']}")
        elif 'projectLocation' in response:
            print(f"  Location: {response['projectLocation']}")

    # Teamwork status
    if 'isTeamwork' in response:
        if response['isTeamwork']:
            print("  Mode: TEAMWORK (BIMcloud)")
            if 'projectName' in response:
                print(f"  Project Name: {response['projectName']}")
        else:
            print("  Mode: Solo Project")

    # Display all available info
    print("\n  All available information:")
    for key, value in response.items():
        print(f"    {key}: {value}")

except Exception as e:
    print(f"✗ Error getting project info: {e}")
    print("\nMake sure:")
    print("  1. Tapir Add-On is installed")
    print("  2. An Archicad project is open")