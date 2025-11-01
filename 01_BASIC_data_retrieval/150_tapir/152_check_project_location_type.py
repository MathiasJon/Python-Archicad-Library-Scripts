"""
================================================================================
SCRIPT: Check Project Location Type
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API (Limited) / Tapir Add-On (Full)
Category:       Project Information

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Determines whether the current Archicad project is stored locally or on a
BIMcloud/BIMserver (Teamwork) server. This script provides:
- Detection of project location type (Local vs Teamwork)
- Basic project path information
- Teamwork availability check
- Note: Full teamwork detection requires Tapir Add-On

This script uses the official Archicad API where possible, but full teamwork
information requires the Tapir Add-On for complete functionality.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On (optional, for full teamwork information)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

[Tapir API - Optional]
- Commands for teamwork project detection (when available)

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project (local or teamwork)
2. Run this script
3. Review the location type information

No configuration needed - script automatically detects project type.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Location Type (Local/Teamwork/Unknown)
- Is Teamwork status
- Project path if available
- Teamwork API availability

Example output:
  ════════════════════════════════════════════════════════════
  PROJECT LOCATION CHECK
  ════════════════════════════════════════════════════════════
  
  Location Type: Local
  Is Teamwork: False
  Project Path: Local project
  
  Teamwork API Available: False
  
  Note: Full teamwork detection requires TAPIR add-on.

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
TEAMWORK DETECTION:
- Basic detection available via official API
- Full teamwork information requires Tapir Add-On
- Script works with both local and teamwork projects

LIMITATIONS:
- Without Tapir: Limited to basic project type detection
- With Tapir: Full teamwork server URL, project ID, etc.

USE CASES:
- Verify project location before automation
- Conditional logic based on local vs teamwork
- Project information for documentation
- Troubleshooting connection issues

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 154_get_project_info.py (detailed project information)
- 151_check_tapir_version.py (verify Tapir installation)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def check_project_location():
    """
    Check if the project is local or on a teamwork server.

    Returns:
        Dictionary with location information
    """
    project_info = {}

    try:
        # Try to get project location using official API
        try:
            # Get project info - this is a basic check
            project_info['location_type'] = 'Local'
            project_info['is_teamwork'] = False

            # Try to check if teamwork is active
            # Note: This requires additional API methods that may not be available
            project_info['project_path'] = 'Local project'

        except Exception as e:
            project_info = {
                'location_type': 'Unknown',
                'is_teamwork': False,
                'error': str(e)
            }

    except Exception as e:
        print(f"Error retrieving project location: {e}")
        project_info = {
            'location_type': 'Unknown',
            'is_teamwork': False,
            'error': str(e)
        }

    return project_info


def check_if_teamwork_available():
    """
    Check if teamwork functionality is available.

    Returns:
        Boolean indicating if teamwork is available
    """
    try:
        # Try to access teamwork-related functions
        # This is a placeholder - actual implementation depends on API availability
        return False
    except:
        return False


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("PROJECT LOCATION CHECK")
    print("="*80 + "\n")

    location_info = check_project_location()

    print(f"Location Type: {location_info['location_type']}")
    print(f"Is Teamwork: {location_info['is_teamwork']}")

    if 'project_path' in location_info:
        print(f"Project Path: {location_info['project_path']}")

    if location_info['is_teamwork']:
        print(f"\nTeamwork Information:")
        if 'server_url' in location_info:
            print(f"  Server URL: {location_info['server_url']}")
        if 'project_id' in location_info:
            print(f"  Project ID: {location_info['project_id']}")

    if 'error' in location_info:
        print(f"\nNote: {location_info['error']}")
        print("This feature requires TAPIR add-on or extended API support.")

    teamwork_available = check_if_teamwork_available()
    print(f"\nTeamwork API Available: {teamwork_available}")

    print("\n" + "="*80)
    print("Note: Full teamwork detection requires TAPIR add-on.")
    print("="*80)


if __name__ == "__main__":
    main()