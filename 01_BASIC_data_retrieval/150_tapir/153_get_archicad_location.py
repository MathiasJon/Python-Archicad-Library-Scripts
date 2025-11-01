"""
================================================================================
SCRIPT: Get Archicad Location
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       System Information

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Gets the file system path of the running Archicad executable. This script
provides:
- Full path to the Archicad executable file
- Installation directory location
- Version folder identification
- Directory accessibility verification
- Key folders/files detection (libraries, add-ons, etc.)

This information is useful for locating resources, debugging installation
issues, and setting up automation scripts.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetArchicadLocation
  Returns the file system path to the Archicad executable

[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetProductInfo()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetProductInfo
  Returns Archicad version and build information

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand
  Executes Tapir commands

[Data Types]
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AddOnCommandId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Ensure Tapir Add-On is installed
2. Open Archicad
3. Run this script
4. Review the installation path information

No configuration needed - script automatically detects Archicad location.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Archicad executable full path
- Installation directory
- Version folder name
- Directory accessibility status
- Key folders/files found
- Archicad version and build
- System information

Example output:
  ═══════════════════════════════════════════════════════
  ARCHICAD INSTALLATION INFO
  ═══════════════════════════════════════════════════════
  
  ✓ Archicad executable found:
    /Applications/GRAPHISOFT/ARCHICAD 27/ARCHICAD 27.app
  
  Directory: /Applications/GRAPHISOFT/ARCHICAD 27
  Executable: ARCHICAD 27.app
  Version folder: ARCHICAD 27
  
  ✓ Directory exists and is accessible
  
  Key folders/files found:
    • Add-Ons
    • Library
    • Templates
  
  ═══════════════════════════════════════════════════════
  ARCHICAD VERSION INFO
  ═══════════════════════════════════════════════════════
  Version: 27
  Build Number: 4030
  Language: INT

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
ARCHICAD LOCATION:
- Returns the actual executable path
- Useful for finding installation directory
- Helps locate resource files and libraries
- Platform-specific paths (Windows/Mac)

USE CASES:
- Locating Archicad library files
- Finding additional resources
- Debugging installation issues
- Setting up automation scripts
- Configuring development environments

PLATFORM DIFFERENCES:
- Windows: Typically C:\Program Files\GRAPHISOFT\ARCHICAD XX\
- macOS: Typically /Applications/GRAPHISOFT/ARCHICAD XX/

TROUBLESHOOTING:
- If command fails: Verify Tapir Add-On is installed
- If directory not accessible: Check file permissions
- If path looks unusual: May be custom installation

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 151_check_tapir_version.py (verify Tapir installation)
- 154_get_project_info.py (project information)

================================================================================
"""

from archicad import ACConnection
import os

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print("\n=== ARCHICAD INSTALLATION INFO ===")

try:
    # Get Archicad location using Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'GetArchicadLocation')
    )

    if 'archicadLocation' in response:
        location = response['archicadLocation']
        print(f"\n✓ Archicad executable found:")
        print(f"  {location}")

        # Parse directory and filename
        directory = os.path.dirname(location)
        filename = os.path.basename(location)

        print(f"\nDirectory: {directory}")
        print(f"Executable: {filename}")

        # Try to determine version from path
        if 'ARCHICAD' in location.upper():
            parts = location.split(os.sep)
            for part in parts:
                if 'ARCHICAD' in part.upper() or 'AC' in part:
                    print(f"Version folder: {part}")
                    break

        # Check if directory exists
        if os.path.exists(directory):
            print(f"\n✓ Directory exists and is accessible")
            
            # List some key files/folders in the directory
            try:
                items = os.listdir(directory)
                key_items = [item for item in items if any(keyword in item.lower() 
                    for keyword in ['library', 'add-ons', 'help', 'template'])]
                
                if key_items:
                    print(f"\nKey folders/files found:")
                    for item in sorted(key_items)[:5]:
                        print(f"  • {item}")
                    if len(key_items) > 5:
                        print(f"  ... and {len(key_items) - 5} more")
            except Exception as e:
                print(f"\n⚠ Could not list directory contents: {e}")
        else:
            print(f"\n⚠ Directory does not exist or is not accessible")

    else:
        print("✗ Could not determine Archicad location")
        print("\nResponse received:")
        for key, value in response.items():
            print(f"  {key}: {value}")

except Exception as e:
    print(f"✗ Error getting Archicad location: {e}")
    print("\nMake sure Tapir Add-On is installed")

# Also get product info from standard API
print("\n=== ARCHICAD VERSION INFO ===")

try:
    product_info = acc.GetProductInfo()
    
    # Handle different return types
    if isinstance(product_info, tuple):
        # It's a tuple: (version, buildNumber, languageCode)
        print(f"Version: {product_info[0]}")
        if len(product_info) > 1:
            print(f"Build Number: {product_info[1]}")
        if len(product_info) > 2:
            print(f"Language: {product_info[2]}")
    elif hasattr(product_info, 'version'):
        # It's an object with attributes
        print(f"Version: {product_info.version}")
        if hasattr(product_info, 'buildNumber'):
            print(f"Build Number: {product_info.buildNumber}")
        if hasattr(product_info, 'languageCode'):
            print(f"Language: {product_info.languageCode}")
    else:
        # Unknown format
        print(f"Product Info: {product_info}")

except Exception as e:
    print(f"Error: {e}")

# System info
print("\n=== SYSTEM INFO ===")
import platform
print(f"OS: {platform.system()} {platform.release()}")
print(f"Python: {platform.python_version()}")

print("\n" + "="*60)
print("💡 Use this information for:")
print("  • Locating library files")
print("  • Finding additional resources")
print("  • Debugging installation issues")
print("  • Setting up automation scripts")