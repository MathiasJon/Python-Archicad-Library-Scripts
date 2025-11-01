"""
================================================================================
SCRIPT: Check Tapir Add-On Version
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       System Verification

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Verifies the installation and version of the Tapir Add-On in Archicad. This
diagnostic script provides:
- Detection of Tapir Add-On installation status
- Version information of the installed Tapir Add-On
- Verification that standard Archicad API is working
- Archicad product version and build information

Tapir is an automation add-on that extends Archicad's capabilities with
additional commands for advanced workflows and integrations.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (optional - script will report if missing)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- IsAlive()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.IsAlive
  Checks if the connection to Archicad is active

- GetProductInfo()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetProductInfo
  Returns information about the running Archicad instance

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand
  Executes a command from a loaded add-on

[Tapir API]
- TapirCommand: GetAddOnVersion
  Returns version information for the Tapir Add-On

[Data Types]
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AddOnCommandId
  Identifier for add-on commands

- ProductInfo
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ProductInfo
  Contains Archicad product information

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. Review the installation status and version information

No configuration needed - script automatically checks both Tapir and standard API.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Tapir Add-On installation status (✓ installed / ✗ not installed)
- Version information if Tapir is installed
- Standard Archicad API status verification
- Archicad version and build number

Example output (Tapir installed):
  ═══════════════════════════════════════════════════
  TAPIR ADD-ON CHECK
  ═══════════════════════════════════════════════════
  ✓ Tapir Add-On is installed
  
  Version information:
    version: 2.0.1
    buildNumber: 1234
  
  ═══════════════════════════════════════════════════
  STANDARD ARCHICAD API CHECK
  ═══════════════════════════════════════════════════
  ✓ Standard Archicad API is working
    Archicad Version: 27
    Build: 4030

Example output (Tapir not installed):
  ✗ Tapir Add-On is NOT installed or not loaded
  
  Error: [error details]
  
  Please install Tapir Add-On from:
  https://github.com/ENZYME-APD/tapir-archicad-automation

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

If you have multiple Archicad instances open:
- The script connects to the first one opened
- Make sure the correct project is open in that instance
- Close other instances if you need to target a specific project

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
ABOUT TAPIR ADD-ON:
- Tapir is an open-source automation add-on for Archicad
- Provides extended functionality beyond standard Archicad API
- Enables advanced automation workflows
- Available at: https://github.com/ENZYME-APD/tapir-archicad-automation

INSTALLATION:
- Download from GitHub releases
- Install the .apx file in Archicad Add-On Manager
- Restart Archicad after installation
- Tapir commands become available after successful installation

VERSION INFORMATION:
- Shows Tapir version and build number if installed
- Helps verify correct Tapir installation
- Useful for troubleshooting add-on issues
- Different Tapir versions may support different features

STANDARD API CHECK:
- Script also verifies standard Archicad API is working
- Shows Archicad version and build number
- Helps diagnose connection issues
- Confirms Python-Archicad communication is functional

TROUBLESHOOTING:
- If Tapir not found: Install from GitHub
- If standard API not working: Check Archicad is running
- If connection fails: Restart Archicad
- If errors persist: Check Archicad version compatibility

USE CASES:
- Verify Tapir installation before using Tapir scripts
- Check Tapir version for compatibility
- Diagnose API connection issues
- Confirm Archicad and Python are communicating
- System verification before automation workflows

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 100_test_connection.py (basic connection test)
- 150_list_installed_addons.py (list all add-ons)
- Any Tapir-specific scripts (require Tapir installation)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print("\n=== TAPIR ADD-ON CHECK ===")

# Try to get Tapir version
try:
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'GetAddOnVersion')
    )
    
    print("✓ Tapir Add-On is installed")
    print(f"\nVersion information:")
    
    # Display version info
    for key, value in response.items():
        print(f"  {key}: {value}")
    
except Exception as e:
    print("✗ Tapir Add-On is NOT installed or not loaded")
    print(f"\nError: {e}")
    print("\nPlease install Tapir Add-On from:")
    print("https://github.com/ENZYME-APD/tapir-archicad-automation")

# Check if standard Archicad API is working
print("\n=== STANDARD ARCHICAD API CHECK ===")
if acc.IsAlive():
    print("✓ Standard Archicad API is working")
    product_info = acc.GetProductInfo()
    print(f"  Archicad Version: {product_info.version}")
    print(f"  Build: {product_info.buildNumber}")
else:
    print("✗ Standard Archicad API is not responding")