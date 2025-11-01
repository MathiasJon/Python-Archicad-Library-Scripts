"""
================================================================================
SCRIPT: Get Product Info
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       System Information

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Gets detailed information about the Archicad installation including version,
build number, language code, API connection status, and Tapir Add-On presence.

This script is useful for system diagnostics and verifying the environment
before running automation scripts.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetProductInfo()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetProductInfo
  
- IsAlive()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.IsAlive
  
- ExecuteAddOnCommand()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Data Types]
 
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.AddOnCommandId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Ensure Archicad is running
2. Run this script
3. View system information in console

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:

1. ARCHICAD PRODUCT INFORMATION:
   - Version number
   - Build number
   - Language code and full language name

2. VERSION ANALYSIS:
   - Major version identification
   - Python API support level assessment

3. API CONNECTION STATUS:
   - Connection active status
   - Archicad response check
   - Connection details (host, port)

4. TAPIR ADD-ON CHECK:
   - Installation status
   - Version if installed

5. SYSTEM SUMMARY:
   - Complete overview of environment
   - Readiness for automation

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
- GetProductInfo() returns a tuple: (version, buildNumber, languageCode)
- Version analysis provides Python API support assessment
- Language codes are mapped to full language names when available
- Tapir Add-On is optional but recommended for advanced automation
- Script works even without Tapir installed

--------------------------------------------------------------------------------
TAPIR ADD-ON:
--------------------------------------------------------------------------------
Tapir is an optional add-on that extends Archicad's automation capabilities:
- Download: github.com/ENZYME-APD/tapir-archicad-automation
- Not required for basic Python API functionality
- Provides additional commands and features

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 101_get_all_classification_systems.py
- 104_get_all_elements.py
- 108_get_all_properties.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

# ============================================================================
# LANGUAGE CODE MAPPING
# ============================================================================
LANGUAGE_NAMES = {
    'AUS': 'Australian',
    'AUT': 'Austrian',
    'BRA': 'Brazilian Portuguese',
    'CZE': 'Czech',
    'DEN': 'Danish',
    'NED': 'Dutch',
    'FIN': 'Finnish',
    'FRA': 'French',
    'GER': 'German',
    'GRE': 'Greek',
    'HUN': 'Hungarian',
    'INT': 'International',
    'ITA': 'Italian',
    'JPN': 'Japanese',
    'KOR': 'Korean',
    'MEX': 'Mexico',
    'NZE': 'New Zealand',
    'NOR': 'Norwegian',
    'POL': 'Polish',
    'POR': 'Portuguese',
    'RUS': 'Russian',
    'CHI': 'Simplified Chinese',
    'SPA': 'Spanish',
    'SWE': 'Swedish',
    'CHE': 'Swiss',
    'TAI': 'Traditional Chinese',
    'TUR': 'Turkish',
    'UKR': 'Ukrainian',
    'UKI': 'United Kingdom & Ireland',
    'USA': 'US'
}


def get_language_name(code):
    """
    Get the full language name from code, or return the code if not found.

    Args:
        code: Language code (e.g., 'FRA', 'USA')

    Returns:
        Full language name or the code itself
    """
    return LANGUAGE_NAMES.get(code, code)


# ============================================================================
# ARCHICAD PRODUCT INFORMATION
# ============================================================================
print("\n" + "="*70)
print("ARCHICAD PRODUCT INFORMATION")
print("="*70)

try:
    # Get product information (returns tuple: (version, buildNumber, languageCode))
    product_info = acc.GetProductInfo()

    print("\n✓ Product information retrieved")

    # Extract values from tuple
    if isinstance(product_info, tuple) and len(product_info) >= 3:
        version = product_info[0]
        build_number = product_info[1]
        language_code = product_info[2]

        # Display version details
        print("\n--- Version Details ---")
        print(f"Version: {version}")
        print(f"Build Number: {build_number}")
        print(f"Language Code: {language_code}")
        print(f"Language: {get_language_name(language_code)}")

        # Analyze version
        print(f"\n--- Version Analysis ---")
        print(f"Major Version: {version}")

        # Version-specific notes
        try:
            v = int(version)
            if v >= 27:
                print("✓ Modern version with full Python API support")
            elif v >= 25:
                print("✓ Good Python API support")
            elif v >= 23:
                print("⚠️  Limited Python API support")
            else:
                print("⚠️  Very limited or no Python API support")
        except:
            print("Unable to analyze version number")
    else:
        print("\n⚠️  Unexpected product info format")
        print(f"Received: {product_info}")

except Exception as e:
    print(f"\n✗ Error getting product info: {e}")
    import traceback
    traceback.print_exc()
    print("Make sure Archicad is running and connected")

# ============================================================================
# API CONNECTION STATUS
# ============================================================================
print("\n" + "="*70)
print("API CONNECTION STATUS")
print("="*70)

try:
    if acc.IsAlive():
        print("\n✓ API connection is active")
        print("✓ Archicad is responding to commands")
    else:
        print("\n✗ API connection issue detected")
except Exception as e:
    print(f"\n✗ Cannot check connection: {e}")

# Display connection details
print("\n--- Connection Details ---")
print(f"Host: {conn.host if hasattr(conn, 'host') else 'localhost'}")
print(f"Port: {conn.port if hasattr(conn, 'port') else 'default'}")

# ============================================================================
# TAPIR ADD-ON CHECK
# ============================================================================
print("\n" + "="*70)
print("TAPIR ADD-ON CHECK")
print("="*70)

tapir_installed = False
try:
    tapir_response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'GetAddOnVersion')
    )
    print("\n✓ Tapir Add-On is installed")
    tapir_installed = True

    if 'version' in tapir_response:
        print(f"Tapir Version: {tapir_response['version']}")
except:
    print("\n✗ Tapir Add-On is not installed")
    print("  Download from: github.com/ENZYME-APD/tapir-archicad-automation")
    print("  Note: Tapir is optional for Python API automation")

# ============================================================================
# SYSTEM SUMMARY
# ============================================================================
print("\n" + "="*70)
print("SYSTEM SUMMARY")
print("="*70)

try:
    product_info = acc.GetProductInfo()

    # Extract from tuple
    if isinstance(product_info, tuple) and len(product_info) >= 3:
        version = product_info[0]
        build = product_info[1]
        language = product_info[2]

        print(f"\nArchicad {version} (Build {build})")
        print(f"Language: {get_language_name(language)} ({language})")
    else:
        print(f"\nArchicad (Product info format: {type(product_info)})")

    print("Python API: Active")

    if tapir_installed:
        print("Tapir Add-On: Installed")
    else:
        print("Tapir Add-On: Not installed")

    print("\n✓ System ready for automation")

except:
    print("\n⚠️  System check incomplete")

print("\n" + "="*70)
print("✓ Information retrieval complete")
print("\n💡 TIP: This script helps verify your environment before running")
print("         automation scripts. Keep Archicad updated for best results.")
