"""
Script: Reload Libraries
API: Tapir Add-On
Description: Reloads all library files in the current project
Usage: Run after modifying GDL objects or adding new library files
Requirements:
    - archicad-api package
    - Tapir Add-On installed
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print("\n=== RELOADING LIBRARIES ===")
print("This will reload all library files in the project...")

# Optional: Get library count before reload
try:
    response_before = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'GetLoadedLibraries')
    )

    if 'libraries' in response_before:
        lib_count = len(response_before['libraries'])
        print(f"Current libraries loaded: {lib_count}")
    elif 'libraryPaths' in response_before:
        lib_count = len(response_before['libraryPaths'])
        print(f"Current libraries loaded: {lib_count}")
except:
    pass

print("\nReloading...")

try:
    # Reload libraries using Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'ReloadLibraries')
    )

    print("\n✓ Libraries reloaded successfully")

    # Check response for any messages
    if isinstance(response, dict):
        if 'message' in response:
            print(f"Message: {response['message']}")

        if 'reloadedCount' in response:
            print(f"Reloaded: {response['reloadedCount']} libraries")

        if 'errors' in response and response['errors']:
            print("\n⚠ Reload warnings/errors:")
            for error in response['errors']:
                print(f"  - {error}")

    print("\n" + "="*60)
    print("RELOAD COMPLETE")
    print("="*60)

    print("\n💡 When to reload libraries:")
    print("  • After editing GDL objects externally")
    print("  • After adding new library files")
    print("  • After updating library folders")
    print("  • When objects appear outdated")

except Exception as e:
    print(f"\n✗ Error reloading libraries: {e}")
    print("\nPossible causes:")
    print("  1. Tapir Add-On not installed")
    print("  2. Library files have errors")
    print("  3. Library folders not accessible")
    print("  4. Insufficient permissions")

print("\n" + "="*60)
print("RELATED SCRIPTS")
print("="*60)
print("• Script 40: View currently loaded libraries")
print("• Script 118: Add new library folders")
print("• Script 109: Create objects from library parts")
