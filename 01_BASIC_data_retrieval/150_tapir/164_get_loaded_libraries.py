"""
Script: Get Loaded Libraries
API: Tapir Add-On
Description: Lists all library files loaded in the current project
Usage: Run to see available library parts and their locations
Requirements:
    - archicad-api package
    - Tapir Add-On installed
"""

from archicad import ACConnection
import os

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types

print("\n=== LOADED LIBRARIES ===")

try:
    # Get loaded libraries using Tapir command
    response = acc.ExecuteAddOnCommand(
        act.AddOnCommandId('TapirCommand', 'GetLibraries')
    )

    if 'libraries' in response:
        libraries = response['libraries']
        print(f"\nTotal libraries loaded: {len(libraries)}")

        # Group libraries by type if possible
        embedded_libs = []
        external_libs = []

        for lib in libraries:
            lib_path = lib.get('path', lib.get('location', 'Unknown'))
            lib_name = lib.get('name', os.path.basename(lib_path))

            # Categorize
            if 'embedded' in lib_path.lower() or lib.get('type') == 'embedded':
                embedded_libs.append((lib_name, lib_path))
            else:
                external_libs.append((lib_name, lib_path))

        # Display embedded libraries
        if embedded_libs:
            print(f"\n{'─'*60}")
            print("📦 EMBEDDED LIBRARIES")
            print(f"{'─'*60}")
            for i, (name, path) in enumerate(embedded_libs):
                print(f"{i+1}. {name}")
                print(f"   {path}")

        # Display external libraries
        if external_libs:
            print(f"\n{'─'*60}")
            print("📁 EXTERNAL LIBRARIES")
            print(f"{'─'*60}")
            for i, (name, path) in enumerate(external_libs):
                print(f"{i+1}. {name}")
                print(f"   {path}")

        # If no categorization, just list all
        if not embedded_libs and not external_libs:
            print("\nLibraries:")
            for i, lib in enumerate(libraries):
                lib_path = lib.get('path', lib.get('location', 'Unknown'))
                lib_name = lib.get('name', os.path.basename(lib_path))
                print(f"{i+1}. {lib_name}")
                print(f"   {lib_path}")

    elif 'libraryPaths' in response:
        # Alternative response format
        library_paths = response['libraryPaths']
        print(f"\nTotal libraries: {len(library_paths)}")

        for i, path in enumerate(library_paths):
            lib_name = os.path.basename(path)
            print(f"{i+1}. {lib_name}")
            print(f"   {path}")

    else:
        print("\nRaw response:")
        for key, value in response.items():
            print(f"{key}: {value}")

except Exception as e:
    print(f"✗ Error getting libraries: {e}")
    print("\nMake sure:")
    print("  1. Tapir Add-On is installed")
    print("  2. A project is open in Archicad")

print("\n" + "="*60)
print("💡 TIPS")
print("="*60)
print("• Use Library Manager in Archicad to add/remove libraries")
print("• Script 117 can reload libraries after changes")
print("• Library parts in these folders can be used in scripts")
print("• Use exact library part names when creating objects")
