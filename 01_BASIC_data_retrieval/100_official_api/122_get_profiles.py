"""
================================================================================
SCRIPT: Get Profiles (Complete)
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Profile attributes in the current Archicad project
with complete detailed information including dimensions, stretchability, and
folder structure organization.

Profiles are custom cross-sectional shapes used in beams, columns, and complex
profiles. This script provides comprehensive profile information:
- List of all profiles with full details
- Dimensions (width, height, minimum sizes)
- Stretchability properties
- Core skin information
- Compatible element types
- Profile modifiers
- Folder hierarchy organization
- Optional export to text file

For profile preview images, see script 115_get_profile_previews.py

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- GetAttributesByType('Profile')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetProfileAttributes(attributeIds)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetProfileAttributes
  Returns detailed profile attributes identified by their GUIDs including:
  * attributeId: The identifier of the attribute
  * name: The name of the profile
  * useWith: List of compatible element types
  * width: Default width (horizontal size)
  * height: Default height (vertical size)
  * minimumWidth: Minimum width constraint
  * minimumHeight: Minimum height constraint
  * widthStretchable: Can width be increased beyond default
  * heightStretchable: Can height be increased beyond default
  * hasCoreSkin: Whether profile has a core skin
  * profileModifiers: List of profile modifier parameters

- GetAttributeFolderStructure('Profile')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

[Data Types]
- ProfileAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.ProfileAttribute

- ProfileAttributeOrError
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.ProfileAttributeOrError

- ProfileModifierListItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.ProfileModifierListItem

- ProfileModifier
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.ProfileModifier
  Profile modifier parameter with:
  * name: The name of the modifier
  * value: The value of the modifier

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project containing profiles
2. Run this script
3. View comprehensive profiles list with all properties
4. View folder structure organization

Optional:
5. Call export_to_file(profiles, 'filename.txt') to save detailed results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of profiles
- For each profile:
  * Name and GUID
  * Dimensions (width × height)
  * Minimum dimensions
  * Stretchability (width/height)
  * Core skin status
  * Compatible element types (useWith)
  * Profile modifiers with name and value
- Folder hierarchy structure
- Summary statistics

Example output for a profile:
  1. Complex Profile Name
     GUID: ABC123...
     Dimensions: 500.00 × 300.00 (W × H)
     Minimum: 200.00 × 150.00
     Stretchable: Width, Height
     Core Skin: Yes
     Use With: Wall, Beam, Column
     Modifiers: 2
        • Width Offset: 50.00
        • Height Scale: 1.25

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
- Profiles define custom cross-sectional shapes for structural elements
- Width refers to horizontal dimension, height to vertical dimension
- Stretchable profiles can be scaled beyond their default size
- Core skin affects how profile layers are treated in composites
- Profile modifiers allow parametric adjustments to profile geometry
- Each modifier has a name and numerical value that controls profile parameters
- Profiles can be restricted to specific element types (useWith)
- This script does NOT generate preview images (see 115_get_profile_previews.py)

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 115_get_profile_previews.py (generate PNG previews of profiles)
- 114_01_get_building_materials.py (materials used in profiles)
- 201_create_profile.py (create new profile attributes)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


def get_all_profiles():
    """
    Get all profile attributes with complete details.
    
    Returns:
        Tuple of (profiles list, attribute IDs list)
    """
    try:
        print("\n" + "="*80)
        print("PROFILES - DETAILED INFORMATION")
        print("="*80)
        
        # Get all profile attribute IDs
        attribute_ids = acc.GetAttributesByType('Profile')
        print(f"\n✓ Found {len(attribute_ids)} profile(s)")

        if not attribute_ids:
            print("\n⚠️  No profiles found in the project")
            return [], []

        # Get detailed profile attributes
        profiles = acc.GetProfileAttributes(attribute_ids)

        print(f"\nProfiles List:")
        print("-"*80)

        for idx, prof_or_error in enumerate(profiles, 1):
            if hasattr(prof_or_error, 'profileAttribute') and prof_or_error.profileAttribute:
                prof = prof_or_error.profileAttribute
                
                # Basic information
                name = prof.name if hasattr(prof, 'name') else 'Unknown'
                guid = prof.attributeId.guid if hasattr(prof, 'attributeId') and prof.attributeId else 'N/A'
                
                print(f"\n{idx}. {name}")
                print(f"   GUID: {guid}")
                
                # Dimensions
                if hasattr(prof, 'width') and hasattr(prof, 'height'):
                    print(f"   Dimensions: {prof.width:.2f} × {prof.height:.2f} (W × H)")
                
                if hasattr(prof, 'minimumWidth') and hasattr(prof, 'minimumHeight'):
                    print(f"   Minimum: {prof.minimumWidth:.2f} × {prof.minimumHeight:.2f}")
                
                # Stretchability
                if hasattr(prof, 'widthStretchable') and hasattr(prof, 'heightStretchable'):
                    stretch_info = []
                    if prof.widthStretchable:
                        stretch_info.append("Width")
                    if prof.heightStretchable:
                        stretch_info.append("Height")
                    
                    if stretch_info:
                        print(f"   Stretchable: {', '.join(stretch_info)}")
                    else:
                        print(f"   Stretchable: No")
                
                # Core skin
                if hasattr(prof, 'hasCoreSkin'):
                    print(f"   Core Skin: {'Yes' if prof.hasCoreSkin else 'No'}")
                
                # Compatible element types
                if hasattr(prof, 'useWith') and prof.useWith:
                    if len(prof.useWith) > 0:
                        print(f"   Use With: {', '.join(prof.useWith)}")
                    else:
                        print(f"   Use With: All element types")
                
                # Profile modifiers
                if hasattr(prof, 'profileModifiers') and prof.profileModifiers:
                    print(f"   Modifiers: {len(prof.profileModifiers)}")
                    for mod_item in prof.profileModifiers:
                        if hasattr(mod_item, 'profileModifier') and mod_item.profileModifier:
                            modifier = mod_item.profileModifier
                            mod_name = modifier.name if hasattr(modifier, 'name') else 'Unknown'
                            mod_value = modifier.value if hasattr(modifier, 'value') else 0.0
                            print(f"      • {mod_name}: {mod_value:.2f}")
            
            elif hasattr(prof_or_error, 'error'):
                print(f"\n{idx}. Error: {prof_or_error.error}")

        return profiles, attribute_ids

    except Exception as e:
        print(f"\n✗ Error getting profiles: {e}")
        import traceback
        traceback.print_exc()
        return [], []


def get_folder_structure(attribute_type='Profile', path=None, indent=0):
    """
    Display the folder structure for profiles.
    
    Args:
        attribute_type: Type of attribute (default: 'Profile')
        path: Current folder path
        indent: Indentation level for display
    
    Returns:
        Folder structure object
    """
    try:
        structure = acc.GetAttributeFolderStructure(attribute_type, path)
        
        prefix = "  " * indent
        folder_name = structure.name if hasattr(structure, 'name') else 'Root'
        
        num_attributes = 0
        if hasattr(structure, 'attributes') and structure.attributes is not None:
            num_attributes = len(structure.attributes)
        
        num_subfolders = 0
        if hasattr(structure, 'subfolders') and structure.subfolders is not None:
            num_subfolders = len(structure.subfolders)
        
        print(f"{prefix}📁 {folder_name} ({num_attributes} profile(s), {num_subfolders} subfolder(s))")
        
        # Recursively display subfolders
        if hasattr(structure, 'subfolders') and structure.subfolders is not None and len(structure.subfolders) > 0:
            for subfolder_item in structure.subfolders:
                if hasattr(subfolder_item, 'attributeFolder'):
                    folder_struct = subfolder_item.attributeFolder
                    if hasattr(folder_struct, 'name'):
                        subfolder_path = path.copy() if path else []
                        subfolder_path.append(folder_struct.name)
                        get_folder_structure(attribute_type, subfolder_path, indent + 1)
        
        return structure
    
    except Exception as e:
        print(f"{prefix}✗ Error: {e}")
        return None


def export_to_file(profiles, filename='profiles_detailed.txt'):
    """
    Export profiles to a detailed text file.
    
    Args:
        profiles: List of ProfileAttributeOrError objects
        filename: Output filename (default: 'profiles_detailed.txt')
    """
    try:
        import os
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD PROFILES - DETAILED INFORMATION\n")
            f.write("="*80 + "\n\n")
            
            for idx, prof_or_error in enumerate(profiles, 1):
                if hasattr(prof_or_error, 'profileAttribute') and prof_or_error.profileAttribute:
                    prof = prof_or_error.profileAttribute
                    
                    # Basic information
                    name = prof.name if hasattr(prof, 'name') else 'Unknown'
                    f.write(f"{idx}. {name}\n")
                    
                    if hasattr(prof, 'attributeId') and prof.attributeId:
                        f.write(f"   GUID: {prof.attributeId.guid}\n")
                    
                    # Dimensions
                    if hasattr(prof, 'width') and hasattr(prof, 'height'):
                        f.write(f"   Dimensions: {prof.width:.2f} × {prof.height:.2f} (W × H)\n")
                    
                    if hasattr(prof, 'minimumWidth') and hasattr(prof, 'minimumHeight'):
                        f.write(f"   Minimum: {prof.minimumWidth:.2f} × {prof.minimumHeight:.2f}\n")
                    
                    # Stretchability
                    if hasattr(prof, 'widthStretchable') and hasattr(prof, 'heightStretchable'):
                        stretch_info = []
                        if prof.widthStretchable:
                            stretch_info.append("Width")
                        if prof.heightStretchable:
                            stretch_info.append("Height")
                        
                        if stretch_info:
                            f.write(f"   Stretchable: {', '.join(stretch_info)}\n")
                        else:
                            f.write(f"   Stretchable: No\n")
                    
                    # Core skin
                    if hasattr(prof, 'hasCoreSkin'):
                        f.write(f"   Core Skin: {'Yes' if prof.hasCoreSkin else 'No'}\n")
                    
                    # Compatible element types
                    if hasattr(prof, 'useWith') and prof.useWith:
                        if len(prof.useWith) > 0:
                            f.write(f"   Use With: {', '.join(prof.useWith)}\n")
                    
                    # Profile modifiers
                    if hasattr(prof, 'profileModifiers') and prof.profileModifiers:
                        f.write(f"   Modifiers: {len(prof.profileModifiers)}\n")
                        for mod_item in prof.profileModifiers:
                            if hasattr(mod_item, 'profileModifier') and mod_item.profileModifier:
                                modifier = mod_item.profileModifier
                                mod_name = modifier.name if hasattr(modifier, 'name') else 'Unknown'
                                mod_value = modifier.value if hasattr(modifier, 'value') else 0.0
                                f.write(f"      • {mod_name}: {mod_value:.2f}\n")
                    
                    f.write("\n")
        
        abs_path = os.path.abspath(filename)
        print(f"\n✓ Profiles exported to:")
        print(f"  {abs_path}")
        
    except Exception as e:
        print(f"\n✗ Error exporting: {e}")


def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("GET PROFILES v1.0")
    print("="*80)
    
    # Get all profiles with details
    profiles, attribute_ids = get_all_profiles()
    
    # Display folder structure
    if profiles:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('Profile')
    
    # Display summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Profiles: {len(profiles)}")
    
    # Count stretchable profiles
    stretchable_count = 0
    core_skin_count = 0
    for prof_or_error in profiles:
        if hasattr(prof_or_error, 'profileAttribute') and prof_or_error.profileAttribute:
            prof = prof_or_error.profileAttribute
            if hasattr(prof, 'widthStretchable') and hasattr(prof, 'heightStretchable'):
                if prof.widthStretchable or prof.heightStretchable:
                    stretchable_count += 1
            if hasattr(prof, 'hasCoreSkin') and prof.hasCoreSkin:
                core_skin_count += 1
    
    print(f"  Stretchable Profiles: {stretchable_count}")
    print(f"  Profiles with Core Skin: {core_skin_count}")
    print("\n" + "="*80)
    
    # Usage hints
    print("\n💡 To export detailed information to file:")
    print("   profiles, ids = get_all_profiles()")
    print("   export_to_file(profiles, 'my_profiles_detailed.txt')")
    
    print("\n💡 To generate profile preview images:")
    print("   See script: 115_get_profile_previews.py")


if __name__ == "__main__":
    main()