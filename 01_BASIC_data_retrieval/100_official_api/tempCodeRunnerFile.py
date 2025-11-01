"""
================================================================================
SCRIPT: Get Composites
================================================================================

Version:        1.0
Date:           2025-10-28
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves and displays all Composite attributes in the current Archicad project
with their folder structure organization.

Composites are multi-layer structural assemblies used in walls, slabs, and roofs.
This script provides:
- List of all composites with skin count
- Folder hierarchy organization
- GUID and index mapping
- Optional export to text file

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
  
- GetAttributesByType('Composite')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesByType

- GetCompositeAttributes()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetCompositeAttributes

- GetAttributeFolderStructure('Composite')
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributeFolderStructure

- GetAttributesIndices() - Archicad 28+
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAttributesIndices

[Data Types]
- CompositeAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.CompositeAttribute

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. View composites list with skin counts and folder structure

Optional:
4. Call export_to_file('filename.txt') to save results

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console display showing:
- Total number of composites
- List of all composites with names and skin count
- Folder structure hierarchy
- GUID/Index mapping for each composite

Optional export creates a text file with all details.

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
- Composites are multi-layer assemblies (skins)
- Each skin has a material, thickness, and function
- Skin count indicates complexity of the assembly
- Composites are organized in folders for better management
- GUIDs are permanent identifiers, indices are project-specific

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 114_01_get_building_materials.py (materials used in composite skins)
- 114_05_get_surfaces.py (surface finish for composites)
- 201_set_element_composites.py (assign composites to elements)

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

acc = conn.commands
act = conn.types


# =============================================================================
# GET COMPOSITES
# =============================================================================

def get_all_composites():
    """Get all composite attributes with details"""
    try:
        # Get IDs
        print("\n" + "="*80)
        print("COMPOSITES")
        print("="*80)
        
        attribute_ids = acc.GetAttributesByType('Composite')
        print(f"\n✓ Found {len(attribute_ids)} composite(s)")

        if not attribute_ids:
            print("\n⚠️  No composites found in the project")
            return []

        # Get details
        composites = acc.GetCompositeAttributes(attribute_ids)

        print(f"\nComposites List:")
        print("-"*80)

        for idx, comp_or_error in enumerate(composites, 1):
            if hasattr(comp_or_error, 'compositeAttribute') and comp_or_error.compositeAttribute:
                comp = comp_or_error.compositeAttribute
                name = comp.name if hasattr(comp, 'name') else 'Unknown'
                
                # Get number of skins
                num_skins = 0
                if hasattr(comp, 'compositeSkins') and comp.compositeSkins:
                    num_skins = len(comp.compositeSkins)
                elif hasattr(comp, 'skins') and comp.skins:
                    num_skins = len(comp.skins)
                
                # Get GUID
                guid = ""
                if hasattr(comp, 'attributeId') and comp.attributeId:
                    guid = comp.attributeId.guid
                
                print(f"  {idx:3}. {name} ({num_skins} skin(s))")
                if guid:
                    print(f"       GUID: {guid}")

        return composites, attribute_ids

    except Exception as e:
        print(f"\n✗ Error getting composites: {e}")
        return [], []


# =============================================================================
# GET FOLDER STRUCTURE
# =============================================================================

def get_folder_structure(attribute_type='Composite', path=None, indent=0):
    """Display the folder structure for composites"""
    try:
        structure = acc.GetAttributeFolderStructure(attribute_type, path)
        
        prefix = "  " * indent
        folder_name = structure.name if hasattr(structure, 'name') else 'Root'
        
        # Count attributes
        num_attributes = 0
        if hasattr(structure, 'attributes') and structure.attributes is not None:
            num_attributes = len(structure.attributes)
        
        # Count subfolders
        num_subfolders = 0
        if hasattr(structure, 'subfolders') and structure.subfolders is not None:
            num_subfolders = len(structure.subfolders)
        
        print(f"{prefix}📁 {folder_name} ({num_attributes} composite(s), {num_subfolders} subfolder(s))")
        
        # Display subfolders recursively
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


# =============================================================================
# GET ATTRIBUTE INDICES
# =============================================================================

def get_indices_mapping(attribute_ids):
    """Get GUID to Index mapping (if API is available)"""
    try:
        if not attribute_ids:
            return []
        
        # Check if GetAttributesIndices is available
        if not hasattr(acc, 'GetAttributesIndices'):
            print("\n" + "="*80)
            print("GUID / INDEX MAPPING")
            print("="*80)
            print("\n⚠️  GetAttributesIndices API not available in this Archicad version")
            print("    This feature requires Archicad 28 or later")
            return []
        
        print("\n" + "="*80)
        print("GUID / INDEX MAPPING")
        print("="*80)
        
        indices = acc.GetAttributesIndices(attribute_ids)
        
        print(f"\nTotal mappings: {len(indices)}")
        print("-"*80)
        
        for idx, result in enumerate(indices, 1):
            if hasattr(result, 'attributeIndexAndGuid') and result.attributeIndexAndGuid:
                index_info = result.attributeIndexAndGuid
                print(f"  {idx:3}. Index: {index_info.index:4} | GUID: {index_info.guid}")
        
        return indices
    
    except Exception as e:
        print(f"\n✗ Error getting indices: {e}")
        return []


# =============================================================================
# EXPORT TO FILE
# =============================================================================

def export_to_file(composites, filename='composites.txt'):
    """Export composites to a text file"""
    try:
        import os
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("ARCHICAD COMPOSITES\n")
            f.write("="*80 + "\n\n")
            
            for idx, comp_or_error in enumerate(composites, 1):
                if hasattr(comp_or_error, 'compositeAttribute') and comp_or_error.compositeAttribute:
                    comp = comp_or_error.compositeAttribute
                    name = comp.name if hasattr(comp, 'name') else 'Unknown'
                    
                    # Get number of skins
                    num_skins = 0
                    if hasattr(comp, 'compositeSkins') and comp.compositeSkins:
                        num_skins = len(comp.compositeSkins)
                    elif hasattr(comp, 'skins') and comp.skins:
                        num_skins = len(comp.skins)
                    
                    f.write(f"{idx}. {name} ({num_skins} skin(s))\n")
                    
                    if hasattr(comp, 'attributeId') and comp.attributeId:
                        f.write(f"   GUID: {comp.attributeId.guid}\n")
                    
                    f.write("\n")
        
        abs_path = os.path.abspath(filename)
        print(f"\n✓ Composites exported to:")
        print(f"  {abs_path}")
        
    except Exception as e:
        print(f"\n✗ Error exporting: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function"""
    print("\n" + "="*80)
    print("GET COMPOSITES v1.0")
    print("="*80)
    
    # Get all composites
    composites, attribute_ids = get_all_composites()
    
    # Show folder structure
    if composites:
        print("\n" + "="*80)
        print("FOLDER STRUCTURE")
        print("="*80)
        print()
        get_folder_structure('Composite')
    
    # Show GUID/Index mapping
    if attribute_ids:
        get_indices_mapping(attribute_ids)
    
    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total Composites: {len(composites)}")
    print("\n" + "="*80)
    
    # Export option
    print("\n💡 To export to file:")
    print("   composites, ids = get_all_composites()")
    print("   export_to_file(composites, 'my_composites.txt')")


if __name__ == "__main__":
    main()
