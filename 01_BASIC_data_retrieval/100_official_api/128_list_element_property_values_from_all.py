"""
================================================================================
SCRIPT: List All Elements with Property Values
================================================================================

Version:        1.0
Date:           2025-10-27
API Type:       Official Archicad API
Category:       Data Presentation

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves ALL elements from the project and displays detailed information
including type, GUID, classification, and property values for each element.

This script automatically analyzes all elements without requiring manual
selection, making it ideal for project audits and property verification.

⚠️  WARNING: For large projects with many elements, this script may take
several minutes to complete.

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)

--------------------------------------------------------------------------------
CONFIGURATION:
--------------------------------------------------------------------------------
Modify these variables in the script:
- MAX_ELEMENTS_TO_ANALYZE : Maximum number of elements to analyze (default: 50)
                            Set to None to analyze all elements (may be slow)
- SHOW_EMPTY_PROPERTIES   : Show empty properties list (default: False)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection
  
- GetAllElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllElements
  
- GetTypesOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetTypesOfElements
  
- GetAllPropertyIds()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllPropertyIds
  
- GetDetailsOfProperties()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetDetailsOfProperties
  
- GetClassificationsOfElements()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetClassificationsOfElements
  
- GetAllClassificationSystems()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationSystems
  
- GetAllClassificationsInSystem()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationsInSystem

[Utilities]
- GetPropertyValuesDictionary()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b3000utilities.Utilities.GetPropertyValuesDictionary 

[Data Types]
- ElementIdArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementIdArrayItem
  
- ElementId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ElementId
  
- TypeOfElementWrapper
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac26.html#archicad.releases.ac26.b3005types.TypeOfElementWrapper 

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Adjust MAX_ELEMENTS_TO_ANALYZE if needed (set to None for all elements)
3. Run this script
4. View detailed element information and properties in console

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
For each element, displays:

1. ELEMENT IDENTIFICATION:
   - Element number and total count
   - Element type
   - GUID
   - Classifications by system (with ID and name)

2. PROPERTY VALUES:
   - Properties with values (sorted alphabetically)
   - Count of empty properties
   - Optional: List of empty properties

3. SUMMARY:
   - Total elements in project
   - Elements analyzed

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
- Default limit is 50 elements to prevent long execution times
- Set MAX_ELEMENTS_TO_ANALYZE = None to analyze all elements
- Property values longer than 50 characters are truncated with "..."
- Empty properties are counted but not listed by default
- Classification names are retrieved by loading all classifications at startup
- For large projects, consider filtering by element type first

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 104_get_all_elements.py
- 106_get_elements_by_type.py
- 111_get_element_full_info.py
- 202_set_property_value.py

================================================================================
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types
acu = conn.utilities

# ============================================================================
# CONFIGURATION
# ============================================================================
MAX_ELEMENTS_TO_ANALYZE = 50  # Set to None to analyze ALL elements
SHOW_EMPTY_PROPERTIES = False  # Set to True to list empty properties

# ============================================================================
# GET ALL ELEMENTS
# ============================================================================
all_elements = acc.GetAllElements()

print(f"\n{'='*70}")
print(f"ELEMENT PROPERTY ANALYSIS")
print(f"{'='*70}")
print(f"Total elements in project: {len(all_elements)}")

if len(all_elements) == 0:
    print("\n⚠️  No elements found in the project")
    exit()

# Determine how many elements to analyze
if MAX_ELEMENTS_TO_ANALYZE is None:
    elements_to_analyze = all_elements
    print(f"Analyzing: ALL {len(all_elements)} elements")
else:
    elements_to_analyze = all_elements[:MAX_ELEMENTS_TO_ANALYZE]
    print(
        f"Analyzing: First {len(elements_to_analyze)} elements (limited by MAX_ELEMENTS_TO_ANALYZE)")
    if len(all_elements) > MAX_ELEMENTS_TO_ANALYZE:
        print(
            f"⚠️  Remaining {len(all_elements) - MAX_ELEMENTS_TO_ANALYZE} elements not analyzed")
        print("   Set MAX_ELEMENTS_TO_ANALYZE = None to analyze all elements")

print("\n⏳ Loading classification systems and property definitions...")

# Get all property IDs once for efficiency
all_property_ids = acc.GetAllPropertyIds()

# ============================================================================
# BUILD CLASSIFICATION LOOKUP DICTIONARY
# ============================================================================
# Get all classification systems and build lookup dictionary
try:
    classification_systems = acc.GetAllClassificationSystems()
    classification_system_ids = [
        system.classificationSystemId for system in classification_systems]
    system_name_lookup = {
        system.classificationSystemId.guid: system.name for system in classification_systems}

    # Build a lookup dictionary for all classification items
    # Format: {item_guid: (item_id, item_name)}
    classification_item_lookup = {}

    def add_items_to_lookup(item_wrapper):
        """Recursively add classification items to lookup dictionary"""
        if hasattr(item_wrapper, 'classificationItem'):
            item = item_wrapper.classificationItem
            if hasattr(item, 'id') and hasattr(item, 'name') and hasattr(item, 'classificationItemId'):
                item_guid = item.classificationItemId.guid
                classification_item_lookup[item_guid] = (item.id, item.name)

            # Process children recursively
            if hasattr(item, 'children') and item.children:
                for child in item.children:
                    add_items_to_lookup(child)

    # Load all classification items from all systems
    for system in classification_systems:
        try:
            all_items = acc.GetAllClassificationsInSystem(
                system.classificationSystemId)
            for item_wrapper in all_items:
                add_items_to_lookup(item_wrapper)
        except Exception as e:
            print(
                f"⚠️  Warning: Could not load classifications for system {system.name}: {e}")

    print(f"✓ Loaded {len(classification_item_lookup)} classification items")

except Exception as e:
    print(f"⚠️  Warning: Unable to retrieve classification systems: {e}")
    classification_systems = []
    classification_system_ids = []
    system_name_lookup = {}
    classification_item_lookup = {}

print("\n⏳ Processing elements...\n")

# ============================================================================
# PROCESS EACH ELEMENT
# ============================================================================
for elem_idx, element in enumerate(elements_to_analyze):
    element_id = element.elementId

    print(f"\n{'='*70}")
    print(f"ELEMENT {elem_idx + 1}/{len(elements_to_analyze)}")
    print(f"{'='*70}")

    # ========================================================================
    # GET ELEMENT TYPE
    # ========================================================================
    element_type = "Unknown"
    try:
        element_types = acc.GetTypesOfElements([element_id])
        if element_types:
            element_type = element_types[0].typeOfElement.elementType
    except Exception as e:
        pass

    print(f"Type: {element_type}")
    print(f"GUID: {element_id.guid}")

# ========================================================================
    # GET CLASSIFICATION
    # ========================================================================
    if classification_system_ids:
        try:
            classifications = acc.GetClassificationsOfElements(
                [element_id], classification_system_ids)

            print("\nClassifications:")

            if classifications and len(classifications) > 0:
                elem_classifications = classifications[0].classificationIds

                # Map classifications by system ID
                classification_by_system = {}
                for class_id_wrapper in elem_classifications:
                    if class_id_wrapper and hasattr(class_id_wrapper, 'classificationId'):
                        class_id = class_id_wrapper.classificationId
                        if hasattr(class_id, 'classificationSystemId'):
                            system_guid = class_id.classificationSystemId.guid
                            classification_by_system[system_guid] = class_id

                # Display classification for each system
                for system in classification_systems:
                    system_guid = system.classificationSystemId.guid
                    system_name = system.name

                    if system_guid in classification_by_system:
                        class_id = classification_by_system[system_guid]
                        # Check if classification item is assigned
                        if hasattr(class_id, 'classificationItemId') and class_id.classificationItemId:
                            item_guid = class_id.classificationItemId.guid

                            # Lookup the classification item name
                            if item_guid in classification_item_lookup:
                                item_id, item_name = classification_item_lookup[item_guid]
                                # Format output based on what data is available
                                if item_id and item_name:
                                    print(
                                        f"  [{system_name}]: {item_id} - {item_name}")
                                elif item_id:
                                    print(f"  [{system_name}]: {item_id}")
                                elif item_name:
                                    print(f"  [{system_name}]: {item_name}")
                                else:
                                    print(
                                        f"  [{system_name}]: Assigned (ID: {item_guid})")
                            else:
                                print(
                                    f"  [{system_name}]: Assigned (ID: {item_guid})")
                        else:
                            print(f"  [{system_name}]: None")
                    else:
                        print(f"  [{system_name}]: None")
            else:
                for system in classification_systems:
                    print(f"  [{system.name}]: None")

        except Exception as e:
            print(f"Classifications: Unable to retrieve - {e}")
    else:
        print("Classifications: No classification systems available")

    # ========================================================================
    # GET PROPERTY VALUES
    # ========================================================================
    try:
        # Get all property values for this element
        all_values_dict = acu.GetPropertyValuesDictionary(
            [element_id], all_property_ids)

        if element_id in all_values_dict:
            props_dict = all_values_dict[element_id]

            print(f"\n{'-'*70}")
            print(f"PROPERTIES ({len(props_dict)} total)")
            print(f"{'-'*70}\n")

            # Separate properties with and without values
            filled_props = {}
            empty_props = []

            for prop_id, value in props_dict.items():
                # Get property name
                try:
                    prop_details = acc.GetDetailsOfProperties([prop_id])
                    if prop_details:
                        prop_name = prop_details[0].propertyDefinition.name
                    else:
                        prop_name = str(prop_id)
                except:
                    prop_name = str(prop_id)

                # Check if value is filled
                if value is not None and str(value).strip():
                    filled_props[prop_name] = value
                else:
                    empty_props.append(prop_name)

            # Display filled properties
            if filled_props:
                print("✓ Properties with values:")
                for prop_name in sorted(filled_props.keys()):
                    value = filled_props[prop_name]
                    # Truncate long values
                    value_str = str(value)
                    if len(value_str) > 50:
                        value_str = value_str[:47] + "..."
                    print(f"  {prop_name}: {value_str}")
            else:
                print("⚠️  No properties have values")

            # Show count of empty properties
            if empty_props:
                print(f"\n○ Empty properties: {len(empty_props)}")

                if SHOW_EMPTY_PROPERTIES:
                    if len(empty_props) <= 10:
                        print("  (Showing all):")
                        for prop_name in sorted(empty_props):
                            print(f"  • {prop_name}")
                    else:
                        print(f"  (Showing first 10):")
                        for prop_name in sorted(empty_props)[:10]:
                            print(f"  • {prop_name}")
                        print(f"  ... and {len(empty_props) - 10} more")
        else:
            print("\n⚠️  No properties available for this element")

    except Exception as e:
        print(f"\n✗ Error reading properties: {e}")

# ============================================================================
# FINAL SUMMARY
# ============================================================================
print("\n" + "="*70)
print("ANALYSIS COMPLETE")
print("="*70)
print(f"Total elements in project: {len(all_elements)}")
print(f"Elements analyzed: {len(elements_to_analyze)}")

if MAX_ELEMENTS_TO_ANALYZE and len(all_elements) > MAX_ELEMENTS_TO_ANALYZE:
    print(f"\n⚠️  Only first {MAX_ELEMENTS_TO_ANALYZE} elements were analyzed")
    print("   To analyze all elements, set MAX_ELEMENTS_TO_ANALYZE = None")

print("\n💡 TIPS:")
print("   - Use 106_get_elements_by_type.py to filter specific element types")
print("   - Use 202_set_property_value.py to modify property values")
print("   - Set SHOW_EMPTY_PROPERTIES = True to see empty property lists")
