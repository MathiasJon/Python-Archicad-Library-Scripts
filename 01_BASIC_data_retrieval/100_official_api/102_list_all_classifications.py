"""
================================================================================
SCRIPT: List All Classifications with Hierarchy
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Retrieval - Classifications

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Lists all classification systems and their complete hierarchical structure in
the current Archicad project. This script provides:
- Complete list of all classification systems
- Full hierarchical tree structure for each system
- Statistics: total items, maximum depth, item counts
- Clear visual representation of parent-child relationships
- Identification of all classifications (all levels can be assigned to elements)

Classification systems organize building elements according to industry
standards like Uniclass, Omniclass, or custom classification schemes.

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

- GetAllClassificationSystems()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationSystems
  Returns the list of all classification systems

- GetAllClassificationsInSystem(classificationSystemId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetAllClassificationsInSystem
  Returns all root classification items in the given system with their
  complete hierarchy including all children

[Data Types]
- ClassificationSystem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationSystem 
  Contains system name and ID

- ClassificationItemArrayItem
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.ClassificationItemArrayItem 
  Wrapper for classification items with hierarchy


--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. Run this script
3. Review the complete classification hierarchy output
4. Note that all classifications can be assigned to elements

No configuration needed - script displays all systems automatically.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows for each system:
- System name and GUID
- Total root classifications
- Complete hierarchical tree
- Visual indentation showing parent-child relationships
- Statistics: total items, maximum depth
- Summary of all systems

Example output:
  ════════════════════════════════════════════════════════════
  📋 SYSTEM 1: Uniclass 2015
  ════════════════════════════════════════════════════════════
  GUID: 12345678-1234-1234-1234-123456789abc
  Total root classifications: 8
  
  ────────────────────────────────────────────────────────────
  COMPLETE HIERARCHY:
  ────────────────────────────────────────────────────────────
  📦 Ss - Systems (level 1)
     └─ Ss_25 - Walls and barriers (level 2)
        └─ Ss_25_10 - Wall systems (level 3)
           └─ Ss_25_10_20 - External walls (level 4)
  
  Total items (including children): 156
  Maximum depth: 4 levels
  
  📊 Grand total: 156 classifications across all systems

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad (using the 
official Python palette or Tapir palette), the script will automatically 
connect to the FIRST instance of Archicad that was launched on your computer.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
CLASSIFICATION HIERARCHY:
- Classifications are organized in tree structures
- Parent classifications provide organizational structure
- All classifications (with or without children) can be assigned to elements
- Depth can vary from 1 to 6+ levels depending on system
- Each item has a unique ID and human-readable name

ASSIGNABILITY RULES:
✓ ALL classifications can be assigned to elements!
- Parent classifications: Can be assigned to elements
- Child classifications: Can be assigned to elements
- Example: Both "Ss - Systems" (parent) AND "Ss_25_10_20 - External walls"
           (child) can be assigned to elements
- Choose the appropriate level of detail for your needs

COMMON CLASSIFICATION SYSTEMS:
- Uniclass 2015: UK classification system
- Omniclass: North American system
- Custom systems: Project-specific classifications
- Multiple systems can coexist in one project

HIERARCHY STRUCTURE:
Level 1 (Root): Broad categories (e.g., "Systems", "Products")
Level 2: Sub-categories (e.g., "Walls and barriers")
Level 3-4+: Specific classifications (e.g., "External walls")

Each level provides more specific categorization.

STATISTICS PROVIDED:
- Total root classifications: Items at top level
- Total items including children: All items at all levels
- Maximum depth: Deepest nesting level in hierarchy
- Grand total: Sum across all classification systems

VISUAL INDICATORS:
- 📦 Root level items
- └─ Child items (indented by level)
- Level number shows depth in hierarchy

USE CASES:
- Understanding available classification systems
- Planning element classification strategy
- Identifying appropriate classifications for elements
- Documenting project classification structure
- Quality control of classification assignments

PERFORMANCE:
- Fast for small systems (< 100 items)
- May take a few seconds for large systems (1000+ items)
- Recursive traversal processes entire tree
- All systems loaded in single script run

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 102_get_classification_system_with_hierarchy.py (single system detail)
- 107_get_elements_by_classification.py (find classified elements)
- 201_set_element_classification.py (assign classifications)
- 111_get_element_full_info.py (view element classifications)

================================================================================
"""

from archicad import ACConnection

# =============================================================================
# CONNECT TO ARCHICAD
# =============================================================================

# Establish connection to running Archicad instance
conn = ACConnection.connect()
assert conn, "Cannot connect to Archicad"

# Initialize command and type objects
acc = conn.commands
act = conn.types


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def print_hierarchy(items, level=0):
    """
    Recursively print classification hierarchy with visual indentation.

    Args:
        items: List of ClassificationItemArrayItem objects
        level: Current depth level (0 = root)
    """
    for item in items:
        # Extract item details
        class_item = item.classificationItem
        item_id = class_item.id if hasattr(class_item, 'id') else 'No ID'
        item_name = class_item.name if hasattr(
            class_item, 'name') else 'No name'

        # Create indentation based on level
        indent = "   " * level

        # Choose symbol based on level
        if level == 0:
            symbol = "📦"
        else:
            symbol = "└─"

        # Check if item has children for recursive display
        has_children = False
        if hasattr(class_item, 'children') and class_item.children:
            has_children = True


        # Display the item
        print(
            f"{indent}{symbol} {item_id} - {item_name} (level {level + 1})")

        # Recursively process children if they exist
        if has_children:
            print_hierarchy(class_item.children, level + 1)


def count_all_items(items):
    """
    Recursively count all classification items including children.

    Args:
        items: List of ClassificationItemArrayItem objects

    Returns:
        Total count of items at all levels
    """
    count = len(items)
    for item in items:
        try:
            class_item = item.classificationItem
            if hasattr(class_item, 'children') and class_item.children:
                count += count_all_items(class_item.children)
        except (AttributeError, TypeError):
            pass
    return count


def get_max_depth(items, current_depth=1):
    """
    Recursively calculate maximum depth of classification hierarchy.

    Args:
        items: List of ClassificationItemArrayItem objects
        current_depth: Current depth level

    Returns:
        Maximum depth in the hierarchy
    """
    max_depth = current_depth
    for item in items:
        try:
            class_item = item.classificationItem
            if hasattr(class_item, 'children') and class_item.children:
                child_depth = get_max_depth(
                    class_item.children, current_depth + 1)
                max_depth = max(max_depth, child_depth)
        except (AttributeError, TypeError):
            pass
    return max_depth


# =============================================================================
# MAIN SCRIPT
# =============================================================================

def main():
    """Main function to list all classification systems with hierarchy."""

    print("\n" + "="*80)
    print("LIST ALL CLASSIFICATIONS WITH HIERARCHY v1.0")
    print("="*80)

    # -------------------------------------------------------------------------
    # 1. GET ALL CLASSIFICATION SYSTEMS
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("CLASSIFICATION SYSTEMS")
    print("="*80)

    try:
        classification_systems = acc.GetAllClassificationSystems()
    except Exception as e:
        print(f"\n✗ Error getting classification systems: {e}")
        import traceback
        traceback.print_exc()
        exit()

    if len(classification_systems) == 0:
        print("\n⚠️  No classification systems found in project")
        print("\nPossible reasons:")
        print("  • Project has no classification systems configured")
        print("  • Classification systems were not imported")
        print("\n💡 Tip: Import a classification system from Archicad Settings")
        exit()

    print(f"\n✓ Found {len(classification_systems)} classification system(s)")

    # -------------------------------------------------------------------------
    # 2. PROCESS EACH SYSTEM
    # -------------------------------------------------------------------------

    system_stats = []

    for sys_idx, system in enumerate(classification_systems, 1):
        system_name = system.name
        system_guid = system.classificationSystemId.guid

        print("\n" + "="*80)
        print(f"📋 SYSTEM {sys_idx}: {system_name}")
        print("="*80)
        print(f"GUID: {system_guid}")

        # Get all classifications in this system
        try:
            all_classifications = acc.GetAllClassificationsInSystem(
                system.classificationSystemId
            )

            num_root = len(all_classifications)
            print(f"Total root classifications: {num_root}")

            if num_root > 0:
                # Display complete hierarchy
                print("\n" + "─"*80)
                print("COMPLETE HIERARCHY:")
                print("─"*80)

                print_hierarchy(all_classifications)

                # Calculate statistics
                total_items = count_all_items(all_classifications)
                max_depth = get_max_depth(all_classifications)

                print(f"\n" + "─"*80)
                print(f"Statistics:")
                print(f"  • Total items (including children): {total_items}")
                print(f"  • Maximum depth: {max_depth} level(s)")
                print(f"  • Root items: {num_root}")

                # Store stats for summary
                system_stats.append({
                    'name': system_name,
                    'total': total_items,
                    'depth': max_depth
                })

            else:
                print("\n  (Empty system - no classifications)")
                system_stats.append({
                    'name': system_name,
                    'total': 0,
                    'depth': 0
                })

        except Exception as e:
            print(f"\n✗ Error loading classifications: {e}")
            import traceback
            traceback.print_exc()
            system_stats.append({
                'name': system_name,
                'total': 0,
                'depth': 0,
                'error': str(e)
            })

    # -------------------------------------------------------------------------
    # 3. SUMMARY
    # -------------------------------------------------------------------------

    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)

    grand_total = sum(stat['total'] for stat in system_stats)

    print(f"\nClassification systems breakdown:")
    for stat in system_stats:
        if 'error' in stat:
            print(f"  • {stat['name']}: Error - {stat['error']}")
        else:
            print(
                f"  • {stat['name']}: {stat['total']} items (depth: {stat['depth']})")

    print(f"\n📊 Grand total: {grand_total} classifications across all systems")
    print("\n" + "="*80)

    # -------------------------------------------------------------------------
    # NEXT STEPS
    # -------------------------------------------------------------------------

    print("\n💡 Understanding the output:")
    print("   • 📦 = Root level classification")
    print("   • └─ = Child classification (indented by level)")
    print("   • Level number = Depth in hierarchy (1 = root)")

    print("\n💡 Key points:")
    print("   • ALL classifications can be assigned to elements")
    print("   • Choose the appropriate level of detail for your project")
    print("   • Use exact IDs or names when assigning classifications")

    print("\n💡 Next steps:")
    print("   • Use 201_set_element_classification.py to assign classifications")
    print("   • Use 107_get_elements_by_classification.py to find classified elements")
    print("   • Use 102_get_classification_system_with_hierarchy.py for detailed view")


if __name__ == "__main__":
    main()