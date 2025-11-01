"""
Simple example: Modify a project info field

This script shows how to modify ONE field in the simplest way possible.
"""

from archicad import ACConnection

# Connect to Archicad
conn = ACConnection.connect()
acc = conn.commands
act = conn.types


# Step 1: Get all available fields to see what we can modify
def show_available_fields():
    """Show all fields you can modify."""
    command_id = act.AddOnCommandId('TapirCommand', 'GetProjectInfoFields')
    response = acc.ExecuteAddOnCommand(command_id, {})

    fields = response.get('fields', [])

    print("\nAvailable fields you can modify:")
    print("="*60)
    for field in fields:
        field_id = field.get('projectInfoId', '')
        field_name = field.get('projectInfoName', '')
        field_value = field.get('projectInfoValue', '(empty)')

        print(f"\nField ID: '{field_id}'")
        print(f"  Name: {field_name}")
        print(f"  Current value: {field_value}")

    return fields


# Step 2: Modify a field (simple function)
def modify_field(field_id, new_value):
    """
    Modify a project info field.

    Args:
        field_id: The ID of the field (from show_available_fields)
        new_value: The new value you want to set
    """
    command_id = act.AddOnCommandId('TapirCommand', 'SetProjectInfoField')

    parameters = {
        'projectInfoId': field_id,
        'projectInfoValue': new_value
    }

    response = acc.ExecuteAddOnCommand(command_id, parameters)

    # Check if it worked
    if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
        print(
            f"✗ Failed: {response.get('error', {}).get('message', 'Unknown error')}")
        return False

    print(f"✓ Successfully set '{field_id}' to '{new_value}'")
    return True


# Main example
if __name__ == "__main__":
    print("\n" + "="*60)
    print("SIMPLE PROJECT INFO MODIFICATION EXAMPLE")
    print("="*60)

    # Show what fields are available
    fields = show_available_fields()

    # Example: Modify the Author field
    print("\n" + "="*60)
    print("EXAMPLE: Modify Author field")
    print("="*60)

    # Use the ID directly (language-independent!)
    # Look at the list above to find the ID you want to modify

    # Common field IDs (check your output above to confirm):
    # 'Author', 'Client', 'ProjectName', 'ProjectNumber', etc.

    print(f"\nChanging Author field (ID: 'Author') to 'John Doe'...")
    modify_field('Author', 'John Doe')

    # You can modify other fields the same way:
    modify_field('PROJECT_DESCRIPTION', 'Archicad Project')
    modify_field('PROJECTNAME', 'The Project')

    # Show the result
    print("\nVerifying change...")
    updated_fields = show_available_fields()

    print("\n" + "="*60)
    print("\n💡 To modify a different field:")
    print("   1. Look at the field IDs shown above")
    print("   2. Call: modify_field('FieldID', 'New Value')")
    print("\n   Example:")
    print("   modify_field('Client', 'ACME Corporation')")
    print("="*60 + "\n")
