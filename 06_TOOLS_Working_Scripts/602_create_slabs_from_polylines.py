"""
================================================================================
SCRIPT: Create Slabs from Polylines
================================================================================

Version:        1.0
Date:           2025-11-01
API Type:       Official Archicad API + Tapir Add-On
Category:       Element Creation

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Converts selected polylines into slabs in Archicad. Properly handles curved 
segments (arcs) from polylines and transfers them to the created slabs.

The script processes all selected polylines and creates slabs at a user-defined
Z level. Arc geometry is preserved during the conversion process.

Features:
- Batch conversion of multiple polylines
- Arc support (curved segments preserved)
- User-defined Z level for all slabs
- Automatic polygon closure handling
- Arc index validation and correction

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27)
- archicad-api package installed (pip install archicad)
- Tapir Add-On installed and active

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- ExecuteAddOnCommand()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Tapir Add-On Commands]
- GetSelectedElements
  Returns all currently selected elements in the project
  
- GetDetailsOfElements
  Returns detailed information about specified elements including coordinates
  
- CreateSlabs
  Creates slab elements from polygon data with arc support

[Data Types]
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AddOnCommandId

[Data Structure]
- Polyline coordinates: List of {x, y} points
- Arc data: {begIndex, endIndex, arcAngle}
- Slab data: {level, polygonCoordinates, polygonArcs, holes}

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Draw polylines in Archicad (can include curved segments)
2. Select the polylines you want to convert
3. Run this script
4. Enter the Z coordinate (level) for the slabs when prompted
5. Script creates slabs from selected polylines

--------------------------------------------------------------------------------
INPUT:
--------------------------------------------------------------------------------
- Selected polylines in Archicad
- Z coordinate (elevation) entered via dialog

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
- Slabs created at specified Z level
- Console log showing:
  * Number of polylines processed
  * Points and arcs for each polyline
  * Success/error messages for each slab
  * Summary of created slabs

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
- Polylines must have at least 3 distinct points
- Duplicate last points are automatically removed
- Arc indices are validated and corrected if needed
- Closed polylines are handled automatically
- If first and last points match, one is removed
- Arc geometry is preserved from polyline to slab

--------------------------------------------------------------------------------
ERROR HANDLING:
--------------------------------------------------------------------------------
- Validates user input for Z coordinate
- Checks for minimum 3 points per polygon
- Validates arc indices against point count
- Reports individual slab creation errors
- Provides detailed error messages via popup

================================================================================
"""

import tkinter as tk
from tkinter import messagebox, simpledialog
from archicad import ACConnection

# ============================================================================
# CONNECTION SETUP
# ============================================================================
# Establish connection with Archicad
conn = ACConnection.connect()
assert conn

# Create shortcuts for commands and types
acc = conn.commands
act = conn.types


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def execute_tapir_command(command_name, parameters=None):
    """
    Execute a Tapir Add-On command using official Archicad API

    Args:
        command_name (str): Name of the Tapir command
        parameters (dict): Command parameters

    Returns:
        dict: Command response or None if error
    """
    if parameters is None:
        parameters = {}

    try:
        command_result = acc.ExecuteAddOnCommand(
            act.AddOnCommandId('TapirCommand', command_name),
            parameters
        )
        return command_result
    except Exception as e:
        print(f"Error executing Tapir command '{command_name}': {str(e)}")
        return None


def show_popup(message, title="Script Message", message_type="error"):
    """
    Display a popup message to the user

    Args:
        message (str): Message to display
        title (str): Window title
        message_type (str): Type of message - 'error', 'info', or 'warning'
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    print(message)

    if message_type == "error":
        messagebox.showerror(title, message)
    elif message_type == "info":
        messagebox.showinfo(title, message)
    elif message_type == "warning":
        messagebox.showwarning(title, message)

    root.destroy()


def show_input_dialog(prompt, title="Input Needed", default="0"):
    """
    Display an input dialog to get user input

    Args:
        prompt (str): Prompt text to display
        title (str): Dialog title
        default (str): Default value

    Returns:
        str: User input or None if cancelled
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    response = simpledialog.askstring(title, prompt, initialvalue=default)
    root.destroy()
    return response


def validate_and_clean_polygon(coordinates, arcs):
    """
    Validate and clean polygon coordinates and arcs

    Args:
        coordinates (list): List of polygon coordinates
        arcs (list): List of arc definitions

    Returns:
        tuple: (cleaned_coordinates, cleaned_arcs, is_valid)
    """
    # Check minimum points requirement
    if len(coordinates) < 3:
        return coordinates, arcs, False

    # Handle polygon closure - remove duplicate last point if present
    first_point = coordinates[0]
    last_point = coordinates[-1]

    if (first_point['x'] == last_point['x'] and
        first_point['y'] == last_point['y'] and
            len(coordinates) > 3):

        original_last_index = len(coordinates) - 1
        coordinates = coordinates[:-1]  # Remove last point

        print(
            f"  Removed duplicate last point (was index {original_last_index})")

        # Update arc indices that reference the removed point
        for arc in arcs:
            if arc['endIndex'] == original_last_index:
                arc['endIndex'] = 0
                print(
                    f"  Updated arc endIndex from {original_last_index} to 0")
            if arc['begIndex'] == original_last_index:
                arc['begIndex'] = 0
                print(
                    f"  Updated arc begIndex from {original_last_index} to 0")

    # Validate arc indices
    max_index = len(coordinates) - 1
    valid_arcs = []

    for arc in arcs:
        if (0 <= arc['begIndex'] <= max_index and
                0 <= arc['endIndex'] <= max_index):
            valid_arcs.append(arc)
        else:
            print(f"  Warning: Removed invalid arc with indices "
                  f"{arc['begIndex']}-{arc['endIndex']} (max valid: {max_index})")

    # Final validation
    is_valid = len(coordinates) >= 3

    return coordinates, valid_arcs, is_valid


# ============================================================================
# STEP 1: GET SELECTED ELEMENTS
# ============================================================================
print("\n" + "="*70)
print("STEP 1: RETRIEVING SELECTED ELEMENTS")
print("="*70)

selectedElementsResponse = execute_tapir_command('GetSelectedElements', {})

if not selectedElementsResponse:
    show_popup(
        "Failed to get selected elements.\nMake sure Tapir Add-On is installed.",
        "Connection Error",
        "error"
    )
    exit(1)

selectedElements = selectedElementsResponse.get('elements', [])

if not selectedElements:
    show_popup(
        "No elements selected!\n\nPlease select at least one polyline.",
        "Selection Error",
        "error"
    )
    exit(1)

print(f"Found {len(selectedElements)} selected elements")


# ============================================================================
# STEP 2: GET ELEMENT DETAILS
# ============================================================================
print("\n" + "="*70)
print("STEP 2: RETRIEVING ELEMENT DETAILS")
print("="*70)

detailsOfElementsResponse = execute_tapir_command(
    'GetDetailsOfElements',
    {'elements': selectedElements}
)

if not detailsOfElementsResponse:
    show_popup(
        "Could not retrieve details of selected elements!",
        "Processing Error",
        "error"
    )
    exit(1)

detailsOfElements = detailsOfElementsResponse.get('detailsOfElements', [])

if not detailsOfElements:
    show_popup(
        "Could not retrieve details of selected elements!",
        "Processing Error",
        "error"
    )
    exit(1)


# ============================================================================
# STEP 3: FILTER POLYLINES
# ============================================================================
print("\n" + "="*70)
print("STEP 3: FILTERING POLYLINES")
print("="*70)

polylines = []

for i in range(len(selectedElements)):
    if detailsOfElements[i]['type'] == 'PolyLine':
        polylines.append((selectedElements[i], detailsOfElements[i]))

if len(polylines) == 0:
    show_popup(
        "No polylines found in the selection!\n\nPlease select at least one polyline.",
        "Polyline Error",
        "error"
    )
    exit(1)

print(f"Found {len(polylines)} polylines to convert")


# ============================================================================
# STEP 4: GET USER INPUT FOR Z LEVEL
# ============================================================================
print("\n" + "="*70)
print("STEP 4: GETTING Z LEVEL FROM USER")
print("="*70)

try:
    level_z = show_input_dialog(
        "Enter the Z coordinate (level) for the slabs:",
        "Slab Configuration",
        "0"
    )

    if level_z is None:  # User clicked Cancel
        print("Script cancelled by user")
        exit(0)

    # Validate input
    level_z = float(level_z)
    print(f"Using Z level: {level_z}")

except ValueError:
    show_popup(
        "Please enter a valid number for the Z coordinate.\nScript will now exit.",
        "Invalid Input",
        "error"
    )
    exit(1)


# ============================================================================
# STEP 5: PROCESS POLYLINES AND PREPARE SLAB DATA
# ============================================================================
print("\n" + "="*70)
print("STEP 5: PROCESSING POLYLINES")
print("="*70)

slabsData = []

for i, (polyline, details) in enumerate(polylines):
    print(f"\nProcessing polyline {i+1} of {len(polylines)}")

    # Extract coordinates from polyline
    polyline_points = details['details']['coordinates']
    polyline_arcs = details['details'].get('arcs', [])

    print(f"  Points: {len(polyline_points)}")
    print(f"  Arcs: {len(polyline_arcs)}")

    # Convert coordinates to expected format
    polygonCoordinates = [
        {'x': point['x'], 'y': point['y']}
        for point in polyline_points
    ]

    # Convert arcs to expected format
    polygonArcs = [
        {
            'begIndex': arc['begIndex'],
            'endIndex': arc['endIndex'],
            'arcAngle': arc['arcAngle']
        }
        for arc in polyline_arcs
    ]

    # Validate and clean polygon
    polygonCoordinates, polygonArcs, is_valid = validate_and_clean_polygon(
        polygonCoordinates,
        polygonArcs
    )

    if not is_valid:
        print(f"  Skipping polyline {i+1}: Not enough valid points")
        continue

    print(
        f"  Final polygon: {len(polygonCoordinates)} points, {len(polygonArcs)} arcs")

    # Create slab data
    slabData = {
        'level': level_z,
        'polygonCoordinates': polygonCoordinates,
        'polygonArcs': polygonArcs,
        'holes': []
    }

    if polygonArcs:
        print(f"  ✓ {len(polygonArcs)} arcs will be transferred to slab")

    slabsData.append(slabData)


# ============================================================================
# STEP 6: CREATE SLABS
# ============================================================================
print("\n" + "="*70)
print("STEP 6: CREATING SLABS")
print("="*70)

try:
    # Format command data
    slabs_command_data = {'slabsData': slabsData}

    print("\nCommand: CreateSlabs")
    print(f"Creating {len(slabsData)} slabs...")

    # Execute command
    createSlabsResponse = execute_tapir_command(
        'CreateSlabs', slabs_command_data)

    # Validate response
    if createSlabsResponse is None:
        show_popup(
            "Error: Received null response from CreateSlabs command.\n"
            "Check script logs for details.",
            "Command Error",
            "error"
        )
        exit(1)

    createdElements = createSlabsResponse.get('elements', [])

    if not createdElements:
        show_popup(
            "No slabs were created.\nCheck the polyline geometry and try again.",
            "Operation Result",
            "warning"
        )
        exit(1)

    # Process results
    success_count = 0
    error_messages = []

    for i, element in enumerate(createdElements):
        if 'error' in element:
            error_code = element['error'].get('code', 'unknown')
            error_msg = element['error'].get('message', 'Unknown error')
            error_messages.append(
                f"Polyline {i+1}: {error_msg} (code: {error_code})"
            )
        else:
            success_count += 1
            if 'elementId' in element:
                print(
                    f"  ✓ Slab {i+1} created - GUID: {element['elementId']['guid']}")

    # Report results
    if error_messages:
        error_text = "\n".join(error_messages)
        print(f"\nErrors encountered:\n{error_text}")

        if success_count == 0:
            show_popup(
                f"Failed to create any slabs.\n\nErrors:\n{error_text}",
                "Operation Failed",
                "error"
            )
            exit(1)
        else:
            show_popup(
                f"Created {success_count} slabs, but encountered "
                f"{len(error_messages)} errors:\n\n{error_text}",
                "Partial Success",
                "warning"
            )
    else:
        print(f"\n✓ Successfully created {success_count} slabs")

except Exception as e:
    show_popup(
        f"Error creating slabs:\n\n{str(e)}",
        "Processing Error",
        "error"
    )
    print(f"Error details: {str(e)}")
    exit(1)


# ============================================================================
# SUMMARY
# ============================================================================
print("\n" + "="*70)
print("SUMMARY")
print("="*70)

if success_count > 0:
    arc_count = sum(len(slab['polygonArcs']) for slab in slabsData)

    print(f"Successfully created: {success_count} slabs")
    print(f"From polylines: {len(slabsData)}")
    print(f"Total arcs transferred: {arc_count}")
    print(f"Z Level: {level_z}")

    message = (
        f"Successfully created {success_count} slabs from "
        f"{len(slabsData)} polylines.\n\n"
        f"Z Level: {level_z}"
    )

    show_popup(message, "Operation Complete", "info")
else:
    show_popup(
        "No slabs were created.\nPlease check your polylines and try again.",
        "Operation Failed",
        "error"
    )
