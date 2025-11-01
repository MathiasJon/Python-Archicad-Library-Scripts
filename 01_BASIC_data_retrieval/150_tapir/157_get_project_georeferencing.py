"""
================================================================================
SCRIPT: Get Project Georeferencing
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Tapir Add-On API
Category:       Project Information - Geolocation

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Retrieves comprehensive geographical location and coordinate reference system
information for the Archicad project. This script provides:
- Project geographical coordinates (latitude, longitude, altitude)
- True north direction
- Survey point position (eastings, northings, elevation)
- Coordinate Reference System (CRS) details
- Geodetic and vertical datum information
- Map projection and zone information
- Formatted coordinate display

Based on TAPIR API documentation v1.1.6

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- Tapir Add-On installed (required)
- An Archicad project must be open
- Geolocation configured in project (optional but recommended)

--------------------------------------------------------------------------------
API COMMANDS USED:
--------------------------------------------------------------------------------
[Tapir API]
- GetGeoLocation
  Returns project location, survey point, and CRS information

[Official API]
- ACConnection.connect()
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.ACConnection

- ExecuteAddOnCommand(addOnCommandId)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.ExecuteAddOnCommand

[Data Types]
- AddOnCommandId
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac24.html#archicad.releases.ac24.b2310types.AddOnCommandId

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project
2. (Optional) Configure geolocation: Options → Project Preferences → Location
3. Run this script
4. Review geographical and coordinate information

No configuration needed - script reads project geolocation settings.

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Console output shows:
- Project location (lat/long/alt)
- True north direction
- Survey point position
- Coordinate Reference System details
- Exported JSON data

Example output (With geolocation):
  ════════════════════════════════════════════════════════════
  PROJECT GEOLOCATION INFORMATION
  ════════════════════════════════════════════════════════════
  
  ────────────────────────────────────────────────────────────
  PROJECT LOCATION
  ────────────────────────────────────────────────────────────
  Geographic Coordinates: 48.856614°N, 2.352222°E, 35.00m
  True North Direction: 15.50° (from project north)
  
  ────────────────────────────────────────────────────────────
  SURVEY POINT POSITION
  ────────────────────────────────────────────────────────────
  Eastings:  600000.000 m
  Northings: 5400000.000 m
  Elevation: 35.000 m
  
  ────────────────────────────────────────────────────────────
  COORDINATE REFERENCE SYSTEM
  ────────────────────────────────────────────────────────────
  CRS Name:         WGS 84 / UTM zone 31N
  Geodetic Datum:   WGS 84
  Vertical Datum:   MSL
  Map Projection:   UTM
  Map Zone:         31N

Example output (No geolocation):
  ✗ No geolocation data available
  
  To set geolocation in Archicad:
    Options → Project Preferences → Location

--------------------------------------------------------------------------------
CONNECTION BEHAVIOR:
--------------------------------------------------------------------------------
⚠️  IMPORTANT: When running this script from outside Archicad, the script will
automatically connect to the FIRST instance of Archicad that was launched.

--------------------------------------------------------------------------------
NOTES:
--------------------------------------------------------------------------------
GEOLOCATION DATA:
- Project Location: Geographic coordinates (WGS84)
- Survey Point: Local coordinate system origin
- CRS: Coordinate Reference System for mapping
- True North: Rotation from project north

PROJECT LOCATION:
- Latitude: North-South position (-90 to +90 degrees)
- Longitude: East-West position (-180 to +180 degrees)
- Altitude: Height above sea level (meters)
- North: True north angle in radians (converted to degrees)

SURVEY POINT:
- Eastings: X coordinate in local CRS (meters)
- Northings: Y coordinate in local CRS (meters)
- Elevation: Z coordinate in local CRS (meters)
- Origin point for project coordinates

COORDINATE REFERENCE SYSTEMS:
- CRS Name: Formal system name (e.g., "WGS 84 / UTM zone 31N")
- Geodetic Datum: Reference ellipsoid (e.g., WGS 84, NAD83)
- Vertical Datum: Elevation reference (e.g., MSL, EGM96)
- Map Projection: Coordinate transformation (e.g., UTM, Lambert)
- Map Zone: Specific zone within projection system

USE CASES:
- Export project location for GIS integration
- Verify geolocation setup
- Coordinate system documentation
- IFC export with proper georeferencing
- Solar analysis preparation
- Site context integration

SETTING GEOLOCATION IN ARCHICAD:
1. Options → Project Preferences → Location
2. Enter latitude/longitude or address
3. Set true north direction
4. Configure survey point (optional)
5. Select coordinate reference system

EXPORT FUNCTIONALITY:
- Script exports data as JSON dictionary
- Useful for further processing
- Integration with other systems
- Coordinate transformation workflows

TROUBLESHOOTING:
- If no data: Set geolocation in Archicad
- If command fails: Ensure Tapir is installed
- If coordinates seem wrong: Verify CRS settings
- For IFC export: Geolocation is important

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 154_get_project_info.py (project information)
- 151_check_tapir_version.py (verify Tapir installation)

================================================================================
"""

from archicad import ACConnection
import math

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


def get_geolocation():
    """
    Get geolocation information for the project using TAPIR GetGeoLocation command.

    Returns:
        Dictionary with geolocation data
    """
    geo_info = {}

    try:
        # Create proper AddOnCommandId for GetGeoLocation
        command_id = act.AddOnCommandId('TapirCommand', 'GetGeoLocation')

        print("Retrieving geolocation data...")
        response = acc.ExecuteAddOnCommand(command_id, {})

        # Check for error response
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ Command failed: {error_msg}")
            geo_info['error'] = error_msg
            return geo_info

        print("✓ Successfully retrieved geolocation data\n")

        # Parse the response
        if isinstance(response, dict):
            geo_info = parse_geolocation_dict(response)
        elif hasattr(response, 'projectLocation'):
            geo_info = parse_geolocation_object(response)
        else:
            print(f"⚠ Unexpected response format: {type(response)}")
            geo_info['error'] = 'Unexpected response format'

    except Exception as e:
        print(f"✗ Error retrieving geolocation: {e}")
        import traceback
        traceback.print_exc()
        geo_info['error'] = str(e)

    return geo_info


def parse_geolocation_dict(response):
    """Parse geolocation data from dictionary response."""
    geo_info = {}

    # Project Location
    if 'projectLocation' in response:
        proj_loc = response['projectLocation']
        geo_info['projectLocation'] = {
            'longitude': proj_loc.get('longitude'),
            'latitude': proj_loc.get('latitude'),
            'altitude': proj_loc.get('altitude'),
            'north': proj_loc.get('north')  # in radians
        }

    # Survey Point
    if 'surveyPoint' in response:
        survey = response['surveyPoint']

        # Position
        position = survey.get('position', {})
        geo_info['surveyPoint'] = {
            'position': {
                'eastings': position.get('eastings'),
                'northings': position.get('northings'),
                'elevation': position.get('elevation')
            }
        }

        # Geo Referencing Parameters
        geo_params = survey.get('geoReferencingParameters', {})
        geo_info['surveyPoint']['geoReferencingParameters'] = {
            'crsName': geo_params.get('crsName', 'Not set'),
            'description': geo_params.get('description', 'Not set'),
            'geodeticDatum': geo_params.get('geodeticDatum', 'Not set'),
            'verticalDatum': geo_params.get('verticalDatum', 'Not set'),
            'mapProjection': geo_params.get('mapProjection', 'Not set'),
            'mapZone': geo_params.get('mapZone', 'Not set')
        }

    return geo_info


def parse_geolocation_object(response):
    """Parse geolocation data from object response."""
    geo_info = {}

    # Project Location
    if hasattr(response, 'projectLocation'):
        proj_loc = response.projectLocation
        geo_info['projectLocation'] = {
            'longitude': getattr(proj_loc, 'longitude', None),
            'latitude': getattr(proj_loc, 'latitude', None),
            'altitude': getattr(proj_loc, 'altitude', None),
            'north': getattr(proj_loc, 'north', None)
        }

    # Survey Point
    if hasattr(response, 'surveyPoint'):
        survey = response.surveyPoint

        # Position
        position = getattr(survey, 'position', None)
        if position:
            geo_info['surveyPoint'] = {
                'position': {
                    'eastings': getattr(position, 'eastings', None),
                    'northings': getattr(position, 'northings', None),
                    'elevation': getattr(position, 'elevation', None)
                }
            }

        # Geo Referencing Parameters
        geo_params = getattr(survey, 'geoReferencingParameters', None)
        if geo_params:
            if 'surveyPoint' not in geo_info:
                geo_info['surveyPoint'] = {}

            geo_info['surveyPoint']['geoReferencingParameters'] = {
                'crsName': getattr(geo_params, 'crsName', 'Not set'),
                'description': getattr(geo_params, 'description', 'Not set'),
                'geodeticDatum': getattr(geo_params, 'geodeticDatum', 'Not set'),
                'verticalDatum': getattr(geo_params, 'verticalDatum', 'Not set'),
                'mapProjection': getattr(geo_params, 'mapProjection', 'Not set'),
                'mapZone': getattr(geo_params, 'mapZone', 'Not set')
            }

    return geo_info


def radians_to_degrees(radians):
    """Convert radians to degrees."""
    if radians is None:
        return None
    return radians * 180.0 / math.pi


def format_coordinates(lat, lon, alt):
    """
    Format geographical coordinates for display.

    Args:
        lat: Latitude in decimal degrees
        lon: Longitude in decimal degrees
        alt: Altitude in meters

    Returns:
        Formatted coordinate string
    """
    if lat is None or lon is None:
        return "Not set"

    lat_dir = "N" if lat >= 0 else "S"
    lon_dir = "E" if lon >= 0 else "W"

    coord_str = f"{abs(lat):.6f}°{lat_dir}, {abs(lon):.6f}°{lon_dir}"

    if alt is not None:
        coord_str += f", {alt:.2f}m"

    return coord_str


def display_geolocation_info(geo_info):
    """
    Display geolocation information in a formatted way.

    Args:
        geo_info: Dictionary with geolocation data
    """
    print("\n" + "="*80)
    print("PROJECT GEOLOCATION INFORMATION")
    print("="*80 + "\n")

    if 'error' in geo_info:
        print(f"Error: {geo_info['error']}")
        print("\nNote: Make sure geolocation is set in Archicad:")
        print("  Options → Project Preferences → Location")
        return

    has_location = 'projectLocation' in geo_info
    has_survey = 'surveyPoint' in geo_info

    if not has_location and not has_survey:
        print("✗ No geolocation data available")
        print("\nTo set geolocation in Archicad:")
        print("  Options → Project Preferences → Location")
        return

    # Project Location
    if has_location:
        print("─"*80)
        print("PROJECT LOCATION")
        print("─"*80)

        proj_loc = geo_info['projectLocation']
        lat = proj_loc.get('latitude')
        lon = proj_loc.get('longitude')
        alt = proj_loc.get('altitude')
        north = proj_loc.get('north')

        print(f"Geographic Coordinates: {format_coordinates(lat, lon, alt)}")

        if north is not None:
            north_deg = radians_to_degrees(north)
            print(
                f"True North Direction: {north_deg:.2f}° (from project north)")

        print()

    # Survey Point
    if has_survey:
        survey = geo_info['surveyPoint']

        # Position
        if 'position' in survey:
            print("─"*80)
            print("SURVEY POINT POSITION")
            print("─"*80)

            pos = survey['position']
            eastings = pos.get('eastings')
            northings = pos.get('northings')
            elevation = pos.get('elevation')

            if eastings is not None:
                print(f"Eastings:  {eastings:.3f} m")
            if northings is not None:
                print(f"Northings: {northings:.3f} m")
            if elevation is not None:
                print(f"Elevation: {elevation:.3f} m")

            print()

        # Geo Referencing Parameters
        if 'geoReferencingParameters' in survey:
            print("─"*80)
            print("COORDINATE REFERENCE SYSTEM")
            print("─"*80)

            geo_params = survey['geoReferencingParameters']

            crs_name = geo_params.get('crsName', 'Not set')
            description = geo_params.get('description', 'Not set')
            geodetic_datum = geo_params.get('geodeticDatum', 'Not set')
            vertical_datum = geo_params.get('verticalDatum', 'Not set')
            map_projection = geo_params.get('mapProjection', 'Not set')
            map_zone = geo_params.get('mapZone', 'Not set')

            print(f"CRS Name:         {crs_name}")
            if description and description != 'Not set':
                print(f"Description:      {description}")
            print(f"Geodetic Datum:   {geodetic_datum}")
            print(f"Vertical Datum:   {vertical_datum}")
            print(f"Map Projection:   {map_projection}")
            if map_zone and map_zone != 'Not set':
                print(f"Map Zone:         {map_zone}")

    print("\n" + "="*80)


def export_to_dict(geo_info):
    """
    Export geolocation info to a simple dictionary format.
    Useful for further processing or export.
    """
    export_data = {}

    if 'projectLocation' in geo_info:
        proj_loc = geo_info['projectLocation']
        export_data['location'] = {
            'latitude': proj_loc.get('latitude'),
            'longitude': proj_loc.get('longitude'),
            'altitude': proj_loc.get('altitude'),
            'north_angle_degrees': radians_to_degrees(proj_loc.get('north'))
        }

    if 'surveyPoint' in geo_info:
        survey = geo_info['surveyPoint']
        if 'position' in survey:
            export_data['survey_point'] = survey['position']
        if 'geoReferencingParameters' in survey:
            export_data['coordinate_system'] = survey['geoReferencingParameters']

    return export_data


def main():
    """Main function to demonstrate the script."""

    # Get geolocation information
    geo_info = get_geolocation()

    # Display the information
    display_geolocation_info(geo_info)

    # Example: Export to dictionary
    if 'error' not in geo_info and ('projectLocation' in geo_info or 'surveyPoint' in geo_info):
        print("\n" + "─"*80)
        print("EXPORTED DATA (for further processing)")
        print("─"*80)

        export_data = export_to_dict(geo_info)

        import json
        print(json.dumps(export_data, indent=2, default=str))
        print()


if __name__ == "__main__":
    main()