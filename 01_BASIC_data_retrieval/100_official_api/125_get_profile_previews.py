"""
================================================================================
SCRIPT: Get Profile Previews (Complete)
================================================================================

Version:        1.0
Date:           2025-10-30
API Type:       Official Archicad API
Category:       Data Retrieval - Attributes

--------------------------------------------------------------------------------
DESCRIPTION:
--------------------------------------------------------------------------------
Generates preview images (PNG) of all profile attributes in the current Archicad 
project. Each profile is rendered as a high-quality image with customizable 
dimensions and background color.

This script provides comprehensive profile visualization:
- Retrieves all profile attributes from the project
- Generates preview images for each profile
- Saves images as individual PNG files
- Creates an interactive HTML gallery for easy viewing
- Optionally exports profile metadata to JSON format
- Provides progress tracking during generation

--------------------------------------------------------------------------------
REQUIREMENTS:
--------------------------------------------------------------------------------
- Archicad instance running (Tested on Archicad 27+)
- archicad package installed (pip install archicad)
- No additional Python packages required (uses standard library)

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

- GetProfileAttributePreview(attributeIds, imageWidth, imageHeight, backgroundColor)
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.html#archicad.Commands.GetProfileAttributePreview
  Returns preview images of requested profile attributes in base64 string format.
  Parameters:
  * attributeIds: List of profile attribute identifiers
  * imageWidth: Width of the preview image in pixels
  * imageHeight: Height of the preview image in pixels
  * backgroundColor: RGBColor object for image background (0.0-1.0 range)
  Returns:
  * List of Image objects with base64 encoded PNG data

[Data Types]
- ProfileAttribute
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.ProfileAttribute

- RGBColor
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.RGBColor
  Color model with red, green, and blue components
  CRITICAL: Values must be in 0.0-1.0 range (NOT 0-255!)

- Image
  https://archicadapi.graphisoft.com/archicadPythonPackage/archicad.releases.ac25.html#archicad.releases.ac25.b3000types.Image
  Image encoded as base64 string

--------------------------------------------------------------------------------
USAGE:
--------------------------------------------------------------------------------
1. Open an Archicad project containing profiles
2. Optionally configure settings in CONFIGURATION section below
3. Run this script
4. Preview images are saved in 'profile_previews/' folder
5. Open 'profile_previews/index.html' in browser to view gallery

Customization Options:
- IMAGE_WIDTH: Preview width in pixels (default: 400)
- IMAGE_HEIGHT: Preview height in pixels (default: 300)
- BACKGROUND_COLOR: RGB tuple with 0.0-1.0 values (default: white)
- OUTPUT_DIR: Output folder name (default: "profile_previews")
- EXPORT_JSON: Export profile metadata to JSON (default: False)

--------------------------------------------------------------------------------
OUTPUT:
--------------------------------------------------------------------------------
Creates a complete folder structure:

profile_previews/
  ├── index.html                    (Interactive HTML gallery)
  ├── profile_001_WallProfile.png   (Individual profile images)
  ├── profile_002_BeamProfile.png
  ├── profile_003_CustomShape.png
  └── ...

Optional (when EXPORT_JSON = True):
  └── profiles.json                 (Profile metadata in JSON format)

Console output shows:
- Total number of profiles found
- Generation progress for each profile
- Success/failure status for each image
- Final summary with file paths

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
RGB COLOR VALUES - CRITICAL INFORMATION:
⚠️  Archicad API uses RGB values between 0.0 and 1.0 (NOT the standard 0-255!)

Correct color values:
- White:      RGBColor(1.0, 1.0, 1.0)
- Black:      RGBColor(0.0, 0.0, 0.0)
- Light gray: RGBColor(0.95, 0.95, 0.95)
- Light blue: RGBColor(0.9, 0.9, 1.0)
- Red:        RGBColor(1.0, 0.0, 0.0)

To convert from standard 0-255 RGB:
- Divide each value by 255.0
- Example: RGB(230, 230, 230) → RGBColor(230/255.0, 230/255.0, 230/255.0)
           = RGBColor(0.902, 0.902, 0.902)

IMAGE FORMAT:
- Images are returned as base64 encoded PNG data
- Automatically decoded and saved as standard PNG files
- Filenames are sanitized (spaces replaced with underscores)
- Invalid filename characters are removed
- Sequential numbering ensures unique filenames

PERFORMANCE CONSIDERATIONS:
- Generating previews can be slow for projects with many profiles
- Time increases with larger image dimensions
- Progress is displayed for each profile being processed
- Consider using smaller dimensions (e.g., 200x150) for faster generation
- Typical speed: 1-3 seconds per profile depending on complexity

HTML GALLERY FEATURES:
- Responsive grid layout adapts to screen size
- Hover effects for better interactivity
- Shows profile name, ID, and GUID
- Images maintain aspect ratio
- Modern, clean design
- Works in all modern browsers

--------------------------------------------------------------------------------
RELATED SCRIPTS:
--------------------------------------------------------------------------------
- 122_get_profiles.py (basic profile list without images)
- 114_01_get_building_materials.py (materials used in profiles)

================================================================================
"""

from archicad import ACConnection
import base64
import os
import json
import re

# =============================================================================
# CONFIGURATION
# =============================================================================

# Image dimensions (pixels)
# Recommended values: 200-600 for width, 150-450 for height
IMAGE_WIDTH = 400
IMAGE_HEIGHT = 300

# Background color (RGB values in 0.0-1.0 range)
# CRITICAL: Do NOT use 0-255 range!
BACKGROUND_COLOR = (1.0, 1.0, 1.0)  # White background

# Alternative background colors (all values 0.0-1.0):
# BACKGROUND_COLOR = (0.95, 0.95, 0.95)  # Light gray
# BACKGROUND_COLOR = (0.0, 0.0, 0.0)      # Black
# BACKGROUND_COLOR = (0.9, 0.9, 1.0)      # Light blue
# BACKGROUND_COLOR = (1.0, 0.95, 0.9)     # Warm white

# Output directory for preview images
OUTPUT_DIR = "profile_previews"

# Export profile metadata to JSON file
# Set to True to create profiles.json with metadata
EXPORT_JSON = False


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

def sanitize_filename(name):
    """
    Sanitize filename by removing invalid characters.

    Removes or replaces characters that are invalid in filenames:
    - Windows: < > : " / \\ | ? *
    - Also replaces spaces with underscores for consistency

    Args:
        name: Original filename string

    Returns:
        Sanitized filename safe for all operating systems
    """
    # Remove invalid filename characters
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    return name


# =============================================================================
# GET PROFILES
# =============================================================================

def get_all_profiles():
    """
    Retrieve all profile attributes from the current project.

    Returns:
        Tuple of (profiles list, attribute IDs list)
        Returns ([], []) if no profiles found or error occurs
    """
    try:
        print("\nRetrieving profiles from project...")

        # Get all profile attribute IDs
        attribute_ids = acc.GetAttributesByType('Profile')
        print(f"✓ Found {len(attribute_ids)} profile(s)")

        if not attribute_ids:
            print("\n⚠️  No profiles found in the project")
            return [], []

        # Get detailed profile attributes
        profiles = acc.GetProfileAttributes(attribute_ids)

        return profiles, attribute_ids

    except Exception as e:
        print(f"\n✗ Error getting profiles: {e}")
        import traceback
        traceback.print_exc()
        return [], []


# =============================================================================
# GENERATE PREVIEW IMAGES
# =============================================================================

def generate_profile_previews():
    """
    Generate preview images for all profiles in the project.

    This function:
    1. Creates output directory if needed
    2. Retrieves all profiles
    3. Generates preview images via API
    4. Decodes base64 image data
    5. Saves images as PNG files
    6. Creates HTML gallery
    7. Optionally exports metadata to JSON

    Returns:
        List of dictionaries containing profile metadata and filenames
        Returns [] if error occurs
    """
    try:
        # Create output directory if it doesn't exist
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            print(f"✓ Created output directory: {OUTPUT_DIR}/")

        # Retrieve all profiles from project
        profiles, attribute_ids = get_all_profiles()

        if not profiles or not attribute_ids:
            print("\n⚠️  No profiles to process")
            return []

        # Prepare background color for API call
        bg_r, bg_g, bg_b = BACKGROUND_COLOR

        # Validate color values are in correct range (0.0-1.0)
        if not (0.0 <= bg_r <= 1.0 and 0.0 <= bg_g <= 1.0 and 0.0 <= bg_b <= 1.0):
            print(f"\n⚠️  WARNING: Background color values must be between 0.0 and 1.0!")
            print(f"   Current values: R={bg_r}, G={bg_g}, B={bg_b}")
            print(f"   Using white background instead.")
            bg_r, bg_g, bg_b = 1.0, 1.0, 1.0

        # Create RGBColor object for API
        bg_rgb = act.RGBColor(bg_r, bg_g, bg_b)

        print(f"\nGenerating preview images...")
        print(f"  Image size: {IMAGE_WIDTH}x{IMAGE_HEIGHT} pixels")
        print(f"  Background: RGB({bg_r:.2f}, {bg_g:.2f}, {bg_b:.2f})")
        print(f"  Output directory: {OUTPUT_DIR}/")
        print("-"*80)

        # Call API to generate all preview images
        # This may take some time depending on number and complexity of profiles
        images = acc.GetProfileAttributePreview(
            attribute_ids,
            IMAGE_WIDTH,
            IMAGE_HEIGHT,
            bg_rgb
        )

        # Storage for profile metadata
        profile_data = []

        # Process and save each profile image
        saved_count = 0
        for idx, (profile_or_error, image_or_error) in enumerate(zip(profiles, images), 1):
            # Extract profile information
            if hasattr(profile_or_error, 'profileAttribute') and profile_or_error.profileAttribute:
                profile = profile_or_error.profileAttribute
                name = profile.name if hasattr(
                    profile, 'name') else f'Profile_{idx}'
                guid = str(profile.attributeId.guid) if hasattr(
                    profile, 'attributeId') else 'unknown'
            else:
                # Fallback for profiles without attributes
                name = f'Profile_{idx}'
                guid = 'unknown'

            # Extract and save image data
            if hasattr(image_or_error, 'image') and image_or_error.image:
                image = image_or_error.image

                try:
                    # Decode base64 encoded PNG data
                    img_data = base64.b64decode(image.content)

                    # Create sanitized filename
                    safe_name = sanitize_filename(name)
                    filename = f"profile_{idx:03d}_{safe_name}.png"
                    filepath = os.path.join(OUTPUT_DIR, filename)

                    # Write PNG file to disk
                    with open(filepath, 'wb') as f:
                        f.write(img_data)

                    print(f"  {idx:3}. ✓ {name} → {filename}")
                    saved_count += 1

                    # Store profile metadata for HTML gallery and JSON
                    profile_data.append({
                        'id': idx,
                        'name': name,
                        'guid': guid,
                        'filename': filename
                    })

                except Exception as e:
                    print(f"  {idx:3}. ✗ {name} - Error saving: {e}")
            else:
                # Handle case where image generation failed
                error_msg = image_or_error.error if hasattr(
                    image_or_error, 'error') else 'No image data'
                print(f"  {idx:3}. ✗ {name} - {error_msg}")

        print("-"*80)
        print(
            f"✓ Successfully saved {saved_count}/{len(profiles)} preview images")

        # Export metadata to JSON if requested
        if EXPORT_JSON and profile_data:
            json_path = os.path.join(OUTPUT_DIR, 'profiles.json')
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(profile_data, f, indent=2, ensure_ascii=False)
            print(f"✓ Exported profile metadata to: {json_path}")

        # Generate interactive HTML gallery
        generate_html_gallery(profile_data)

        return profile_data

    except Exception as e:
        print(f"\n✗ Error generating previews: {e}")
        import traceback
        traceback.print_exc()
        return []


# =============================================================================
# GENERATE HTML GALLERY
# =============================================================================

def generate_html_gallery(profile_data):
    """
    Generate an interactive HTML gallery to view all profile previews.

    Creates a modern, responsive HTML page with:
    - Grid layout that adapts to screen size
    - Hover effects for better UX
    - Profile names, IDs, and GUIDs
    - Clean, professional design

    Args:
        profile_data: List of dictionaries with profile metadata
    """
    try:
        html_path = os.path.join(OUTPUT_DIR, 'index.html')

        # HTML template with embedded CSS
        html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Archicad Profile Previews</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: #f5f5f5;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
        }
        
        h1 {
            color: #333;
            margin-bottom: 10px;
        }
        
        .info {
            color: #666;
            margin-bottom: 30px;
            font-size: 14px;
        }
        
        .gallery {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
        }
        
        .profile-card {
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            transition: transform 0.2s, box-shadow 0.2s;
        }
        
        .profile-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.15);
        }
        
        .profile-image {
            width: 100%;
            height: 200px;
            object-fit: contain;
            background: #fafafa;
            padding: 10px;
        }
        
        .profile-info {
            padding: 15px;
        }
        
        .profile-name {
            font-weight: 600;
            color: #333;
            margin-bottom: 5px;
            font-size: 16px;
        }
        
        .profile-id {
            color: #999;
            font-size: 12px;
            font-family: monospace;
        }
        
        .profile-guid {
            color: #999;
            font-size: 11px;
            font-family: monospace;
            margin-top: 5px;
            word-break: break-all;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📐 Archicad Profile Previews</h1>
        <div class="info">
            Total profiles: {count} | Generated: {date}
        </div>
        
        <div class="gallery">
"""

        # Add individual profile cards
        for profile in profile_data:
            html_content += f"""
            <div class="profile-card">
                <img src="{profile['filename']}" 
                     alt="{profile['name']}" 
                     class="profile-image">
                <div class="profile-info">
                    <div class="profile-name">{profile['name']}</div>
                    <div class="profile-id">Profile #{profile['id']}</div>
                    <div class="profile-guid">{profile['guid']}</div>
                </div>
            </div>
"""

        # Close HTML structure
        html_content += """
        </div>
    </div>
</body>
</html>
"""

        # Replace template placeholders
        from datetime import datetime
        html_content = html_content.replace('{count}', str(len(profile_data)))
        html_content = html_content.replace(
            '{date}', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

        # Write HTML file
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # Display absolute path for easy access
        abs_path = os.path.abspath(html_path)
        print(f"✓ Generated HTML gallery: {abs_path}")
        print(f"\n💡 Open this file in your browser to view all profiles:")
        print(f"   file://{abs_path}")

    except Exception as e:
        print(f"\n✗ Error generating HTML gallery: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("GET PROFILE PREVIEWS v1.0")
    print("="*80)

    print("\n" + "="*80)
    print("PROFILE PREVIEW GENERATOR")
    print("="*80)

    # Generate preview images for all profiles
    profile_data = generate_profile_previews()

    # Display summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"  Total profiles processed: {len(profile_data)}")
    print(f"  Output directory: {os.path.abspath(OUTPUT_DIR)}/")
    if profile_data:
        print(
            f"  HTML gallery: {os.path.abspath(os.path.join(OUTPUT_DIR, 'index.html'))}")
    print("\n" + "="*80)

    # Usage hints
    if profile_data:
        print("\n💡 Next steps:")
        print("   1. Open the HTML gallery in your browser")
        print("   2. Review the generated preview images")
        print("   3. Use profile GUIDs in other scripts if needed")


if __name__ == "__main__":
    main()
