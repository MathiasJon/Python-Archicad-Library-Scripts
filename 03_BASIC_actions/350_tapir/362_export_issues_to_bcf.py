"""
Export Issues to BCF (BIM Collaboration Format)

This script exports project issues to BCF format using TAPIR commands.
BCF is the standard format for BIM issue tracking and collaboration.

Uses TAPIR commands:
- GetIssues (v1.0.2): Get all issues
- ExportIssuesToBCF (v1.0.6): Export to BCF file

Based on TAPIR API documentation.
"""

from archicad import ACConnection
import os
from datetime import datetime

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


# =============================================================================
# GET ISSUES
# =============================================================================

def get_all_issues():
    """
    Get all issues from the project.
    Uses TAPIR GetIssues command.

    Returns:
        List of issue dictionaries
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'GetIssues')

        parameters = {}

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ GetIssues failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'issues' in response:
            return response['issues']

        return []

    except Exception as e:
        print(f"✗ Error getting issues: {e}")
        return []


# =============================================================================
# EXPORT TO BCF
# =============================================================================

def export_issues_to_bcf(issues_to_export, export_path, use_external_id=True, align_by_survey_point=False):
    """
    Export issues to a BCF file.
    Uses TAPIR ExportIssuesToBCF command.

    Args:
        issues_to_export: List of issue objects (from GetIssues) OR list of issue GUIDs
        export_path: Full path to BCF file (including filename .bcf or .bcfzip)
        use_external_id: Use external IFC ID (default: True)
        align_by_survey_point: Align by survey point vs project origin (default: False)

    Returns:
        Boolean indicating success
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'ExportIssuesToBCF')

        # Prepare issues list
        # If we receive issue objects (dicts with 'issueId' key), pass them directly
        # If we receive GUIDs (strings), wrap them
        if issues_to_export and isinstance(issues_to_export[0], dict):
            # Full issue objects - pass directly (like official script)
            issues = issues_to_export
        else:
            # Just GUIDs - wrap them
            issues = [{'issueId': {'guid': guid}} for guid in issues_to_export]

        parameters = {
            'issues': issues,
            'exportPath': export_path,
            'useExternalId': use_external_id,
            'alignBySurveyPoint': align_by_survey_point
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'success' in response:
            if response['success']:
                print(f"✓ Exported {len(issues)} issue(s) to BCF")
                print(f"  File: {export_path}")
                return True
            else:
                error = response.get('error', {})
                print(
                    f"✗ ExportIssuesToBCF failed: {error.get('message', 'Unknown error')}")
                return False

        return False

    except Exception as e:
        print(f"✗ Error exporting to BCF: {e}")
        import traceback
        traceback.print_exc()
        return False


def export_all_issues_to_bcf(export_path, use_external_id=True, align_by_survey_point=False):
    """
    Export all issues to BCF.

    Args:
        export_path: Full path to BCF file
        use_external_id: Use external IFC ID (default: True)
        align_by_survey_point: Align by survey point (default: False)

    Returns:
        Boolean indicating success
    """
    issues = get_all_issues()

    if not issues:
        print("⚠ No issues to export")
        return False

    print(f"Exporting {len(issues)} issue(s)...")

    # Pass full issue objects (like official script)
    return export_issues_to_bcf(issues, export_path, use_external_id, align_by_survey_point)


def export_issues_by_filter(filter_function, export_path, use_external_id=True, align_by_survey_point=False):
    """
    Export issues that match a filter function.

    Args:
        filter_function: Function that takes an issue dict and returns True/False
        export_path: Full path to BCF file
        use_external_id: Use external IFC ID (default: True)
        align_by_survey_point: Align by survey point (default: False)

    Returns:
        Boolean indicating success
    """
    issues = get_all_issues()

    if not issues:
        print("⚠ No issues found")
        return False

    # Filter issues - keep full objects
    filtered_issues = [issue for issue in issues if filter_function(issue)]

    if not filtered_issues:
        print("⚠ No issues match the filter")
        return False

    print(f"Exporting {len(filtered_issues)} filtered issue(s)...")

    # Pass full issue objects (like official script)
    return export_issues_to_bcf(filtered_issues, export_path, use_external_id, align_by_survey_point)


def export_issues_by_tag(tag_text, export_path, use_external_id=True, align_by_survey_point=False):
    """
    Export issues that contain specific tag text.

    Args:
        tag_text: Tag text to search for (case-insensitive substring match)
        export_path: Full path to BCF file
        use_external_id: Use external IFC ID
        align_by_survey_point: Align by survey point

    Returns:
        Boolean indicating success
    """
    def has_tag(issue):
        return tag_text.lower() in issue.get('tagText', '').lower()

    return export_issues_by_filter(has_tag, export_path, use_external_id, align_by_survey_point)


def export_recent_issues(days=7, export_path="recent_issues.bcf", use_external_id=True):
    """
    Export issues modified in the last N days.

    Args:
        days: Number of days (default: 7)
        export_path: Full path to BCF file
        use_external_id: Use external IFC ID

    Returns:
        Boolean indicating success
    """
    from datetime import timedelta

    cutoff_time = datetime.now() - timedelta(days=days)
    cutoff_timestamp = cutoff_time.timestamp()

    def is_recent(issue):
        modi_time = issue.get('modiTime', 0)
        return modi_time > cutoff_timestamp

    print(f"Exporting issues modified in the last {days} day(s)...")

    return export_issues_by_filter(is_recent, export_path, use_external_id)


# =============================================================================
# SUMMARY & REPORTING
# =============================================================================

def create_bcf_export_summary(issues, export_path):
    """
    Create a text summary of exported issues.

    Args:
        issues: List of issue dictionaries
        export_path: BCF file path (summary will be created alongside)
    """
    summary_path = export_path.replace(
        '.bcf', '_summary.txt').replace('.bcfzip', '_summary.txt')

    try:
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("BCF EXPORT SUMMARY\n")
            f.write("="*80 + "\n\n")
            f.write(
                f"Export Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"BCF File: {os.path.basename(export_path)}\n")
            f.write(f"Total Issues: {len(issues)}\n\n")

            # Group by tags
            all_tags = {}
            for issue in issues:
                tags = issue.get('tagText', '')
                if tags:
                    for tag in tags.split(','):
                        tag = tag.strip()
                        if tag:
                            all_tags[tag] = all_tags.get(tag, 0) + 1

            if all_tags:
                f.write("Issues by Tag:\n")
                f.write("-"*80 + "\n")
                for tag, count in sorted(all_tags.items(), key=lambda x: x[1], reverse=True):
                    f.write(f"  {tag}: {count}\n")
                f.write("\n")

            # Issue list
            f.write("Issue List:\n")
            f.write("-"*80 + "\n")
            for idx, issue in enumerate(issues, 1):
                name = issue.get('name', 'Unnamed')
                guid = issue.get('issueId', {}).get('guid', 'N/A')
                tags = issue.get('tagText', 'No tags')
                crea_time = issue.get('creaTime', 0)

                f.write(f"\n{idx}. {name}\n")
                f.write(f"   GUID: {guid}\n")
                f.write(f"   Tags: {tags}\n")
                f.write(
                    f"   Created: {datetime.fromtimestamp(crea_time) if crea_time > 0 else 'N/A'}\n")

        abs_path = os.path.abspath(summary_path)
        print(f"\n✓ Export summary created:")
        print(f"  {abs_path}")

    except Exception as e:
        print(f"✗ Error creating summary: {e}")


def display_issues_summary(issues):
    """
    Display a summary of issues.

    Args:
        issues: List of issue dictionaries
    """
    if not issues:
        print("  No issues")
        return

    print(f"\nTotal issues: {len(issues)}")
    print("="*80)

    # Group by tags
    all_tags = {}
    for issue in issues:
        tags = issue.get('tagText', '')
        if tags:
            for tag in tags.split(','):
                tag = tag.strip()
                if tag:
                    all_tags[tag] = all_tags.get(tag, 0) + 1

    if all_tags:
        print("\nIssues by tag:")
        for tag, count in sorted(all_tags.items(), key=lambda x: x[1], reverse=True):
            print(f"  {tag}: {count}")

    print("\nIssue list:")
    print("─"*80)
    for idx, issue in enumerate(issues, 1):
        name = issue.get('name', 'Unnamed')
        tags = issue.get('tagText', 'No tags')

        print(f"  {idx}. {name}")
        print(f"      Tags: {tags}")


# =============================================================================
# MAIN DEMONSTRATION
# =============================================================================

def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("EXPORT ISSUES TO BCF")
    print("="*80)

    # Get all issues
    print("\n" + "─"*80)
    print("Getting Issues...")
    print("─"*80)

    all_issues = get_all_issues()

    if not all_issues:
        print("\n⚠ No issues found in the project")
        print("   Create issues first using issue management tools")
        return

    # Display summary
    display_issues_summary(all_issues)

    # Example 1: Export all issues
    print("\n" + "─"*80)
    print("EXAMPLE 1: Export All Issues")
    print("─"*80 + "\n")

    export_path = os.path.abspath("all_issues.bcf")
    print(f"💡 To export all issues:")
    print(f"   export_all_issues_to_bcf('{export_path}')")

    # Uncomment to actually export:
    success = export_all_issues_to_bcf(export_path)
    if success:
        create_bcf_export_summary(all_issues, export_path)

    # Example 2: Export by tag
    print("\n" + "─"*80)
    print("EXAMPLE 2: Export by Tag")
    print("─"*80 + "\n")

    print("💡 To export issues with 'Priority' tag:")
    print("   export_issues_by_tag('Priority', 'priority_issues.bcf')")

    # Uncomment to actually export:
    # export_issues_by_tag('Priority', 'priority_issues.bcf')

    # Example 3: Export recent issues
    print("\n" + "─"*80)
    print("EXAMPLE 3: Export Recent Issues")
    print("─"*80 + "\n")

    print("💡 To export issues from last 7 days:")
    print("   export_recent_issues(days=7, export_path='recent.bcf')")

    # Uncomment to actually export:
    # export_recent_issues(days=7, export_path='recent.bcf')

    # Example 4: Custom filter
    print("\n" + "─"*80)
    print("EXAMPLE 4: Export with Custom Filter")
    print("─"*80 + "\n")

    print("💡 To export with custom filter:")
    print("   def my_filter(issue):")
    print("       return 'structural' in issue.get('name', '').lower()")
    print("   export_issues_by_filter(my_filter, 'structural.bcf')")

    print("\n" + "="*80)
    print("✓ Demonstration complete!")
    print("="*80)
    print("\n💡 BCF Format:")
    print("   BCF is the standard format for BIM issue tracking")
    print("   Compatible with: Revit, Navisworks, BIM 360, and more")
    print("   Includes: Issue data, 3D viewpoints, snapshots, comments")
    print("\n💡 File locations:")
    print(f"   Current directory: {os.getcwd()}")
    print("="*80)


if __name__ == "__main__":
    main()
