"""
Issue Management Script - 100% TAPIR Commands

This script manages Issues (BCF topics) in Archicad using TAPIR commands:
- CreateIssue (v1.0.2): Create new issues
- GetIssues (v1.0.2): List all issues
- DeleteIssue (v1.0.2): Delete issues
- AddCommentToIssue (v1.0.6): Add comments
- GetCommentsFromIssue (v1.0.6): Get comments
- AttachElementsToIssue (v1.0.6): Attach elements
- DetachElementsFromIssue (v1.0.6): Detach elements
- GetElementsAttachedToIssue (v1.0.6): Get attached elements
- ExportIssuesToBCF (v1.0.6): Export to BCF file
- ImportIssuesFromBCF (v1.0.6): Import from BCF file

Based on TAPIR API documentation.
"""

from archicad import ACConnection
from datetime import datetime

# Connect to Archicad
conn = ACConnection.connect()
assert conn, "Failed to connect to Archicad"

acc = conn.commands
act = conn.types
acu = conn.utilities


# =============================================================================
# CREATE & DELETE ISSUES
# =============================================================================

def create_issue(name, parent_issue_id=None, tag_text=""):
    """
    Create a new issue.
    Uses TAPIR CreateIssue command.

    Args:
        name: Issue name (required)
        parent_issue_id: Parent issue GUID (optional, for sub-issues)
        tag_text: Tag/label text (optional)

    Returns:
        Issue GUID or None
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'CreateIssue')

        parameters = {
            'name': name
        }

        if parent_issue_id:
            parameters['parentIssueId'] = {'guid': parent_issue_id}

        if tag_text:
            parameters['tagText'] = tag_text

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ CreateIssue failed: {error_msg}")
            return None

        # Parse response
        if isinstance(response, dict) and 'issueId' in response:
            issue_guid = response['issueId'].get('guid')
            if issue_guid:
                print(f"✓ Issue created: {name}")
                print(f"  GUID: {issue_guid}")
                return issue_guid

        return None

    except Exception as e:
        print(f"✗ Error creating issue: {e}")
        return None


def delete_issue(issue_guid, accept_all_elements=False):
    """
    Delete an issue.
    Uses TAPIR DeleteIssue command.

    Args:
        issue_guid: Issue GUID to delete
        accept_all_elements: Accept all creation/deletion/modification (default: False)

    Returns:
        Boolean indicating success
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'DeleteIssue')

        parameters = {
            'issueId': {'guid': issue_guid},
            'acceptAllElements': accept_all_elements
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'success' in response:
            if response['success']:
                print(f"✓ Issue deleted: {issue_guid}")
                return True
            else:
                error = response.get('error', {})
                print(
                    f"✗ DeleteIssue failed: {error.get('message', 'Unknown error')}")
                return False

        return False

    except Exception as e:
        print(f"✗ Error deleting issue: {e}")
        return False


# =============================================================================
# GET ISSUES
# =============================================================================

def get_all_issues():
    """
    Get all issues in the project.
    Uses TAPIR GetIssues command.

    Returns:
        List of issue dictionaries
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'GetIssues')

        # No parameters needed
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


def list_all_issues():
    """
    List all issues with formatted output.
    """
    issues = get_all_issues()

    if not issues:
        print("⚠ No issues found in project")
        return

    print(f"\nTotal issues: {len(issues)}")
    print("="*80)

    for idx, issue in enumerate(issues, 1):
        issue_id = issue.get('issueId', {}).get('guid', 'N/A')
        name = issue.get('name', 'Unnamed')
        tag_text = issue.get('tagText', '')
        crea_time = issue.get('creaTime', 0)
        modi_time = issue.get('modiTime', 0)
        parent_id = issue.get('parentIssueId', {}).get('guid', None)

        print(f"\n{idx}. {name}")
        print(f"   GUID: {issue_id}")
        if tag_text:
            print(f"   Tags: {tag_text}")
        if parent_id:
            print(f"   Parent: {parent_id}")
        print(
            f"   Created: {datetime.fromtimestamp(crea_time) if crea_time > 0 else 'N/A'}")
        print(
            f"   Modified: {datetime.fromtimestamp(modi_time) if modi_time > 0 else 'N/A'}")


def find_issue_by_name(issue_name):
    """
    Find an issue by its name.

    Args:
        issue_name: Name to search for

    Returns:
        Issue dictionary or None
    """
    issues = get_all_issues()

    for issue in issues:
        if issue.get('name', '').lower() == issue_name.lower():
            return issue

    return None


# =============================================================================
# COMMENTS
# =============================================================================

def add_comment_to_issue(issue_guid, text, author="", status="Unknown"):
    """
    Add a comment to an issue.
    Uses TAPIR AddCommentToIssue command.

    Args:
        issue_guid: Issue GUID
        text: Comment text (required)
        author: Comment author (optional)
        status: Comment status: "Unknown", "Error", "Info", "Warning" (optional)

    Returns:
        Boolean indicating success
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'AddCommentToIssue')

        parameters = {
            'issueId': {'guid': issue_guid},
            'text': text
        }

        if author:
            parameters['author'] = author

        if status:
            parameters['status'] = status

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'success' in response:
            if response['success']:
                print(f"✓ Comment added to issue")
                return True
            else:
                error = response.get('error', {})
                print(
                    f"✗ AddCommentToIssue failed: {error.get('message', 'Unknown error')}")
                return False

        return False

    except Exception as e:
        print(f"✗ Error adding comment: {e}")
        return False


def get_comments_from_issue(issue_guid):
    """
    Get all comments from an issue.
    Uses TAPIR GetCommentsFromIssue command.

    Args:
        issue_guid: Issue GUID

    Returns:
        List of comment dictionaries
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'GetCommentsFromIssue')

        parameters = {
            'issueId': {'guid': issue_guid}
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ GetCommentsFromIssue failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'comments' in response:
            return response['comments']

        return []

    except Exception as e:
        print(f"✗ Error getting comments: {e}")
        return []


def display_comments(issue_guid):
    """
    Display all comments for an issue.

    Args:
        issue_guid: Issue GUID
    """
    comments = get_comments_from_issue(issue_guid)

    if not comments:
        print("  No comments")
        return

    print(f"  {len(comments)} comment(s):")
    for idx, comment in enumerate(comments, 1):
        author = comment.get('author', 'Unknown')
        text = comment.get('text', '')
        status = comment.get('status', 'Unknown')
        crea_time = comment.get('creaTime', 0)

        print(f"\n    [{idx}] {author} ({status})")
        print(f"        {text}")
        print(
            f"        {datetime.fromtimestamp(crea_time) if crea_time > 0 else 'N/A'}")


# =============================================================================
# ATTACH/DETACH ELEMENTS
# =============================================================================

def attach_elements_to_issue(issue_guid, element_guids, attachment_type="Highlight"):
    """
    Attach elements to an issue.
    Uses TAPIR AttachElementsToIssue command.

    Args:
        issue_guid: Issue GUID
        element_guids: List of element GUIDs
        attachment_type: "Highlight", "Hidden", "Selection", "Problem" (default: "Highlight")

    Returns:
        Boolean indicating success
    """
    try:
        command_id = act.AddOnCommandId(
            'TapirCommand', 'AttachElementsToIssue')

        # Prepare elements list
        elements = [
            {'elementId': {'guid': guid}} for guid in element_guids
        ]

        parameters = {
            'issueId': {'guid': issue_guid},
            'elements': elements,
            'type': attachment_type
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'success' in response:
            if response['success']:
                print(
                    f"✓ Attached {len(element_guids)} element(s) as '{attachment_type}'")
                return True
            else:
                error = response.get('error', {})
                print(
                    f"✗ AttachElementsToIssue failed: {error.get('message', 'Unknown error')}")
                return False

        return False

    except Exception as e:
        print(f"✗ Error attaching elements: {e}")
        return False


def detach_elements_from_issue(issue_guid, element_guids):
    """
    Detach elements from an issue.
    Uses TAPIR DetachElementsFromIssue command.

    Args:
        issue_guid: Issue GUID
        element_guids: List of element GUIDs to detach

    Returns:
        Boolean indicating success
    """
    try:
        command_id = act.AddOnCommandId(
            'TapirCommand', 'DetachElementsFromIssue')

        # Prepare elements list
        elements = [
            {'elementId': {'guid': guid}} for guid in element_guids
        ]

        parameters = {
            'issueId': {'guid': issue_guid},
            'elements': elements
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'success' in response:
            if response['success']:
                print(f"✓ Detached {len(element_guids)} element(s)")
                return True
            else:
                error = response.get('error', {})
                print(
                    f"✗ DetachElementsFromIssue failed: {error.get('message', 'Unknown error')}")
                return False

        return False

    except Exception as e:
        print(f"✗ Error detaching elements: {e}")
        return False


def get_elements_attached_to_issue(issue_guid, attachment_type="Highlight"):
    """
    Get elements attached to an issue.
    Uses TAPIR GetElementsAttachedToIssue command.

    Args:
        issue_guid: Issue GUID
        attachment_type: "Highlight", "Hidden", "Selection", "Problem"

    Returns:
        List of element GUIDs
    """
    try:
        command_id = act.AddOnCommandId(
            'TapirCommand', 'GetElementsAttachedToIssue')

        parameters = {
            'issueId': {'guid': issue_guid},
            'type': attachment_type
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'succeeded' in response and not response['succeeded']:
            error_msg = response.get('error', {}).get(
                'message', 'Unknown error')
            print(f"✗ GetElementsAttachedToIssue failed: {error_msg}")
            return []

        # Parse response
        if isinstance(response, dict) and 'elements' in response:
            elements = response['elements']
            guids = []
            for elem in elements:
                if isinstance(elem, dict) and 'elementId' in elem:
                    guid = elem['elementId'].get('guid')
                    if guid:
                        guids.append(guid)
            return guids

        return []

    except Exception as e:
        print(f"✗ Error getting attached elements: {e}")
        return []


# =============================================================================
# BCF IMPORT/EXPORT
# =============================================================================

def export_issues_to_bcf(issue_guids, export_path, use_external_id=True, align_by_survey_point=False):
    """
    Export issues to a BCF file.
    Uses TAPIR ExportIssuesToBCF command.

    Args:
        issue_guids: List of issue GUIDs to export
        export_path: Full path to BCF file (including filename)
        use_external_id: Use external IFC ID (default: True)
        align_by_survey_point: Align by survey point vs project origin (default: False)

    Returns:
        Boolean indicating success
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'ExportIssuesToBCF')

        # Prepare issues list
        issues = [
            {'issueId': {'guid': guid}} for guid in issue_guids
        ]

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
                print(f"✓ Exported {len(issue_guids)} issue(s) to BCF")
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
        return False


def import_issues_from_bcf(import_path, align_by_survey_point=False):
    """
    Import issues from a BCF file.
    Uses TAPIR ImportIssuesFromBCF command.

    Args:
        import_path: Full path to BCF file (including filename)
        align_by_survey_point: Align by survey point vs project origin (default: False)

    Returns:
        Boolean indicating success
    """
    try:
        command_id = act.AddOnCommandId('TapirCommand', 'ImportIssuesFromBCF')

        parameters = {
            'importPath': import_path,
            'alignBySurveyPoint': align_by_survey_point
        }

        response = acc.ExecuteAddOnCommand(command_id, parameters)

        # Check for error
        if isinstance(response, dict) and 'success' in response:
            if response['success']:
                print(f"✓ Imported issues from BCF")
                print(f"  File: {import_path}")
                return True
            else:
                error = response.get('error', {})
                print(
                    f"✗ ImportIssuesFromBCF failed: {error.get('message', 'Unknown error')}")
                return False

        return False

    except Exception as e:
        print(f"✗ Error importing from BCF: {e}")
        return False


# =============================================================================
# MAIN DEMONSTRATION
# =============================================================================

def main():
    """Main function to demonstrate the script."""
    print("\n" + "="*80)
    print("ISSUE MANAGEMENT - 100% TAPIR")
    print("="*80)

    # Example 1: List existing issues
    print("\n" + "─"*80)
    print("EXAMPLE 1: List All Issues")
    print("─"*80)

    list_all_issues()

    # Example 2: Create a new issue
    print("\n" + "─"*80)
    print("EXAMPLE 2: Create New Issue")
    print("─"*80 + "\n")

    issue_guid = create_issue(
        name="Wall alignment issue in Room 301",
        tag_text="High Priority, Structural"
    )

    # Example 3: Add comment to issue
    if issue_guid:
        print("\n" + "─"*80)
        print("EXAMPLE 3: Add Comment")
        print("─"*80 + "\n")

        add_comment_to_issue(
            issue_guid,
            "Wall appears misaligned with structural grid. Needs verification.",
            author="John Architect",
            status="Warning"
        )

        # Display comments
        print("\nComments for this issue:")
        display_comments(issue_guid)

    # Example 4: Attach selected elements
    if issue_guid:
        print("\n" + "─"*80)
        print("EXAMPLE 4: Attach Selected Elements")
        print("─"*80 + "\n")

        selected = acc.GetSelectedElements()

        if selected:
            element_guids = [elem.elementId.guid for elem in selected]
            attach_elements_to_issue(issue_guid, element_guids, "Problem")

            # Show attached elements
            attached = get_elements_attached_to_issue(issue_guid, "Problem")
            print(f"\n  Attached elements: {len(attached)}")
        else:
            print("  ⚠ No elements selected")

    # Example 5: Create sub-issue
    if issue_guid:
        print("\n" + "─"*80)
        print("EXAMPLE 5: Create Sub-Issue")
        print("─"*80 + "\n")

        sub_issue_guid = create_issue(
            name="Verify grid alignment",
            parent_issue_id=issue_guid,
            tag_text="Sub-task"
        )

    # Example 6: Export to BCF (commented out - requires file path)
    print("\n" + "─"*80)
    print("EXAMPLE 6: Export to BCF")
    print("─"*80 + "\n")

    print("💡 To export issues to BCF:")
    print("   issues = get_all_issues()")
    print("   issue_guids = [i['issueId']['guid'] for i in issues]")
    print("   export_issues_to_bcf(issue_guids, 'C:/path/to/issues.bcf')")

    # Example 7: Final list
    print("\n" + "─"*80)
    print("EXAMPLE 7: Updated Issue List")
    print("─"*80)

    list_all_issues()

    print("\n" + "="*80)
    print("✓ All operations completed using TAPIR commands only!")
    print("="*80)
    print("\n💡 TAPIR Commands used:")
    print("   • CreateIssue - Create new issues")
    print("   • GetIssues - List all issues")
    print("   • DeleteIssue - Delete issues")
    print("   • AddCommentToIssue - Add comments")
    print("   • GetCommentsFromIssue - Get comments")
    print("   • AttachElementsToIssue - Attach elements")
    print("   • GetElementsAttachedToIssue - Get attached elements")
    print("   • ExportIssuesToBCF - Export to BCF")
    print("   • ImportIssuesFromBCF - Import from BCF")
    print("="*80)


if __name__ == "__main__":
    main()
