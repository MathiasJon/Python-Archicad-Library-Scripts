    success = export_all_issues_to_bcf(export_path)
    if success:
        create_bcf_export_summary(all_issues, export_path)