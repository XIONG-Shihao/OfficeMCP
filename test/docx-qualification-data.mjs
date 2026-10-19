export const QUALIFICATION_GATES = Object.freeze([
  "valid-target-success",
  "stale-target-refusal",
  "non-mutating-preview",
  "multi-step-rollback",
  "save-reopen-semantic-equality",
  "unrelated-content-preservation",
  "representative-render",
]);

export const QUALIFIED_ACTIONS = Object.freeze([
  "insert_paragraphs", "insert_heading", "replace_text", "delete_text",
  "delete_blocks", "append_list", "create_table", "rewrite_block",
  "normalize_body_font_size", "set_font_family", "apply_letter_spacing",
  "insert_toc", "insert_table_row", "insert_table_column", "delete_table_row",
  "delete_table_column", "convert_list", "format_text", "apply_style",
  "delete_table", "add_hyperlink", "write_table_cells", "merge_table_cells",
  "unmerge_table_cells", "insert_image", "replace_image", "delete_image",
  "resize_image", "set_image_alt_text", "rotate_image",
  "accept_tracked_change", "reject_tracked_change", "set_content_control_text",
  "set_checkbox_control", "set_date_control", "set_choice_control",
  "edit_repeating_section", "update_toc", "format_table_cell",
  "format_table_row", "format_paragraph_advanced", "materialize_header_footer",
  "replace_header_footer_text", "patch_hyperlink", "update_footnote",
  "insert_caption", "update_caption", "remove_caption",
  "insert_cross_reference", "update_endnote",
]);

export const LIMITED_ACTIONS = Object.freeze([
  "accept_tracked_changes",
  "reject_tracked_changes",
]);

export function evidenceFor(action) {
  return {
    "valid-target-success": `docx-agent:action-fixture:${action}`,
    "stale-target-refusal": "docx-agent:revision-bound-targets",
    "non-mutating-preview": "docx-agent:previewable-plans",
    "multi-step-rollback": "docx-contract:batch-failure-rollback",
    "save-reopen-semantic-equality": "docx-agent:facade-save-validation",
    "unrelated-content-preservation": "docx-agent:package-validation",
    "representative-render": "docx-agent:libreoffice-page-evidence",
  };
}
