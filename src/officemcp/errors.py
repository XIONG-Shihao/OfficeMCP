"""Refusals: every failure an agent can act on has a stable code, a message and a hint."""

# Each code an agent may receive, with what it means. The catalog and the agent
# guide list these; a code not here is a bug.
CODES = {
    "invalid_request": "The request does not match the tool's or action's schema.",
    "unknown_action": "No action of that name exists for the file's format.",
    "unsupported_format": "The file's format is not PDF, DOCX, XLSX or PPTX, or the conversion is not offered.",
    "path_not_allowed": "The path is outside the folders OfficeMCP may use.",
    "not_found": "The file does not exist.",
    "unreadable": "The file is not a valid document of its format, or is encrypted.",
    "revision_mismatch": "The file changed since the revision the request names.",
    "target_not_found": "The address does not name anything in the file.",
    "expected_mismatch": "The content at the address is not what the request expected.",
    "no_match": "The text to find does not occur.",
    "ambiguous_match": "The key or text matches more than once, or not the number of times expected.",
    "unsupported_content": "The target holds content this action cannot change safely (fields, objects, merged cells, ...).",
    "fidelity_loss": "Saving would lose content the file has; nothing was written. allowLoss saves anyway.",
    "engine_unavailable": "The request needs LibreOffice, which was not found.",
    "output_exists": "A file already exists at the new path; only office_edit without output.path replaces a file.",
    "too_large": "The file or the requested range is larger than the configured limits.",
}


class Refusal(Exception):
    """A request OfficeMCP will not carry out. Nothing is written."""

    def __init__(self, code: str, message: str, hint: str | None = None, **details):
        assert code in CODES, code
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint
        self.details = details

    def result(self) -> dict:
        out = {"status": "refused", "code": self.code, "message": self.message}
        if self.hint:
            out["hint"] = self.hint
        if self.details:
            out["details"] = self.details
        return out
