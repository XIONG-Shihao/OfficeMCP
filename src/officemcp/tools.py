"""The seven MCP tools: names, descriptions and input schemas. Their behaviour is in service.py."""

FORMAT = {"type": "string", "enum": ["docx", "xlsx", "pptx", "pdf"]}
PATH = {"type": "string", "minLength": 1, "maxLength": 4096,
        "description": "The file: absolute, or relative to the first folder OfficeMCP may use (the project)."}
NEW_PATH = {"type": "string", "minLength": 1, "maxLength": 4096,
            "description": "A new file's path, in an existing folder OfficeMCP may use; an existing file is never overwritten."}
ACTION_LIST = {
    "type": "array", "minItems": 1, "maxItems": 500,
    "items": {"type": "object", "required": ["action"], "properties": {"action": {"type": "string"}},
              "description": "One action: {\"action\": name, ...the fields office_actions lists for it}."},
    "description": "Actions applied in order, all or none.",
}
OUTPUT = {"type": "object", "properties": {"path": NEW_PATH}, "required": ["path"], "additionalProperties": False}

TOOLS = [
    {
        "name": "office_actions",
        "description": "List the actions office_edit and office_create accept, with each action's fields. Filter by format (docx, xlsx, pptx, pdf) and family.",
        "inputSchema": {"type": "object", "properties": {"format": FORMAT, "family": {"type": "string", "maxLength": 100}}, "additionalProperties": False},
        "readOnly": True,
    },
    {
        "name": "office_inspect",
        "description": "Outline a file: its structure with addresses, its revision (needed by office_edit), and anything its library cannot keep when saving.",
        "inputSchema": {"type": "object", "properties": {"path": PATH}, "required": ["path"], "additionalProperties": False},
        "readOnly": True,
    },
    {
        "name": "office_read",
        "description": "Read the content at an address (body/p[3], Sheet!A1:D20, slide[2]/shape[5], page[4], ...), or page through the whole file without one.",
        "inputSchema": {"type": "object", "properties": {
            "path": PATH, "address": {"type": "string", "maxLength": 500},
            "offset": {"type": "integer", "minimum": 0, "default": 0}, "limit": {"type": "integer", "minimum": 1, "maximum": 1000, "default": 100}},
            "required": ["path"], "additionalProperties": False},
        "readOnly": True,
    },
    {
        "name": "office_find",
        "description": "Find text in a file; returns the addresses where it occurs.",
        "inputSchema": {"type": "object", "properties": {
            "path": PATH, "text": {"type": "string", "minLength": 1, "maxLength": 10000},
            "matchCase": {"type": "boolean", "default": False}, "limit": {"type": "integer", "minimum": 1, "maximum": 1000, "default": 100}},
            "required": ["path", "text"], "additionalProperties": False},
        "readOnly": True,
    },
    {
        "name": "office_edit",
        "description": "Apply actions to a file, all or none, and save it in place (or as a new file with output.path). Refuses if the file changed since expectedRevision, and if saving would lose content unless allowLoss is true. dryRun checks everything without writing.",
        "inputSchema": {"type": "object", "properties": {
            "path": PATH, "expectedRevision": {"type": "string", "pattern": "^sha256:[a-f0-9]{64}$", "description": "The revision office_inspect or office_read returned."},
            "actions": ACTION_LIST, "output": {**OUTPUT, "description": "Save as this new file instead of in place."},
            "dryRun": {"type": "boolean", "default": False},
            "allowLoss": {"type": "boolean", "default": False, "description": "Save even if the fidelity check finds content the save would lose; the result lists it."},
            "recalculate": {"type": "boolean", "default": False, "description": "Workbooks: recalculate formulas with LibreOffice before writing."}},
            "required": ["path", "expectedRevision", "actions"], "additionalProperties": False},
        "readOnly": False,
    },
    {
        "name": "office_create",
        "description": "Create a new file at a path, applying actions to it. A new PDF is built from compose actions (add_heading, add_paragraph, add_table, add_page_break).",
        "inputSchema": {"type": "object", "properties": {
            "format": FORMAT, "path": {**NEW_PATH, "description": NEW_PATH["description"] + " It ends in the format's extension."},
            "actions": ACTION_LIST},
            "required": ["format", "path"], "additionalProperties": False},
        "readOnly": False,
    },
    {
        "name": "office_convert",
        "description": "Convert a file to a new file, next to it unless output.path says where: to txt from any format, to csv from a workbook (one sheet), to pdf from docx, xlsx or pptx (needs LibreOffice).",
        "inputSchema": {"type": "object", "properties": {
            "path": PATH, "to": {"type": "string", "enum": ["txt", "csv", "pdf"]}, "output": {**OUTPUT, "description": "Where to write the result."},
            "sheet": {"type": "string", "maxLength": 31, "description": "csv: the sheet; the first otherwise."}},
            "required": ["path", "to"], "additionalProperties": False},
        "readOnly": False,
    },
]

RESULT = {
    "type": "object",
    "properties": {
        "status": {"type": "string", "enum": ["ok", "refused", "error"]},
        "code": {"type": "string", "description": "When refused: a stable reason code."},
        "message": {"type": "string"},
        "hint": {"type": "string", "description": "When refused: what to do next."},
    },
    "required": ["status"],
}
