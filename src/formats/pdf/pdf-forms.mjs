import { PDFDocument } from "pdf-lib";
import { writeFileSync } from "node:fs";
import { digest } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { verifyPdfReplay } from "./pdf-replay.mjs";
import { documentFonts } from "./pdf-fonts.mjs";
const value = (f) =>
  f.constructor.name === "PDFTextField"
    ? (f.getText() ?? "")
    : f.constructor.name === "PDFCheckBox"
      ? f.isChecked()
      : f.constructor.name === "PDFRadioGroup"
        ? (f.getSelected() ?? "")
        : (f.getSelected?.() ?? null);
export async function formFields(bytes) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false });
  return doc
    .getForm()
    .getFields()
    .map((f) => ({
      kind: "form-field",
      scope: "document",
      name: f.getName(),
      type: f.constructor.name,
      value: value(f),
      readOnly: f.isReadOnly(),
      options: f.getOptions?.() ?? null,
    }));
}
async function apply(bytes, a) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false }),
    form = doc.getForm();
  requireThat(
    !form.hasXFA(),
    "UNSUPPORTED_FEATURE",
    "XFA forms are not supported.",
  );
  requireThat(
    !form.getFields().some((f) => f.constructor.name === "PDFSignature"),
    "UNSUPPORTED_FEATURE",
    "Signature fields require a separate signing workflow.",
  );
  for (const edit of a.fields) {
    const f = form.getFieldMaybe(edit.name);
    requireThat(
      f && !f.isReadOnly(),
      "TARGET_CONFLICT",
      "Form field is missing or read-only.",
    );
    requireThat(
      JSON.stringify(value(f)) === JSON.stringify(edit.expectedValue),
      "TARGET_CONFLICT",
      "Form value differs from the expected value.",
    );
    const type = f.constructor.name;
    if (type === "PDFTextField") {
      requireThat(
        typeof edit.value === "string",
        "INVALID_ARGUMENT",
        "Text fields require strings.",
      );
      f.setText(edit.value);
    } else if (type === "PDFCheckBox") {
      requireThat(
        typeof edit.value === "boolean",
        "INVALID_ARGUMENT",
        "Checkboxes require booleans.",
      );
      edit.value ? f.check() : f.uncheck();
    } else if (
      ["PDFDropdown", "PDFOptionList", "PDFRadioGroup"].includes(type)
    ) {
      const requested = Array.isArray(edit.value) ? edit.value : [edit.value];
      requireThat(
        requested.every(
          (v) => typeof v === "string" && f.getOptions().includes(v),
        ),
        "INVALID_ARGUMENT",
        "Selection must match the field options.",
      );
      f.select(type === "PDFRadioGroup" ? requested[0] : requested);
    } else
      requireThat(false, "UNSUPPORTED_FEATURE", "Unsupported form field type.");
  }
  const fonts = await documentFonts(
    doc,
    a.fields.flatMap((e) =>
      Array.isArray(e.value)
        ? e.value
        : typeof e.value === "string"
          ? [e.value]
          : [],
    ),
  );
  form.updateFieldAppearances(fonts.regular);
  const out = Buffer.from(await doc.save({ useObjectStreams: false })),
    saved = await formFields(out);
  for (const e of a.fields)
    requireThat(
      JSON.stringify(saved.find((f) => f.name === e.name)?.value) ===
        JSON.stringify(e.value),
      "VALIDATION_FAILED",
      "Saved form field differs from requested value.",
    );
  return out;
}
export async function editForms(task, bytes) {
  if (task.kind === "pdf_fill_form") {
    const out = await apply(bytes, task.args);
    writeFileSync(task.output, out, { flag: "wx", mode: 0o600 });
    return {
      revision: digest(out),
      affectedTargets: task.args.fields.map((f) => f.name),
    };
  }
  await verifyPdfReplay(task, bytes, apply, "PDF form replay differs from the candidate.");
  return {
    passed: true,
    revision: digest(bytes),
    profile: "pdf.forms.v1",
    checks: ["form-edit-replay", "saved-field-values"],
    notVerified: ["visual-appearance", "non-Latin-font-support"],
  };
}
