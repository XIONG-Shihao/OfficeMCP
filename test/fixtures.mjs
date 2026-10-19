import { zipSync, strToU8 } from "fflate";
import { writeFileSync } from "node:fs";
import { join } from "node:path";
export const originalText = "Payment within 30 days";
export function docxBytes(
  paragraphs = [originalText, "Unrelated paragraph must remain unchanged."],
) {
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    paragraphs
      .map((t) => "<w:p><w:r><w:t>" + t + "</w:t></w:r></w:p>")
      .join("") +
    "<w:sectPr/></w:body></w:document>";
  return Buffer.from(
    zipSync(
      {
        "[Content_Types].xml": strToU8(
          '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        ),
        "_rels/.rels": strToU8(
          '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        ),
        "word/document.xml": strToU8(xml),
        "docProps/custom.xml": strToU8(
          '<fixture xmlns="urn:officemcp:fixture">UNRELATED-PART-MUST-SURVIVE</fixture>',
        ),
      },
      { level: 6 },
    ),
  );
}
export function pdfBytes() {
  const stream =
    "BT /F1 18 Tf 72 700 Td (OfficeMCP synthetic PDF. Invoice total: 125 dollars.) Tj ET";
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
  ];
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (let i = 0; i < objects.length; i++) {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${i + 1} 0 obj\n${objects[i]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(pdf);
  pdf +=
    `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` +
    offsets
      .slice(1)
      .map((n) => String(n).padStart(10, "0") + " 00000 n \n")
      .join("") +
    `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(pdf);
}
export function writeFixtures(root) {
  writeFileSync(join(root, "contract.docx"), docxBytes());
  writeFileSync(join(root, "invoice.pdf"), pdfBytes());
}
