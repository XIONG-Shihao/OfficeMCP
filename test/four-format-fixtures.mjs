import { zipSync, strToU8 } from "fflate";
import { writeFileSync } from "node:fs";
import { join } from "node:path";
const X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rels = (rows) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rows.map(([Id, Type, Target, external]) => `<Relationship Id="${Id}" Type="${R}/${Type}" Target="${Target}"${external ? ' TargetMode="External"' : ""}/>`).join("")}</Relationships>`;
const pack = (parts) =>
  Buffer.from(
    zipSync(
      Object.fromEntries(
        Object.entries(parts).map(([n, v]) => [
          n,
          typeof v === "string" ? strToU8(v) : v,
        ]),
      ),
      { level: 6 },
    ),
  );
const types = (overrides) =>
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/>${overrides.map(([n, t]) => `<Override PartName="/${n}" ContentType="application/vnd.openxmlformats-officedocument.${t}+xml"/>`).join("")}</Types>`;
export function xlsxBytes() {
  return pack({
    "[Content_Types].xml": types([
      ["xl/workbook.xml", "spreadsheetml.sheet.main"],
      ["xl/worksheets/budget.xml", "spreadsheetml.worksheet"],
      ["xl/worksheets/hidden.xml", "spreadsheetml.worksheet"],
      ["xl/sharedStrings.xml", "spreadsheetml.sharedStrings"],
      ["xl/styles.xml", "spreadsheetml.styles"],
    ]),
    "_rels/.rels": rels([["root", "officeDocument", "xl/workbook.xml"]]),
    "xl/workbook.xml": `<workbook xmlns="${X}" xmlns:r="${R}"><workbookPr date1904="1"/><sheets><sheet name="Budget" sheetId="7" r:id="budget"/><sheet name="Internal" sheetId="11" state="veryHidden" r:id="hidden"/></sheets><definedNames><definedName name="InvoiceTotal">Budget!$B$2</definedName></definedNames><calcPr calcId="123" fullCalcOnLoad="1"/></workbook>`,
    "xl/_rels/workbook.xml.rels": rels([
      ["budget", "worksheet", "worksheets/budget.xml"],
      ["hidden", "worksheet", "worksheets/hidden.xml"],
      ["strings", "sharedStrings", "sharedStrings.xml"],
      ["styles", "styles", "styles.xml"],
    ]),
    "xl/worksheets/budget.xml": `<worksheet xmlns="${X}"><dimension ref="A1:E4"/><cols><col min="5" max="5" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>125</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><v>#DIV/0!</v></c><c r="E1"><v>9007199254740993</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><r><t>Invoice </t></r><r><t>total</t></r></is></c><c r="B2" s="1"><f>SUM(B1,25)</f><v>150</v></c><c r="C2"><f t="shared" si="0" ref="C2:C3">B2*2</f><v>300</v></c><c r="D2"><f>NOW()</f></c></row><row r="3" hidden="1"><c r="C3"><f t="shared" si="0"/><v>350</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A4:B4"/></mergeCells></worksheet>`,
    "xl/worksheets/hidden.xml": `<worksheet xmlns="${X}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Hidden reserve 77</t></is></c></row></sheetData></worksheet>`,
    "xl/sharedStrings.xml": `<sst xmlns="${X}" count="1" uniqueCount="1"><si><r><t>Unit </t></r><r><t>cost</t></r><rPh sb="0" eb="1"><t>phonetic-not-visible-text</t></rPh></si></sst>`,
    "xl/styles.xml": `<styleSheet xmlns="${X}"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>`,
    "xl/worksheets/_rels/budget.xml.rels": rels([
      ["external", "hyperlink", "https://example.invalid/must-not-fetch", true],
    ]),
  });
}
const textShape = (id, text) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000" cy="1000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const tree = (body) =>
  `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${body}</p:spTree></p:cSld>`;
export const imageBytes = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aWQAAAABJRU5ErkJggg==",
  "base64",
);
export function pptxBytes() {
  return pack({
    "[Content_Types].xml": types([
      ["ppt/presentation.xml", "presentationml.presentation.main"],
      ["ppt/slides/slide9.xml", "presentationml.slide"],
      ["ppt/slides/slide2.xml", "presentationml.slide"],
      ["ppt/notesSlides/notesSlide1.xml", "presentationml.notesSlide"],
    ]),
    "_rels/.rels": rels([["root", "officeDocument", "ppt/presentation.xml"]]),
    "ppt/presentation.xml": `<p:presentation xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="401" r:id="first"/><p:sldId id="402" r:id="second"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": rels([
      ["first", "slide", "slides/slide9.xml"],
      ["second", "slide", "slides/slide2.xml"],
    ]),
    "ppt/slides/slide9.xml": `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}" show="0">${tree(textShape(2, "Quarterly review") + `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000" cy="4000"/><a:chOff x="0" y="0"/><a:chExt cx="4000" cy="4000"/></a:xfrm></p:grpSpPr>${textShape(4, "Revenue 150")}</p:grpSp><p:pic><p:nvPicPr><p:cNvPr id="5" name="Logo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="logo"/></p:blipFill><p:spPr/></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Margin 42</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`)}</p:sld>`,
    "ppt/slides/slide2.xml": `<p:sld xmlns:p="${P}" xmlns:a="${A}">${tree(textShape(2, "Next steps"))}</p:sld>`,
    "ppt/slides/_rels/slide9.xml.rels": rels([
      ["notes", "notesSlide", "../notesSlides/notesSlide1.xml"],
      ["logo", "image", "../media/logo.png"],
      ["external", "hyperlink", "https://example.invalid/no-fetch", true],
    ]),
    "ppt/notesSlides/notesSlide1.xml": `<p:notes xmlns:p="${P}" xmlns:a="${A}">${tree(textShape(2, "Speaker note: reserve 77 is confidential fixture text."))}</p:notes>`,
    "ppt/media/logo.png": imageBytes,
  });
}
export function writeFourFormatFixtures(root) {
  writeFileSync(join(root, "budget.xlsx"), xlsxBytes());
  writeFileSync(join(root, "review.pptx"), pptxBytes());
}
