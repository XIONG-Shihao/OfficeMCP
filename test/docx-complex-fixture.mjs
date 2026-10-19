import { Document, Packer, Paragraph, Math as WordMath, MathRun } from 'docx';
import { unzipSync, zipSync, strFromU8 } from 'fflate';
import ExcelJS from 'exceljs';
export async function complexFixture({ formula = false, external = false, literal = false } = {}) {
  const parts = unzipSync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Outside stays.'), new Paragraph({ children: [new WordMath({ children: [new MathRun('x+y')] })] })] }] })));
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', C = 'http://schemas.openxmlformats.org/drawingml/2006/chart', A = 'http://schemas.openxmlformats.org/drawingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const textbox = `<w:p><w:r><w:pict><v:shapetype xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="Box1" type="#_x0000_t202" style="width:150pt;height:50pt"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Payment 30 days.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
  const drawing = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="4000000" cy="3000000"/><wp:docPr id="50" name="Sales chart"/><a:graphic xmlns:a="${A}"><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="officeChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  parts['word/document.xml'] = Buffer.from(strFromU8(parts['word/document.xml']).replace('<w:sectPr', textbox + drawing + '<w:sectPr'));
  const points = (values, numeric) => `<c:ptCount val="${values.length}"/>` + values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  const cat = literal ? `<c:strLit>${points(['Jan', 'Feb'], false)}</c:strLit>` : `<c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache>${points(['Jan', 'Feb'], false)}</c:strCache></c:strRef>`;
  const val = literal ? `<c:numLit>${points([10, 20], true)}</c:numLit>` : `<c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode>${points([10, 20], true)}</c:numCache></c:numRef>`;
  parts['word/charts/chart1.xml'] = Buffer.from(`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>Original Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat>${cat}</c:cat><c:val>${val}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/><c:crosses val="autoZero"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart>${literal ? '' : '<c:externalData r:id="workbook"><c:autoUpdate val="0"/></c:externalData>'}</c:chartSpace>`);
  const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  parts['word/_rels/document.xml.rels'] = Buffer.from(strFromU8(parts['word/_rels/document.xml.rels']).replace('</Relationships>', `<Relationship Id="officeChart1" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`));
  parts['[Content_Types].xml'] = Buffer.from(strFromU8(parts['[Content_Types].xml']).replace('</Types>', `<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>${literal ? '' : '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'}</Types>`));
  if (!literal) {
    const book = new ExcelJS.Workbook(), sheet = book.addWorksheet('Sheet1'); sheet.addRows([['Month', 'Revenue', 'Unrelated'], ['Jan', 10, 'Keep A'], ['Feb', 20, 'Keep B']]); sheet.getCell('B2').numFmt = '$0.00';
    if (formula) sheet.getCell('B2').value = { formula: '5+5', result: 10 };
    parts['word/embeddings/data.xlsx'] = Buffer.from(await book.xlsx.writeBuffer());
    parts['word/charts/_rels/chart1.xml.rels'] = Buffer.from(`<Relationships xmlns="${rel}"><Relationship Id="workbook" Type="${R}/package" Target="${external ? 'https://example.invalid/data.xlsx' : '../embeddings/data.xlsx'}"${external ? ' TargetMode="External"' : ''}/></Relationships>`);
  }
  parts['word/media/keep.bin'] = new Uint8Array([4, 9, 12]);
  parts['[Content_Types].xml'] = Buffer.from(strFromU8(parts['[Content_Types].xml']).replace('</Types>', '<Default Extension="bin" ContentType="application/octet-stream"/></Types>'));
  return Buffer.from(zipSync(parts));
}
