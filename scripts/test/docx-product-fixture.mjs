import { Document, Packer, Paragraph, TextRun, Footer, Table, TableRow, TableCell, ImageRun, CommentRangeStart, CommentRangeEnd, CommentReference } from 'docx';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

/** Synthetic only. Repeated text, mixed formatting, stories, table and a real SDT. */
export async function productFixture({ references = 12 } = {}) {
  const bytes = await Packer.toBuffer(new Document({ comments: { children: [{ id: 0, author: 'Reviewer', children: [new Paragraph('Approved by legal')] }] }, sections: [{
    footers: { default: new Footer({ children: [new Paragraph('Liability is capped at USD 500.')] }) },
    children: [
      new Paragraph({ children: [new TextRun('Payment within '), new TextRun({ text: '30', bold: true }), new TextRun(' days.')] }),
      new Paragraph({ children: [new TextRun({ text: 'Deliverables are due on 30 September.', italics: true })] }),
      new Paragraph({ children: [new CommentRangeStart(0), new TextRun('Review clause remains effective.'), new CommentRangeEnd(0), new TextRun({ children: [new CommentReference(0)] })] }),
      new Paragraph('Duplicate risk'), new Paragraph('Duplicate risk'),
      new Paragraph('CLIENT_CONTROL_MARKER'),
      new Paragraph({ children: [new ImageRun({
        data: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=', 'base64'),
        type: 'png', transformation: { width: 16, height: 16 },
        altText: { title: 'Company mark', description: 'Company mark', name: 'Company mark' },
      })] }),
      new Table({ rows: [['Service', 'Price USD'], ['Support', '125'], ['Training', '300']].map(row =>
        new TableRow({ children: row.map(text => new TableCell({ children: [new Paragraph(text)] })) })) }),
      ...Array.from({ length: references }, (_, i) => new Paragraph(`Reference paragraph ${i}.`)),
    ],
  }] }));
  const parts = unzipSync(bytes), xml = strFromU8(parts['word/document.xml']);
  const marker = /<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*?CLIENT_CONTROL_MARKER(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/;
  if (!marker.test(xml)) throw new Error('Fixture marker missing');
  parts['word/document.xml'] = strToU8(xml.replace(marker,
    '<w:p><w:sdt><w:sdtPr><w:id w:val="104"/><w:tag w:val="client"/><w:alias w:val="Client name"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Original client</w:t></w:r></w:sdtContent></w:sdt></w:p>'));
  return Buffer.from(zipSync(parts));
}

export const PRODUCT_JOBS = Object.freeze([
  { id: 'formatted-clause', instruction: 'Change the payment deadline from 30 days to 45 days. Preserve the bold number and every other part of the document. Read and report the liability qualification in the footer.' },
  { id: 'table-by-key', instruction: 'In the table with Service and Price USD headers, change the Support price from 125 to 150. Use the row business key and column heading to identify the cell. Leave Training and all other content unchanged.' },
  { id: 'template-control', instruction: 'Set the text content control tagged client to Meridian Labs. Preserve the control itself and all other content.' },
  { id: 'comment-read', instruction: 'Read the existing comment and report its text and open/resolved status. Save an unchanged copy. Do not change, resolve or reply to any comment.' },
  { id: 'save-recovery', instruction: 'Save an unchanged copy of the document. If the save response is unavailable, recover its completed receipt by operation ID; do not produce a second output.' },
  { id: 'ambiguous-refusal', instruction: 'Try to change exactly one Duplicate risk paragraph to Reviewed risk without choosing between the two matching paragraphs. If ambiguous, report the refusal and do not edit or save.' },
]);
