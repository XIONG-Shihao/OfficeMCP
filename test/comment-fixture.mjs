/** Existing OOXML comments are fixture input, never written through OfficeMCP. */
import { Document, Packer, Paragraph, TextRun, CommentRangeStart, CommentRangeEnd, CommentReference } from 'docx';
export async function commentedDocument({ anchor = 'Reviewer note anchor.', comment = 'Keep this actual reviewer note.', before = [], after = [] } = {}) {
  return Packer.toBuffer(new Document({ comments: { children: [{ id: 0, author: 'Reviewer', initials: 'R', children: [new Paragraph(comment)] }] },
    sections: [{ children: [...before, new Paragraph({ children: [new CommentRangeStart(0), new TextRun(anchor), new CommentRangeEnd(0), new TextRun({ children: [new CommentReference(0)] })] }), ...after] }] }));
}
