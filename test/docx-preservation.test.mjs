import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { zipSync, strToU8 } from "fflate";
import { packageEvidence } from "../src/formats/opc/engine-evidence.mjs";

const contentTypes = strToU8(
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
);
const rootRelationships = strToU8(
  '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
);
const document = strToU8(
  '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
);
const relationships = (
  first = "media/one.png",
  includeSecond = true,
  second = "media/two.png",
) =>
  strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${first}"/>${includeSecond ? `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${second}"/>` : ""}</Relationships>`,
  );

const sha256 = (bytes) =>
  `sha256:${createHash("sha256").update(bytes).digest("hex")}`;

function packageBytes({
  replaceFirst = false,
  dropSecond = false,
  orphan = false,
  retargetSecond = false,
} = {}) {
  return Buffer.from(
    zipSync({
      "[Content_Types].xml": contentTypes,
      "_rels/.rels": rootRelationships,
      "word/document.xml": document,
      "word/_rels/document.xml.rels": relationships(
        replaceFirst ? "media/three.png" : "media/one.png",
        !dropSecond,
        retargetSecond ? "media/one.png" : "media/two.png",
      ),
      ...(!replaceFirst ? { "word/media/one.png": Uint8Array.from([1, 2, 3]) } : {}),
      ...(replaceFirst ? { "word/media/three.png": Uint8Array.from([7, 8, 9]) } : {}),
      ...(!dropSecond ? { "word/media/two.png": Uint8Array.from([4, 5, 6]) } : {}),
      ...(orphan ? { "word/media/orphan.png": Uint8Array.from([10]) } : {}),
    }),
  );
}

test("DOCX media validation permits only the targeted binary delta", () => {
  const before = packageBytes();
  const valid = packageEvidence(before, packageBytes({ replaceFirst: true }), 1024 * 1024, {
    mediaChanges: [{
      action: "replace_image",
      partName: "media/one.png",
      replacementHash: sha256(Uint8Array.from([7, 8, 9])),
    }],
  });
  assert.deepEqual(valid.binaryChanges.allowedChangedParts, ["word/media/one.png"]);
  assert.equal(valid.binaryChanges.preserveAllOtherBinaryMembers, true);
  assert.equal(valid.mediaRelationshipChanges.changed.length, 1);
  assert.equal(
    valid.mediaRelationshipChanges.changed[0].after.part,
    "word/media/three.png",
  );

  assert.throws(
    () =>
      packageEvidence(
        before,
        packageBytes({ replaceFirst: true, dropSecond: true }),
        1024 * 1024,
        {
          mediaChanges: [{
            action: "replace_image",
            partName: "media/one.png",
            replacementHash: sha256(Uint8Array.from([7, 8, 9])),
          }],
        },
      ),
    /unrelated binary package member.*two\.png/,
  );
  assert.throws(
    () =>
      packageEvidence(before, packageBytes({ orphan: true }), 1024 * 1024),
    /orphaned media.*orphan\.png/,
  );
  assert.throws(
    () =>
      packageEvidence(
        before,
        packageBytes({ replaceFirst: true }),
        1024 * 1024,
        {
          mediaChanges: [{
            action: "replace_image",
            partName: "media/one.png",
            replacementHash: sha256(Uint8Array.from([99])),
          }],
        },
      ),
    /unauthorized binary package members|exact supplied image bytes/,
  );
  assert.throws(
    () =>
      packageEvidence(
        before,
        packageBytes({ retargetSecond: true }),
        1024 * 1024,
        {
          mediaChanges: [{
            action: "replace_image",
            partName: "media/one.png",
            replacementHash: sha256(Uint8Array.from([1, 2, 3])),
          }],
        },
      ),
    /orphaned media|unauthorized package relationship/,
  );
});
