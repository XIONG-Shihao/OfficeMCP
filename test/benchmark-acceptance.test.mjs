import test from "./support/test.mjs";
import assert from "node:assert/strict";
import { zipSync, strToU8 } from "fflate";
import { benchmarkVerification, benchmarkSemanticPass } from "../scripts/benchmark-acceptance.mjs";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const p = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table = (price = "125", other = "9") => `<w:tbl><w:tr><w:tc>${p("Service")}</w:tc><w:tc>${p(price)}</w:tc></w:tr><w:tr><w:tc>${p("Other")}</w:tc><w:tc>${p(other)}</w:tc></w:tr></w:tbl>`;
function fixture({ body = table(), parts = {} } = {}) {
  return zipSync({
    "word/document.xml": strToU8(`<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`),
    ...Object.fromEntries(Object.entries(parts).map(([name, value]) => [name, typeof value === "string" ? strToU8(value) : value])),
  });
}
const verify = (originalBytes, outputBytes) => benchmarkVerification({ originalBytes, outputBytes });

test("table acceptance rejects document-wide price decoys and unrelated cell changes", () => {
  const original = fixture();
  assert.equal(verify(original, fixture({ body: table() + p("Service 150") + p("Service fees") })).tableUpdated, false);
  assert.equal(verify(original, fixture({ body: table("150", "10") + p("Service fees") })).tableUpdated, false);
  const correct = verify(original, fixture({ body: table("150") + p("Service fees") }));
  assert.equal(correct.tableUpdated, true);
  assert.equal(correct.unrelatedTableCellsPreserved, true);
  assert.equal(correct.captionAdded, true);
  assert.equal(verify(original, fixture({ body: table("150") + p("Unrelated paragraph") + p("Service fees") })).captionAdded, false);
});

function commentFixture({ resolved = "1", resolutionId = "ABC123", anchorText = "Review clause remains effective." } = {}) {
  return fixture({
    body: `<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>${anchorText}</w:t></w:r><w:commentRangeEnd w:id="2"/></w:p>`,
    parts: {
      "word/comments.xml": `<w:comments xmlns:w="${W}" xmlns:w14="${W14}"><w:comment w:id="2"><w:p w14:paraId="ABC123"><w:r><w:t>Approved by legal</w:t></w:r></w:p></w:comment></w:comments>`,
      "word/commentsExtended.xml": `<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="${resolutionId}" w15:done="${resolved}"/></w15:commentsEx>`,
    },
  });
}

test("existing comment acceptance requires unchanged text, anchor and status", () => {
  const original = commentFixture({ resolved: "0" });
  const unchanged = verify(original, commentFixture({ resolved: "0" }));
  assert.equal(unchanged.existingComment, true);
  assert.equal(unchanged.commentsPreserved, true);
  assert.equal(unchanged.existingCommentResolved, false);
  assert.equal(verify(original, commentFixture()).commentsPreserved, false);
  assert.equal(verify(original, commentFixture({ anchorText: "Wrong paragraph" })).commentsPreserved, false);
});

function imageFixture({ size = 152400, shapeSize = size, bytes = Uint8Array.from([1, 2, 3]), extra = false, directory = false, alt = "Company mark" } = {}) {
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  return fixture({
    body: `<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="${R}"><wp:inline><wp:extent cx="${size}" cy="${size}"/><wp:docPr id="1" descr="${alt}"/><a:graphic><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="${shapeSize}" cy="${shapeSize}"/></a:xfrm></pic:spPr></pic:pic></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
    parts: {
      "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/image" Target="media/image.png"/></Relationships>`,
      "word/media/image.png": bytes,
      ...(directory ? { "word/media/": new Uint8Array() } : {}),
      ...(extra ? { "word/media/extra.png": Uint8Array.from([4, 5, 6]) } : {}),
    },
  });
}

test("image acceptance checks effective dimensions, relationship target and binary hashes", () => {
  const original = imageFixture();
  const unchangedSize = verify(original, imageFixture({ alt: "Updated company mark" }));
  assert.equal(unchangedSize.imageAltUpdated, true);
  assert.equal(unchangedSize.imageDimensionsUpdated, false);
  const correct = verify(original, imageFixture({ size: 304800, alt: "Updated company mark" }));
  assert.equal(correct.imageDimensionsUpdated, true);
  assert.equal(correct.imageBytesPreserved, true);
  assert.equal(verify(original, imageFixture({ size: 304800, shapeSize: 152400 })).imageDimensionsUpdated, false);
  assert.equal(verify(original, imageFixture({ size: 304800, bytes: Uint8Array.from([9, 9, 9]) })).imageBytesPreserved, false);
  assert.equal(verify(original, imageFixture({ size: 304800, extra: true })).imageBytesPreserved, false);
  assert.equal(verify(imageFixture({ directory: true }), imageFixture({ size: 304800 })).imageBytesPreserved, true, "ZIP directory records are not image bytes");
});

test("bold deadline acceptance rejects explicitly disabled formatting", () => {
  const payment = (value, bold) => `<w:p><w:r><w:t>Payment within </w:t></w:r><w:r><w:rPr><w:b w:val="${bold}"/></w:rPr><w:t>${value}</w:t></w:r><w:r><w:t> days.</w:t></w:r></w:p>`;
  const original = fixture({ body: payment("30", "1") });
  assert.equal(verify(original, fixture({ body: payment("45", "0") })).boldNumberPreserved, false);
  assert.equal(verify(original, fixture({ body: payment("45", "1") })).boldNumberPreserved, true);
});

test("semantic acceptance requires successful completion, source preservation and every scenario signal", () => {
  const scenario = { expectedSignals: ["outputExists", "commentsPreserved"] };
  const result = { exit: 0, timedOut: false, originalUnchanged: true, toolCalls: 1, verification: { outputExists: true, commentsPreserved: true } };
  assert.equal(benchmarkSemanticPass(result, scenario), true);
  for (const changed of [{ exit: 1 }, { timedOut: true }, { originalUnchanged: false }, { toolCalls: 0 }, { verification: { outputExists: true } }])
    assert.equal(benchmarkSemanticPass({ ...result, ...changed }, scenario), false);
  assert.equal(verify(fixture(), Uint8Array.from([1, 2, 3])).invalidOutput, true);
});
