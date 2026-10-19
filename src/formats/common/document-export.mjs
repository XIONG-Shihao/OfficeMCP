import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { pathToFileURL, fileURLToPath } from "node:url";
import * as mupdf from "mupdf";
import { requireThat } from "../../infrastructure/errors.mjs";
import { regularFile } from "../../infrastructure/files.mjs";
import { runLocal } from "../../engines/local-process.mjs";

const escapeXml = (text) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");

/** Document conversion only: no page images, layout maps, or visual comparison. */
export async function exportDocumentPdf(task) {
  requireThat(task.format === "docx", "UNSUPPORTED_FORMAT", "This export operation accepts DOCX input only.");
  requireThat(task.engines?.libreOffice, "ENGINE_UNAVAILABLE", "Configure engines.libreOffice before converting this document to PDF.");
  const profile = path.join(task.scratch, "export-pdf");
  const output = path.join(task.scratch, "export-pdf-output");
  const user = path.join(profile, "user");
  const fontDirectory = fileURLToPath(new URL("../../assets/fonts/", import.meta.url));
  const fontConfig = path.join(profile, "fonts.conf");
  mkdirSync(user, { recursive: true });
  mkdirSync(output, { recursive: true });
  writeFileSync(path.join(user, "registrymodifications.xcu"),
    '<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item></oor:items>');
  const fonts = [fontDirectory,
    ...["/System/Library/Fonts", "/Library/Fonts", "/usr/share/fonts", "/usr/local/share/fonts"].filter(existsSync)];
  writeFileSync(fontConfig,
    '<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>' +
    (existsSync("/etc/fonts/fonts.conf") ? '<include ignore_missing="yes">/etc/fonts/fonts.conf</include>' : "") +
    fonts.map((directory) => `<dir>${escapeXml(directory)}</dir>`).join("") +
    `<cachedir>${escapeXml(path.join(profile, "font-cache"))}</cachedir></fontconfig>`);
  await runLocal(task.engines.libreOffice, [
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    "--headless", "--convert-to", "pdf", "--outdir", output, task.input,
  ], { env: { ...process.env, FONTCONFIG_FILE: fontConfig, SAL_FONTPATH_PRIVATE: fontDirectory } });
  const converted = path.join(output, path.basename(task.input, path.extname(task.input)) + ".pdf");
  requireThat(existsSync(converted), "ENGINE_FAILED", "LibreOffice did not produce PDF output.");
  const bytes = regularFile(converted, task.maxFileBytes ?? task.maxExpandedBytes);
  const document = new mupdf.PDFDocument(bytes);
  let pageCount;
  try {
    document.disableJS();
    pageCount = document.countPages();
    requireThat(pageCount > 0 && pageCount <= 2000, "LIMIT_EXCEEDED", "Exported PDF page count is invalid.");
  } finally { document.destroy(); }
  writeFileSync(task.output, bytes, { flag: "wx", mode: 0o600 });
  return { format: "pdf", size: bytes.length, pageCount,
    engine: "LibreOffice conversion + MuPDF.js 1.28.1 verification",
    notVerified: ["exact-Microsoft-Office-layout", "font-equivalence"] };
}
