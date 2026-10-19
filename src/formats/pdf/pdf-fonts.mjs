import { readFileSync } from "node:fs";
import { StandardFonts } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import { requireThat } from "../../infrastructure/errors.mjs";
export async function documentFonts(pdf, texts) {
  if (texts.some((t) => /[^\x20-\x7e\r\n\t]/.test(t))) {
    pdf.registerFontkit(fontkit);
    const font = await pdf.embedFont(
      readFileSync(new URL("../../assets/fonts/NotoSansSC.ttf", import.meta.url)),
      { subset: true },
    );
    const supported = new Set(font.getCharacterSet());
    requireThat(
      texts.every((t) =>
        Array.from(t).every(
          (c) => "\r\n\t".includes(c) || supported.has(c.codePointAt(0)),
        ),
      ),
      "UNSUPPORTED_GLYPH",
      "The bundled font cannot represent a requested character; use a supported font or revise that character.",
    );
    return { regular: font, bold: font };
  }
  return {
    regular: await pdf.embedFont(StandardFonts.Helvetica),
    bold: await pdf.embedFont(StandardFonts.HelveticaBold),
  };
}
