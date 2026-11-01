import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { Canvas, FontLibrary as fontLibrary } from "skia-canvas";
import { createPptxTextLayout } from "./pptxTextLayout";

const parse = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  trimValues: false,
  parseTagValue: false,
});
type Node = Record<string, unknown>;
function find(nodes: Node[], name: string): Node[] {
  return nodes.flatMap((node) =>
    Object.entries(node).flatMap(([key, value]) =>
      key === ":@"
        ? []
        : [
            ...(key === name ? [node] : []),
            ...(Array.isArray(value) ? find(value as Node[], name) : []),
          ],
    ),
  );
}
function content(node: Node): string {
  return Object.entries(node)
    .filter(([key]) => key !== ":@")
    .map(([key, value]) =>
      key === "#text"
        ? String(value)
        : Array.isArray(value)
          ? (value as Node[]).map(content).join("")
          : "",
    )
    .join("");
}
function visible(pages: string[]): string {
  return pages
    .flatMap((page) => find(parse.parse(page), "a:t").map(content))
    .join("");
}
const escape = (value: string) =>
  value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
function run(value: string, family = "Arial", extra = ""): string {
  return `<a:r><a:rPr sz="4800" b="1" ${extra}><a:latin typeface="${family}"/></a:rPr><a:t>${escape(value)}</a:t></a:r>`;
}
function paragraph(
  value: string,
  properties = "",
  family = "Arial",
  extra = "",
): string {
  return `<a:p><a:pPr marL="0" marR="0" indent="0">${properties}</a:pPr>${run(value, family, extra)}</a:p>`;
}
function shape(
  paragraphs: string,
  width: number,
  height: number,
  id = 1,
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Measured ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width * 12700}" cy="${height * 12700}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}
function slide(
  paragraphs: string,
  width = 500,
  height = 230,
  extraShape = "",
): string {
  return `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>${shape(paragraphs, width, height)}${extraShape}</p:spTree></p:cSld></p:sld>`;
}
function validateGeometry(
  pages: string[],
  widths: number | number[],
  heights: number | number[],
): void {
  const ctx = new Canvas(1, 1).getContext("2d");
  pages.forEach((page, pageIndex) => {
    assert.equal(XMLValidator.validate(page), true);
    const parsed = parse.parse(page) as Node[];
    const width = Array.isArray(widths) ? widths[pageIndex ? 1 : 0] : widths;
    const height = Array.isArray(heights)
      ? heights[pageIndex ? 1 : 0]
      : heights;
    for (const tx of find(parsed, "p:txBody")) {
      let used = 0;
      for (const p of find([tx], "a:p")) {
        const pPr = find([p], "a:pPr")[0];
        const propertyTags = (pPr["a:pPr"] as Node[]).map((node) =>
          Object.keys(node).find((key) => key !== ":@"),
        );
        if (
          propertyTags.includes("a:buNone") &&
          propertyTags.includes("a:defRPr")
        ) {
          assert.ok(
            propertyTags.indexOf("a:buNone") < propertyTags.indexOf("a:defRPr"),
            "PowerPoint requires the bullet override before default run properties",
          );
        }
        const spacing = (name: string): number => {
          const points = find(find([pPr], name), "a:spcPts")[0];
          return (
            Number((points?.[":@"] as Record<string, string>)?.["@_val"] ?? 0) /
            100
          );
        };
        const lineHeight = spacing("a:lnSpc");
        let lineWidth = 0,
          lineCount = 1;
        const entries = p["a:p"] as Node[];
        for (const entry of entries) {
          if (entry["a:br"]) {
            assert.ok(lineWidth <= width + 0.02, `${lineWidth} > ${width}`);
            lineWidth = 0;
            lineCount++;
          }
          if (!entry["a:r"]) continue;
          const properties = find([entry], "a:rPr")[0];
          const attributes = properties[":@"] as Record<string, string>;
          const latin = find([properties], "a:latin")[0][":@"] as Record<
            string,
            string
          >;
          const family = latin["@_typeface"];
          ctx.font = `${attributes["@_i"] === "1" ? "italic" : "normal"} ${family === "Arial Rounded MT Bold" ? 400 : 700} 48px "${family}"`;
          ctx.letterSpacing = `${Number(attributes["@_spc"] ?? 0) / 100}px`;
          const value = content(find([entry], "a:t")[0]);
          lineWidth += ctx.measureText(value).width;
        }
        assert.ok(lineWidth <= width + 0.02, `${lineWidth} > ${width}`);
        used +=
          spacing("a:spcBef") + lineCount * lineHeight + spacing("a:spcAft");
      }
      assert.ok(
        used <= height + 0.02,
        `page ${pageIndex + 1}: ${used} > ${height}`,
      );
    }
  });
}

const dirs = process.env.ENGLISH_MASS_FONT_DIR
  ? [process.env.ENGLISH_MASS_FONT_DIR]
  : [
      path.join(process.cwd(), "app/fonts/english-mass"),
      "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts",
      "/Applications/Microsoft Word.app/Contents/Resources/DFonts",
      "/Library/Fonts",
      "/System/Library/Fonts/Supplemental",
    ];
const faceFiles = [
  ["arialbd.ttf", "Arial Bold.ttf"],
  ["arialbi.ttf", "Arial Bold Italic.ttf"],
  ["Book Antiqua Bold Italic.ttf"],
  ["ArialRoundedMTBold.ttf", "Arial Rounded Bold.ttf"],
].map((names) =>
  dirs
    .flatMap((dir) => names.map((name) => path.join(dir, name)))
    .find((file) => fs.existsSync(file)),
);
const missingFonts = faceFiles.some((file) => !file)
  ? "Requires exact licensed Arial bold/bold italic, Book Antiqua bold italic, and Arial Rounded MT Bold; provision ENGLISH_MASS_FONT_DIR."
  : false;
if (!missingFonts) fontLibrary.use(faceFiles as string[]);

test(
  "English Mass uses exact shaped font geometry, not equal character counts",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const source = slide(paragraph("(teks)"), 310, 180);
    const narrow = "iiiiiiii ".repeat(12),
      wide = "WWWWWWWW ".repeat(12);
    assert.equal(narrow.length, wide.length);
    const narrowPages = layout.paginate("slide1.xml", source, "teks", [narrow]);
    const widePages = layout.paginate("slide1.xml", source, "teks", [wide]);
    assert.ok(widePages.length > narrowPages.length);
    assert.equal(visible(narrowPages), narrow.trim());
    assert.equal(visible(widePages), wide.trim());
    validateGeometry(narrowPages, 310, 180);
    validateGeometry(widePages, 310, 180);
  },
);

test(
  "English Mass preserves paragraphs and hard breaks without blank-line double charging",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const pages = layout.paginate(
      "slide1.xml",
      slide(
        paragraph(
          "(teks)",
          '<a:spcBef><a:spcPts val="500"/></a:spcBef><a:spcAft><a:spcPts val="700"/></a:spcAft>',
        ),
        900,
        500,
      ),
      "teks",
      ["Alpha\nBeta\n\nGamma\r\nDelta"],
    );
    assert.equal(pages.length, 1);
    assert.equal(find(parse.parse(pages[0]), "a:p").length, 2);
    assert.equal(find(parse.parse(pages[0]), "a:br").length, 2);
    assert.equal(visible(pages), "AlphaBetaGammaDelta");
    validateGeometry(pages, 900, 500);
  },
);

test(
  "English Mass shares one shape budget across multiple placeholders and responses",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const first = "First verse is measured. ".repeat(11).trim();
    const second = "The refrain must follow exactly once. ".repeat(6).trim();
    const source = slide(
      paragraph("HEADING") + paragraph("(teks)") + paragraph("(teks)"),
      500,
      230,
    );
    const pages = layout.paginate("slide1.xml", source, "teks", [
      first,
      second,
    ]);
    assert.ok(pages.length > 2);
    assert.equal(visible(pages), `HEADING${first}${second}`);
    assert.equal(visible(pages).split("HEADING").length, 2);
    validateGeometry(pages, 500, 230);
  },
);

test(
  "English Mass preserves split-run token prefix/suffix style and XML entities",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const source = slide(
      `<a:p><a:pPr/>${run("MC & <leader>: (")}${run("te", "Arial", 'i="1"')}${run("ks)")}${run(" — Amen & thanks", "Arial", 'i="1"')}</a:p>`,
      500,
      230,
      shape(paragraph("UNTOUCHED & <shape>"), 500, 230, 2),
    );
    const originalStatic = source.match(/<p:sp>[\s\S]*?<\/p:sp>/g)![1];
    const body = "Text & truth <remain> exactly once. ".repeat(9).trim();
    const pages = layout.paginate("slide1.xml", source, "teks", [body]);
    const dynamicText = pages
      .map((page) => content(find(parse.parse(page), "p:sp")[0]))
      .join("");
    assert.equal(dynamicText, `MC & <leader>: ${body} — Amen & thanks`);
    assert.ok(pages.every((page) => page.includes(originalStatic)));
    assert.ok(pages.join("").includes("&lt;remain&gt;"));
    const suffix = pages
      .flatMap((page) => find(parse.parse(page), "a:r"))
      .find((node) => content(node).includes("thanks"));
    assert.equal(
      (find([suffix!], "a:rPr")[0][":@"] as Record<string, string>)["@_i"],
      "1",
    );
  },
);

test(
  "English Mass removes decorative title spacers and continues metadata before body",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const title = "The exact title and its rich epigraph ".repeat(10).trim();
    const citation = "Book of Wisdom chapter three";
    const source = slide(
      paragraph("FIRST READING") +
        paragraph("") +
        paragraph("(teks)", "", "Book Antiqua", 'i="1"') +
        paragraph("") +
        paragraph("A reading from the (teks)"),
      600,
      250,
    );
    const continuation = slide(paragraph("(teks)"), 600, 250);
    const pages = layout.paginate(
      "slide39.xml",
      source,
      "teks",
      [title, citation],
      { continuationXml: continuation, firstOnlyTexts: ["FIRST READING"] },
    );
    assert.ok(pages.length > 1);
    assert.equal(
      visible(pages),
      `FIRST READING${title}A reading from the ${citation}`,
    );
    assert.ok(pages.slice(1).every((page) => !page.includes("FIRST READING")));
    assert.ok(
      find(parse.parse(pages[0]), "a:latin").some(
        (node) =>
          (node[":@"] as Record<string, string>)["@_typeface"] ===
          "Book Antiqua",
      ),
    );
    validateGeometry(pages, 600, 250);
  },
);

test(
  "English Mass keeps first/last policies once and never repeats dialogue in the body",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const source = slide(
      paragraph("OPENING PRAYER") +
        paragraph("MC: Let us pray.") +
        paragraph("MC: (teks)") +
        paragraph("Cs: Amen."),
      550,
      250,
    );
    const body =
      "A prayer with enough measured lines to continue several times. "
        .repeat(10)
        .trim();
    const pages = layout.paginate("slide37.xml", source, "teks", [body], {
      firstOnlyTexts: ["OPENING PRAYER", "MC: Let us pray.", "MC: (teks)"],
      lastOnlyTexts: ["Cs: Amen."],
    });
    assert.equal(
      visible(pages),
      `OPENING PRAYERMC: Let us pray.MC: ${body}Cs: Amen.`,
    );
    assert.ok(
      pages
        .slice(1)
        .every(
          (page) =>
            !page.includes("OPENING PRAYER") && !page.includes("Let us pray."),
        ),
    );
    assert.ok(pages.slice(0, -1).every((page) => !page.includes("Cs: Amen.")));
    assert.ok(pages[pages.length - 1].includes("Cs: Amen."));
    validateGeometry(pages, 550, 250);
  },
);

test(
  "English Mass empty/exact-fit input emits no unexpected blank pages",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const ctx = new Canvas(1, 1).getContext("2d");
    ctx.font = 'normal 700 48px "Arial"';
    const metrics = ctx.measureText("One line");
    const height =
      Math.ceil(
        (metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent) * 100,
      ) / 100;
    for (const value of ["", "One line"]) {
      const pages = layout.paginate(
        "slide1.xml",
        slide(paragraph("(teks)"), 900, height),
        "teks",
        [value],
      );
      assert.equal(pages.length, 1);
      assert.equal(visible(pages), value);
      validateGeometry(pages, 900, height);
    }
  },
);

test(
  "English Mass oversized words split only at graphemes; no-break groups fail visibly",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const source = slide(paragraph("(teks)"), 180, 180);
    const word = "e\u0301".repeat(80);
    const pages = layout.paginate("slide1.xml", source, "teks", [word]);
    assert.equal(visible(pages), word);
    assert.ok(
      pages
        .flatMap((page) => find(parse.parse(page), "a:t"))
        .every((node) => !/^\p{Mark}/u.test(content(node))),
    );
    assert.throws(
      () => layout.paginate("slide1.xml", source, "teks", ["WWWW\u00a0WWWW"]),
      /nonbreaking text group exceeds/,
    );
    assert.throws(
      () =>
        layout.paginate(
          "slide1.xml",
          slide(paragraph("(teks)"), 900, 10),
          "teks",
          ["Impossible"],
        ),
      /one text line.*cannot fit/,
    );
    validateGeometry(pages, 180, 180);
  },
);

test(
  "English Mass supports differing continuation geometry and independent touched shapes",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const source = slide(paragraph("(teks)"), 600, 180);
    const continuation = slide(paragraph("(teks)"), 350, 180);
    const body =
      "Words that continue in a narrower box without losing content. "
        .repeat(14)
        .trim();
    const pages = layout.paginate("slide1.xml", source, "teks", [body], {
      continuationXml: continuation,
    });
    assert.equal(visible(pages), body);
    validateGeometry(pages, [600, 350], 180);
    const two = slide(
      paragraph("(teks)"),
      500,
      230,
      shape(paragraph("(teks)"), 500, 230, 2),
    );
    const pair = layout.paginate("slide1.xml", two, "teks", [
      "first shape",
      "second shape",
    ]);
    assert.equal(visible(pair), "first shapesecond shape");
  },
);

test(
  "English Mass resolves real template hierarchy and leaves unrelated song lyric placeholder intact",
  { skip: missingFonts },
  async () => {
    const zip = await JSZip.loadAsync(
      fs.readFileSync(
        path.join(process.cwd(), "app/template/EM_Template.pptx"),
      ),
    );
    const layout = await createPptxTextLayout(zip);
    const original = await zip.file("ppt/slides/slide14.xml")!.async("string");
    const title = "A measured song title & refrain ".repeat(15).trim();
    const pages = layout.paginate("slide14.xml", original, "judul", [title]);
    assert.equal(visible(pages), `${title}(teks)`);
    assert.ok(pages.length > 1);
    assert.equal(
      await zip.file("ppt/slides/slide14.xml")!.async("string"),
      original,
      "layout must not mutate the input ZIP",
    );
    assert.ok(pages.every((page) => XMLValidator.validate(page) === true));
    const reading = await zip.file("ppt/slides/slide40.xml")!.async("string");
    const readingPages = layout.paginate("slide40.xml", reading, "teks", [
      "The reading remains plain prose, without inherited bullets. "
        .repeat(12)
        .trim(),
    ]);
    assert.ok(find(parse.parse(readingPages[0]), "a:buNone").length > 0);
    assert.equal(find(parse.parse(readingPages[0]), "a:buChar").length, 0);
    validateGeometry(readingPages, 945.6, 532.8);
  },
);

test(
  "English Mass carries one tail line with the final response and preserves source hard breaks",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const ctx = new Canvas(1, 1).getContext("2d");
    ctx.font = 'normal 700 48px "Arial"';
    const metrics = ctx.measureText("Alpha");
    const lineHeight =
      Math.ceil(
        (metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent) * 100,
      ) / 100;
    const body = `<a:p><a:pPr/>${run("(teks)")}</a:p>`;
    const source = slide(body + paragraph("Cs: Amen."), 900, lineHeight * 3);
    const pages = layout.paginate(
      "slide1.xml",
      source,
      "teks",
      ["Alpha\nBeta\nGamma"],
      { lastOnlyTexts: ["Cs: Amen."] },
    );
    assert.equal(pages.length, 2);
    assert.equal(visible([pages[0]]), "AlphaBeta");
    assert.equal(visible([pages[1]]), "GammaCs: Amen.");
    assert.equal(find(parse.parse(pages[0]), "a:br").length, 1);
    assert.equal(find(parse.parse(pages[1]), "a:br").length, 0);
    assert.ok(pages.every((page) => page.includes('xml:space="preserve"')));
    validateGeometry(pages, 900, lineHeight * 3);
    const originalBreak = slide(
      `<a:p><a:pPr/>${run("Before")}<a:br/>${run("(teks)")}</a:p>`,
      900,
      lineHeight * 3,
    );
    const originalBreakPages = layout.paginate(
      "slide1.xml",
      originalBreak,
      "teks",
      ["After"],
    );
    assert.equal(visible(originalBreakPages), "BeforeAfter");
    assert.equal(find(parse.parse(originalBreakPages[0]), "a:br").length, 1);
    const narrower = slide(body + paragraph("Cs: Amen."), 180, lineHeight * 3);
    const carried = layout.paginate(
      "slide1.xml",
      source,
      "teks",
      ["Alpha\nBeta\nGamma Gamma Gamma"],
      {
        continuationXml: narrower,
        lastOnlyTexts: ["Cs: Amen."],
      },
    );
    assert.equal(visible(carried), "AlphaBetaGamma Gamma GammaCs: Amen.");
    assert.ok(
      carried
        .slice(0, -1)
        .every((page) => !visible([page]).includes("Cs: Amen.")),
    );
    validateGeometry(carried, [900, 180], lineHeight * 3);
  },
);

test(
  "English Mass inherits real font/spacing/extents and rejects grouped dynamic shapes",
  { skip: missingFonts },
  async () => {
    const zip = new JSZip();
    const rel = (type: string, target: string) =>
      `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/></Relationships>`;
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      rel("slideLayout", "../slideLayouts/slideLayout1.xml"),
    );
    zip.file(
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
      rel("slideMaster", "../slideMasters/slideMaster1.xml"),
    );
    zip.file(
      "ppt/slideLayouts/slideLayout1.xml",
      '<p:sldLayout><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>',
    );
    zip.file(
      "ppt/slideMasters/slideMaster1.xml",
      `<p:sldMaster><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:ext cx="${520 * 12700}" cy="${250 * 12700}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr lIns="127000" rIns="127000" tIns="127000" bIns="127000"/><a:lstStyle/></p:txBody></p:sp></p:spTree></p:cSld><p:txStyles><p:bodyStyle><a:lvl1pPr marL="0" indent="0"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:defRPr sz="4800" b="1"><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>`,
    );
    const source =
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>(teks)</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>';
    const layout = await createPptxTextLayout(zip);
    const body = "Inherited exact typography and geometry. ".repeat(10).trim();
    const pages = layout.paginate("slide1.xml", source, "teks", [body]);
    assert.equal(visible(pages), body);
    validateGeometry(pages, 500, 230);
    const grouped = source
      .replace("<p:spTree>", "<p:spTree><p:grpSp>")
      .replace("</p:spTree>", "</p:grpSp></p:spTree>");
    assert.throws(
      () => layout.paginate("slide1.xml", grouped, "teks", ["Grouped"]),
      /Ungroup.*actual slide-space extents/,
    );
  },
);

test(
  "template gaps survive while fitting and are reduced only as much as overflow requires",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip());
    const ctx = new Canvas(1, 1).getContext("2d");
    ctx.font = 'normal 700 48px "Arial"';
    const metrics = ctx.measureText("Alpha");
    const lineHeight =
      Math.ceil(
        (metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent) * 100,
      ) / 100;
    const paragraphs =
      paragraph("HEADING") +
      paragraph("") +
      paragraph("") +
      paragraph("(teks)") +
      paragraph("") +
      paragraph("(teks)");
    for (const [capacity, remainingGaps, pageCount] of [
      [6, 3, 1],
      [5, 2, 1],
      [3, 0, 1],
      [2, 0, 2],
    ]) {
      const pages = layout.paginate(
        "slide1.xml",
        slide(paragraphs, 900, lineHeight * capacity),
        "teks",
        ["Alpha", "Beta"],
        { preserveTemplateSpacers: true, firstOnlyTexts: ["HEADING"] },
      );
      const output = pages.flatMap((page) => find(parse.parse(page), "a:p"));
      assert.equal(pages.length, pageCount);
      assert.equal(
        output.filter((p) => !content(p).trim()).length,
        remainingGaps,
      );
      assert.equal(visible(pages), "HEADINGAlphaBeta");
      if (capacity === 5) {
        assert.deepEqual(
          output.map(content),
          ["HEADING", "", "Alpha", "", "Beta"],
          "reduce the double gap first, keeping a gap between each section",
        );
      }
      validateGeometry(pages, 900, lineHeight * capacity);
    }
    const hardBreaks = layout.paginate(
      "slide1.xml",
      slide(paragraphs, 900, lineHeight * 4),
      "teks",
      ["Alpha\nBeta", "Gamma"],
      { preserveTemplateSpacers: true, firstOnlyTexts: ["HEADING"] },
    );
    assert.equal(visible(hardBreaks), "HEADINGAlphaBetaGamma");
    assert.equal(
      find(parse.parse(hardBreaks[0]), "a:br").length,
      1,
      "a source hard break is not a disposable template gap",
    );
    assert.equal(
      find(parse.parse(hardBreaks[0]), "a:p").filter((p) => !content(p).trim())
        .length,
      0,
    );
    validateGeometry(hardBreaks, 900, lineHeight * 4);
  },
);

test("Indonesian font profile is exclusive and captured per layout instance", async () => {
  const previous = process.env.INDONESIAN_MASS_FONT_DIR;
  process.env.INDONESIAN_MASS_FONT_DIR = path.join(
    process.cwd(),
    "missing-indonesian-fonts",
  );
  let layout;
  try {
    layout = await createPptxTextLayout(new JSZip(), {
      fontProfile: "indonesian",
    });
  } finally {
    if (previous === undefined) delete process.env.INDONESIAN_MASS_FONT_DIR;
    else process.env.INDONESIAN_MASS_FONT_DIR = previous;
  }
  assert.throws(
    () =>
      layout.paginate("slide1.xml", slide(paragraph("(teks)")), "teks", [
        "Doa",
      ]),
    /Indonesian Mass.*INDONESIAN_MASS_FONT_DIR.*Font substitution is disabled/,
  );
});

test(
  "structured paragraphs retain token color, prefix, suffix and explicit face overrides",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip(), {
      fontProfile: "indonesian",
    });
    const source = slide(
      `<a:p><a:pPr/>${run("MC: (teks) — selesai").replace("<a:latin", '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin')}</a:p>`,
      600,
      220,
    );
    const body = "Isi doa yang harus lengkap. ".repeat(16).trim();
    const pages = layout.paginate("slide1.xml", source, "teks", [
      {
        paragraphs: [
          { text: "Judul", bold: true },
          { text: body, bold: false },
        ],
      },
    ]);
    assert.equal(visible(pages), `MC: Judul${body} — selesai`);
    assert.ok(pages.length > 1);
    const runs = pages.flatMap((page) => find(parse.parse(page), "a:r"));
    for (const entry of runs) {
      const properties = find([entry], "a:rPr")[0];
      assert.equal(
        (find([properties], "a:srgbClr")[0][":@"] as Record<string, string>)[
          "@_val"
        ],
        "FF0000",
      );
      if (content(entry).includes("Isi")) {
        assert.equal((properties[":@"] as Record<string, string>)["@_b"], "0");
      }
    }
    assert.equal(visible(pages).split("MC: ").length - 1, 1);
  },
);

test(
  "dynamic final response shares a final budget or spans ordered pages without loss",
  { skip: missingFonts },
  async () => {
    const layout = await createPptxTextLayout(new JSZip(), {
      fontProfile: "indonesian",
    });
    const source = slide(
      paragraph("(teks)") +
        paragraph("(teks)").replace(
          "<a:latin",
          '<a:solidFill><a:srgbClr val="00AA00"/></a:solidFill><a:latin',
        ),
      500,
      200,
    );
    const body = "Intensi umat. ".repeat(25).trim();
    for (const response of [
      "U\tKabulkanlah doa kami.",
      "U\t" + "Jawaban umat lengkap. ".repeat(45).trim(),
    ]) {
      const pages = layout.paginate("slide1.xml", source, "teks", [
        { paragraphs: [{ text: body, bold: false }] },
        {
          paragraphs: [{ text: response, bold: false, italic: true }],
          policy: "last",
        },
      ]);
      assert.equal(visible(pages), body + response);
      const responseRuns = pages
        .flatMap((page) => find(parse.parse(page), "a:r"))
        .filter((entry) => find([entry], "a:srgbClr").length);
      for (const entry of responseRuns) {
        const attributes = find([entry], "a:rPr")[0][":@"] as Record<
          string,
          string
        >;
        assert.equal(attributes["@_i"], "1");
        assert.equal(attributes["@_b"], "0");
      }
      pages.forEach((page) => assert.equal(XMLValidator.validate(page), true));
    }
  },
);

const indonesianDirs = process.env.INDONESIAN_MASS_FONT_DIR
  ? [process.env.INDONESIAN_MASS_FONT_DIR]
  : [
      path.join(process.cwd(), "app/fonts/indonesian-mass"),
      "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts",
      "/System/Library/Fonts/Supplemental",
    ];
const missingWeeklyFaces = [
  "Arial Narrow Bold.ttf",
  "Wingdings.ttf",
  "Lucida Sans Demibold Italic.ttf",
  "Arial.ttf",
  "Arial Bold.ttf",
].some(
  (name) => !indonesianDirs.some((dir) => fs.existsSync(path.join(dir, name))),
);
test(
  "weekly metadata measures symbol-only Wingdings and mixed 60/48/44pt exact faces",
  { skip: missingWeeklyFaces ? "Requires licensed weekly font faces" : false },
  async () => {
    const layout = await createPptxTextLayout(new JSZip(), {
      fontProfile: "indonesian",
    });
    const source = slide(
      '<a:p><a:pPr/><a:r><a:rPr sz="6000"><a:sym typeface="Wingdings"/></a:rPr><a:t>\uF026</a:t></a:r>' +
        run(" (teks)", "Arial Narrow").replace('sz="4800"', 'sz="6000"') +
        "</a:p>" +
        paragraph("Epigraf", "", "Lucida Sans", 'i="1"') +
        paragraph("hening").replace('sz="4800"', 'sz="4400"'),
      960,
      540,
    );
    const pages = layout.paginate("slide35.xml", source, "teks", [
      "Bacaan dari Kitab Kebijaksanaan",
    ]);
    assert.equal(pages.length, 1);
    assert.equal(
      visible(pages),
      "\uF026 Bacaan dari Kitab KebijaksanaanEpigrafhening",
    );
    assert.match(pages[0], /a:sym typeface="Wingdings"/);
    for (const size of ["6000", "4800", "4400"])
      assert.ok(pages[0].includes(`sz="${size}"`));
    assert.equal(XMLValidator.validate(pages[0]), true);
  },
);
