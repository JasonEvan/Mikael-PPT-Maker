import fs from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { Canvas, FontLibrary as fontLibrary } from "skia-canvas";

// DrawingML uses EMUs for geometry and hundredths of a point for typography.
const EMU = 12700;
type Xml = { [name: string]: Xml[] | string | Record<string, string> };
const xmlOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
};
const parser = new XMLParser(xmlOptions);
const builder = new XMLBuilder(xmlOptions);
// CT_TextParagraphProperties has a schema-defined sequence. Inherited
// bullet overrides after defRPr can otherwise be ignored by PowerPoint.
const paragraphPropertyOrder = [
  "a:lnSpc",
  "a:spcBef",
  "a:spcAft",
  "a:buClrTx",
  "a:buClr",
  "a:buSzTx",
  "a:buSzPct",
  "a:buSzPts",
  "a:buFontTx",
  "a:buFont",
  "a:buNone",
  "a:buAutoNum",
  "a:buChar",
  "a:buBlip",
  "a:tabLst",
  "a:defRPr",
  "a:extLst",
];
const parse = (xml: string): Xml[] => parser.parse(xml) as Xml[];
const clone = <T>(value: T): T => structuredClone(value);
const tag = (node: Xml): string =>
  Object.keys(node).find((key) => key !== ":@") ?? "";
const children = (node?: Xml): Xml[] =>
  node && Array.isArray(node[tag(node)]) ? (node[tag(node)] as Xml[]) : [];
const child = (node: Xml | undefined, name: string): Xml | undefined =>
  children(node).find((entry) => tag(entry) === name);
const attrs = (node?: Xml): Record<string, string> =>
  (node?.[":@"] ?? {}) as Record<string, string>;
const attr = (node: Xml | undefined, name: string): string | undefined =>
  attrs(node)[`@_${name}`];
const element = (
  name: string,
  content: Xml[] = [],
  attributes: Record<string, string> = {},
): Xml => ({
  [name]: content,
  ...(Object.keys(attributes).length
    ? {
        ":@": Object.fromEntries(
          Object.entries(attributes).map(([key, value]) => [`@_${key}`, value]),
        ),
      }
    : {}),
});
function descendants(nodes: Xml[], name: string): Xml[] {
  return nodes.flatMap((node) => [
    ...(tag(node) === name ? [node] : []),
    ...descendants(children(node), name),
  ]);
}
function text(node?: Xml): string {
  if (!node) return "";
  if (tag(node) === "#text") return String(node["#text"]);
  if (tag(node) === "a:br") return "\n";
  return children(node).map(text).join("");
}
function merge(name: string, nodes: (Xml | undefined)[]): Xml {
  const result = element(name);
  for (const node of nodes) {
    if (!node) continue;
    result[":@"] = { ...attrs(result), ...attrs(node) };
    for (const entry of children(node)) {
      if (["a:buNone", "a:buChar", "a:buAutoNum"].includes(tag(entry))) {
        result[name] = children(result).filter(
          (candidate) =>
            !["a:buNone", "a:buChar", "a:buAutoNum"].includes(tag(candidate)),
        );
      }
      const existing = children(result).findIndex(
        (candidate) => tag(candidate) === tag(entry),
      );
      if (existing < 0) children(result).push(clone(entry));
      else children(result)[existing] = clone(entry);
    }
  }
  return result;
}

const fontNames: Record<string, string[]> = {
  "Arial|400|normal": ["arial.ttf", "Arial.ttf"],
  "Arial|700|normal": ["arialbd.ttf", "Arial Bold.ttf"],
  "Arial|400|italic": ["ariali.ttf", "Arial Italic.ttf"],
  "Arial|700|italic": ["arialbi.ttf", "Arial Bold Italic.ttf"],
  "Book Antiqua|400|normal": ["Book Antiqua.ttf"],
  "Book Antiqua|700|normal": ["Book Antiqua Bold.ttf"],
  "Book Antiqua|400|italic": ["Book Antiqua Italic.ttf"],
  "Book Antiqua|700|italic": ["Book Antiqua Bold Italic.ttf"],
  "Arial Rounded MT Bold|700|normal": [
    "ArialRoundedMTBold.ttf",
    "Arial Rounded Bold.ttf",
  ],
  "Arial Rounded MT Bold|400|normal": [
    "ArialRoundedMTBold.ttf",
    "Arial Rounded Bold.ttf",
  ],
  "Arial Narrow|400|normal": ["Arial Narrow.ttf"],
  "Arial Narrow|700|normal": ["Arial Narrow Bold.ttf"],
  "Arial Narrow|400|italic": ["Arial Narrow Italic.ttf"],
  "Arial Narrow|700|italic": ["Arial Narrow Bold Italic.ttf"],
  "Calibri|700|normal": ["Calibri Bold.ttf"],
  "Calibri|700|italic": ["Calibri Bold Italic.ttf"],
  "Lucida Sans|700|italic": ["Lucida Sans Demibold Italic.ttf"],
  "Wingdings|400|normal": ["Wingdings.ttf"],
};
type FontFace = { alias: string; weight: number; style: string };
function fontDirectories(profile: "english" | "indonesian"): string[] {
  const override = process.env[`${profile.toUpperCase()}_MASS_FONT_DIR`];
  if (override) return [path.resolve(override)];
  return [path.join(process.cwd(), "app/fonts/pptx")];
}
function exactFont(
  family: string,
  bold: boolean,
  italic: boolean,
  profile: "english" | "indonesian",
  dirs: string[],
  registeredFonts: Map<string, FontFace>,
) {
  const key = `${family}|${bold ? 700 : 400}|${italic ? "italic" : "normal"}`;
  const cacheKey = `${dirs.join(";")}|${key}`;
  const cached = registeredFonts.get(cacheKey);
  if (cached) return cached;
  const candidates = fontNames[key] ?? [];
  const file = dirs
    .flatMap((dir) => candidates.map((name) => path.join(dir, name)))
    .find((candidate) => fs.existsSync(candidate));
  if (!file)
    throw new Error(
      `${profile === "english" ? "English" : "Indonesian"} Mass layout requires the exact licensed ${family} ${bold ? "bold" : "regular"}${italic ? " italic" : ""} font. Set ${profile.toUpperCase()}_MASS_FONT_DIR to a directory containing ${candidates.join(" or ") || "this exact font face"}. Searched: ${dirs.join(", ")}. Font substitution is disabled.`,
    );
  // Register the physical face explicitly; verify its internal family/style rather
  // than trusting a filename or allowing Canvas to synthesize a missing face.
  const faces = fontLibrary.use([file]);
  const face = faces.find(
    (candidate) =>
      candidate.family === family &&
      candidate.style === (italic ? "italic" : "normal") &&
      (family === "Arial Rounded MT Bold" ||
        candidate.weight ===
          (family === "Lucida Sans" && bold ? 600 : bold ? 700 : 400)),
  );
  if (!face)
    throw new Error(
      `${profile === "english" ? "English" : "Indonesian"} Mass font ${file} does not contain the required ${key} face; provision the exact licensed font, not a renamed substitute.`,
    );
  const result = { alias: family, weight: face.weight, style: face.style };
  registeredFonts.set(cacheKey, result);
  return result;
}

type Span = { value: string; properties: Xml };
type Paragraph = {
  properties: Xml;
  end?: Xml;
  spans: Span[];
  policy: "flow" | "first" | "last";
  dynamicLast?: boolean;
};
type Geometry = { width: number; height: number; body: Xml };
type Line = {
  spans: Span[];
  height: number;
  before: number;
  after: number;
  paragraph: Paragraph;
  start: boolean;
  end: boolean;
  from: number;
};
type Shape = { raw: string; node: Xml; index: number };
type ReplacementRun = { text: string; bold?: boolean; italic?: boolean };
type ReplacementParagraph = {
  bold?: boolean;
  italic?: boolean;
  policy?: "flow" | "last";
} & (
  | { text: string; runs?: never }
  | { text?: never; runs: readonly ReplacementRun[] }
);
export type PptxTextReplacement =
  | string
  | {
      paragraphs: readonly ReplacementParagraph[];
      policy?: "flow" | "last";
    };
export interface PptxTextLayout {
  paginate(
    slideFileName: string,
    slideXml: string,
    token: string,
    replacements: readonly PptxTextReplacement[],
    options?: {
      continuationXml?: string;
      firstOnlyTexts?: readonly string[];
      lastOnlyTexts?: readonly string[];
      preserveTemplateSpacers?: boolean;
    },
  ): string[];
}
function shapes(xml: string): Shape[] {
  return [...xml.matchAll(/<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g)].map(
    (match, index) => ({ raw: match[0], node: parse(match[0])[0], index }),
  );
}
function slices(spans: Span[], start: number, end: number): Span[] {
  let offset = 0;
  return spans.flatMap((span) => {
    const from = Math.max(0, start - offset),
      to = Math.min(span.value.length, end - offset);
    offset += span.value.length;
    return to > from
      ? [{ value: span.value.slice(from, to), properties: span.properties }]
      : [];
  });
}
const spanText = (spans: Span[]): string =>
  spans.map((span) => span.value).join("");
// CRLF/CR become LF; ordinary horizontal whitespace collapses to one space and
// line-edge ordinary spaces are removed. NBSP/narrow NBSP/word joiners are retained.
// Two LFs denote a new paragraph; one LF is a hard line break, not an extra paragraph.
function normalize(value: string): string {
  return value
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => line.replace(/[ \t\f\v]+/g, " ").replace(/^ +| +$/g, ""))
    .join("\n");
}

// Normalize a whole styled paragraph, never individual runs: a style seam is
// not a word boundary. Keep the original face on every retained character.
function normalizeRuns(spans: Span[]): Span[] {
  const value = spanText(spans);
  const result: Span[] = [];
  const append = (span: Span) => {
    const previous = result[result.length - 1];
    if (
      previous &&
      (previous.properties === span.properties ||
        JSON.stringify(previous.properties) === JSON.stringify(span.properties))
    ) {
      previous.value += span.value;
    } else {
      result.push(span);
    }
  };
  for (const match of value.matchAll(/\r\n?|\n|[ \f\v]+|\t|[^ \f\v\t\r\n]+/g)) {
    const start = match.index!;
    const raw = match[0];
    if (/^[ \f\v]+$/.test(raw)) {
      const before = value[start - 1];
      const after = value[start + raw.length];
      if (!before || !after || /[\r\n\t]/.test(before + after)) continue;
      append({
        value: " ",
        properties: slices(spans, start, start + 1)[0].properties,
      });
    } else if (raw[0] === "\r") {
      append({
        value: "\n",
        properties: slices(spans, start, start + 1)[0].properties,
      });
    } else {
      for (const span of slices(spans, start, start + raw.length)) append(span);
    }
  }
  return result;
}
function placeholder(token: string): RegExp {
  return new RegExp(
    `\\(\\s*${token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*\\)`,
    "gi",
  );
}

export async function createPptxTextLayout(
  zip: JSZip,
  options: { fontProfile?: "english" | "indonesian" } = {},
): Promise<PptxTextLayout> {
  const profile = options.fontProfile ?? "english";
  const dirs = fontDirectories(profile);
  const registeredFonts = new Map<string, FontFace>();
  const documents = new Map<string, Xml[]>();
  await Promise.all(
    Object.values(zip.files)
      .filter((file) => /\.xml(?:\.rels)?$|\.rels$/.test(file.name))
      .map(async (file) => {
        documents.set(file.name, parse(await file.async("string")));
      }),
  );
  const ctx = new Canvas(1, 1).getContext("2d");
  const presentation = documents.get("ppt/presentation.xml") ?? [];
  function related(file: string, type: string): string | undefined {
    const rels =
      documents.get(
        path.posix.join(
          path.posix.dirname(file),
          "_rels",
          `${path.posix.basename(file)}.rels`,
        ),
      ) ?? [];
    const rel = descendants(rels, "Relationship").find((entry) =>
      attr(entry, "Type")?.endsWith(`/${type}`),
    );
    const target = attr(rel, "Target");
    return target
      ? path.posix.normalize(path.posix.join(path.posix.dirname(file), target))
      : undefined;
  }
  function inherited(file: string, shape: Xml) {
    const layoutFile = related(file, "slideLayout");
    const masterFile = layoutFile
      ? related(layoutFile, "slideMaster")
      : undefined;
    const layout = layoutFile ? (documents.get(layoutFile) ?? []) : [];
    const master = masterFile ? (documents.get(masterFile) ?? []) : [];
    const ph = descendants([shape], "p:ph")[0];
    function matching(nodes: Xml[], masterMatch = false): Xml | undefined {
      if (!ph) return undefined;
      return descendants(nodes, "p:sp").find((candidate) => {
        const candidatePh = descendants([candidate], "p:ph")[0];
        if (!candidatePh) return false;
        if (
          attr(ph, "idx") !== undefined &&
          attr(ph, "idx") === attr(candidatePh, "idx")
        )
          return true;
        const type = attr(ph, "type") ?? "body";
        return (
          (masterMatch || attr(ph, "idx") === undefined) &&
          (attr(candidatePh, "type") ?? "body") === type
        );
      });
    }
    const layoutShape = matching(layout),
      masterShape = matching(master, true);
    const themeFile = masterFile ? related(masterFile, "theme") : undefined;
    const theme = themeFile ? (documents.get(themeFile) ?? []) : [];
    return {
      layoutShape,
      masterShape,
      master,
      theme,
      type: attr(ph, "type") ?? "body",
    };
  }
  function geometry(file: string, shape: Xml): Geometry {
    const inherit = inherited(file, shape);
    const chain = [inherit.masterShape, inherit.layoutShape, shape];
    let ext: Xml | undefined;
    let transform: Xml | undefined;
    for (const candidate of chain) {
      const xfrm = child(child(candidate, "p:spPr"), "a:xfrm");
      if (child(xfrm, "a:ext")) {
        ext = child(xfrm, "a:ext");
        transform = xfrm;
      }
    }
    const body = merge(
      "a:bodyPr",
      chain.map((candidate) => child(child(candidate, "p:txBody"), "a:bodyPr")),
    );
    if (
      Number(attr(transform, "rot") ?? 0) % 21600000 !== 0 ||
      ![undefined, "horz"].includes(attr(body, "vert")) ||
      Number(attr(body, "numCol") ?? 1) !== 1
    )
      throw new Error(
        `${file}: English Mass layout requires an unrotated, horizontal, single-column text shape.`,
      );
    const width =
      Number(attr(ext, "cx")) / EMU -
      Number(attr(body, "lIns") ?? 91440) / EMU -
      Number(attr(body, "rIns") ?? 91440) / EMU;
    const height =
      Number(attr(ext, "cy")) / EMU -
      Number(attr(body, "tIns") ?? 45720) / EMU -
      Number(attr(body, "bIns") ?? 45720) / EMU;
    if (!(width > 0 && height > 0))
      throw new Error(
        `${file}: English Mass text shape is missing usable inherited extents/insets.`,
      );
    return { width, height, body };
  }
  function paragraph(file: string, shape: Xml, source: Xml): Paragraph {
    const inherit = inherited(file, shape);
    const own = child(source, "a:pPr"),
      level = Number(attr(own, "lvl") ?? 0);
    const levelTag = `a:lvl${level + 1}pPr`;
    const type = ["title", "ctrTitle"].includes(inherit.type)
      ? "p:titleStyle"
      : inherit.type === "body"
        ? "p:bodyStyle"
        : "p:otherStyle";
    const defaultStyle = descendants(presentation, "p:defaultTextStyle")[0];
    const masterStyle = descendants(inherit.master, type)[0];
    const pStyles = [
      child(defaultStyle, "a:defPPr"),
      child(defaultStyle, levelTag),
      child(masterStyle, levelTag),
      ...[inherit.masterShape, inherit.layoutShape, shape].map((candidate) =>
        child(child(child(candidate, "p:txBody"), "a:lstStyle"), levelTag),
      ),
      own,
    ];
    const properties = merge("a:pPr", pStyles);
    const defaultRun = merge(
      "a:rPr",
      pStyles.map((style) => child(style, "a:defRPr")),
    );
    function runProperties(run?: Xml): Xml {
      const result = merge("a:rPr", [defaultRun, run]);
      const latin = child(result, "a:latin");
      const family = attr(latin, "typeface");
      if (family?.startsWith("+")) {
        const scheme = descendants(
          inherit.theme,
          family.startsWith("+mj") ? "a:majorFont" : "a:minorFont",
        )[0];
        const resolved = attr(child(scheme, "a:latin"), "typeface");
        if (resolved)
          children(result).splice(
            children(result).findIndex((node) => tag(node) === "a:latin"),
            1,
            element("a:latin", [], { typeface: resolved }),
          );
      }
      if (
        !(
          attr(child(result, "a:sym"), "typeface") ||
          attr(child(result, "a:latin"), "typeface")
        ) ||
        !attr(result, "sz")
      )
        throw new Error(
          `${file}: a dynamic paragraph has no font family/size in its template style hierarchy.`,
        );
      return result;
    }
    const spans: Span[] = [];
    for (const node of children(source)) {
      if (tag(node) === "a:r" || tag(node) === "a:fld")
        spans.push({
          value: text(child(node, "a:t")),
          properties: runProperties(child(node, "a:rPr")),
        });
      else if (tag(node) === "a:br") {
        const preceding = spans[spans.length - 1]?.properties;
        spans.push({
          value: "\n",
          properties: runProperties(
            merge("a:rPr", [preceding, child(node, "a:rPr")]),
          ),
        });
      }
    }
    if (!spans.length)
      spans.push({
        value: "",
        properties: runProperties(child(source, "a:endParaRPr")),
      });
    return {
      properties,
      spans,
      end: child(source, "a:endParaRPr"),
      policy: "flow",
    };
  }
  function style(properties: Xml) {
    const family = (attr(child(properties, "a:sym"), "typeface") ||
      attr(child(properties, "a:latin"), "typeface"))!;
    const font = exactFont(
      family,
      attr(properties, "b") === "1" || attr(properties, "b") === "true",
      attr(properties, "i") === "1" || attr(properties, "i") === "true",
      profile,
      dirs,
      registeredFonts,
    );
    ctx.font = `${font.style} ${font.weight} ${Number(attr(properties, "sz")) / 100}px "${font.alias}"`;
    ctx.letterSpacing = `${Number(attr(properties, "spc") ?? 0) / 100}px`;
    ctx.fontKerning = "normal";
  }
  function measure(
    spans: Span[],
    properties: Xml,
  ): { width: number; natural: number; ink: number } {
    let width = 0,
      left = 0,
      right = 0,
      natural = 0,
      ink = 0;
    const tabSize = Number(attr(properties, "defTabSz") ?? 457200) / EMU;
    const tabStops = children(child(properties, "a:tabLst"))
      .map((node) => Number(attr(node, "pos")) / EMU)
      .filter(Number.isFinite)
      .sort((a, b) => a - b);
    // Merge matching adjacent run styles before shaping (including split-run placeholders).
    const groups: Span[] = [];
    for (const span of spans) {
      const previous = groups[groups.length - 1];
      if (
        previous &&
        JSON.stringify(previous.properties) === JSON.stringify(span.properties)
      )
        previous.value += span.value;
      else groups.push({ ...span });
    }
    for (const span of groups) {
      style(span.properties);
      const parts = span.value.split(/([\t\n])/);
      parts.forEach((part) => {
        if (part === "\t") {
          width =
            tabStops.find((stop) => stop > width) ??
            (Math.floor(width / tabSize) + 1) * tabSize;
          return;
        }
        if (part === "\n") {
          width = 0;
          return;
        }
        const metrics = ctx.measureText(part || " ");
        if (part) {
          left = Math.min(left, width - metrics.actualBoundingBoxLeft);
          right = Math.max(right, width + metrics.actualBoundingBoxRight);
          width += metrics.width;
        }
        natural = Math.max(
          natural,
          metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent,
        );
        ink = Math.max(
          ink,
          metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent,
        );
      });
    }
    return { width: Math.max(width, right) - left, natural, ink };
  }
  function spacing(node: Xml | undefined, natural: number): number {
    const points = child(node, "a:spcPts");
    return points
      ? Number(attr(points, "val")) / 100
      : Math.round(
          (Number(attr(child(node, "a:spcPct"), "val") ?? 0) / 100000) *
            natural *
            100,
        ) / 100;
  }
  function lines(paragraph: Paragraph, box: Geometry): Line[] {
    const value = spanText(paragraph.spans);
    const metric = measure(paragraph.spans, paragraph.properties);
    const lineSpacing = child(paragraph.properties, "a:lnSpc");
    const height =
      Math.ceil(
        Math.max(
          metric.ink,
          lineSpacing ? spacing(lineSpacing, metric.natural) : metric.natural,
        ) * 100,
      ) / 100;
    const before = spacing(
      child(paragraph.properties, "a:spcBef"),
      metric.natural,
    );
    const after = spacing(
      child(paragraph.properties, "a:spcAft"),
      metric.natural,
    );
    const margin =
      (Number(attr(paragraph.properties, "marL") ?? 0) +
        Number(attr(paragraph.properties, "marR") ?? 0)) /
      EMU;
    const indent = Number(attr(paragraph.properties, "indent") ?? 0) / EMU;
    const result: Line[] = [];
    let position = 0;
    const graphemes = new Intl.Segmenter("en", { granularity: "grapheme" });
    do {
      const first = result.length === 0;
      const available = box.width - margin - (first ? Math.max(0, indent) : 0);
      const hardEnd = value.indexOf("\n", position);
      const end = hardEnd === -1 ? value.length : hardEnd;
      let fit = position;
      // Candidate widths are shaped whole lines, never sums of character widths.
      const boundaries = [...graphemes.segment(value.slice(position, end))].map(
        (entry) => position + entry.index + entry.segment.length,
      );
      let low = 0,
        high = boundaries.length - 1;
      while (low <= high) {
        const middle = (low + high) >>> 1;
        if (
          measure(
            slices(paragraph.spans, position, boundaries[middle]),
            paragraph.properties,
          ).width <=
          available + 0.001
        ) {
          fit = boundaries[middle];
          low = middle + 1;
        } else high = middle - 1;
      }
      if (fit < end) {
        const candidate = value.slice(position, fit);
        const breakAt = Math.max(
          candidate.lastIndexOf(" "),
          candidate.lastIndexOf("\t"),
        );
        if (breakAt > 0) fit = position + breakAt + 1;
        else {
          // An oversized ordinary word may break at grapheme boundaries, but an
          // explicit nonbreaking group is indivisible and must fail visibly.
          const wordEnd = value.slice(position, end).search(/[ \t]/);
          const word = value.slice(
            position,
            wordEnd < 0 ? end : position + wordEnd,
          );
          if (/[\u00a0\u202f\u2060\ufeff]/u.test(word))
            throw new Error(
              "English Mass layout: a nonbreaking text group exceeds the shape width. Add an intentional break or widen the template shape.",
            );
        }
      }
      if (fit === position && end > position)
        throw new Error(
          `English Mass layout: one grapheme cannot fit in the ${available.toFixed(2)}pt text width at the template font size. Widen the shape; font shrinking is disabled.`,
        );
      const spans = slices(paragraph.spans, position, fit);
      if (!spans.length)
        spans.push({ value: "", properties: paragraph.spans[0].properties });
      result.push({
        spans,
        height,
        before: first ? before : 0,
        after: 0,
        paragraph,
        start: first,
        end: false,
        from: position,
      });
      position = fit;
      if (fit === end && hardEnd !== -1) position++;
      // A trailing hard break has a real final empty line.
      if (position === value.length && value.endsWith("\n")) {
        result.push({
          spans: [
            {
              value: "",
              properties:
                paragraph.spans[paragraph.spans.length - 1].properties,
            },
          ],
          height,
          before: 0,
          after: 0,
          paragraph,
          start: false,
          end: false,
          from: position,
        });
        break;
      }
    } while (position < value.length);
    result[result.length - 1].after = after;
    result[result.length - 1].end = true;
    return result;
  }
  const cost = (line: Line) => line.before + line.height + line.after;
  function render(shape: Shape, page: Line[], box: Geometry): string {
    const node = clone(shape.node),
      tx = child(node, "p:txBody")!;
    const body = clone(box.body);
    body["a:bodyPr"] = children(body).filter(
      (entry) =>
        !["a:spAutoFit", "a:normAutofit", "a:noAutofit"].includes(tag(entry)),
    );
    children(body).push(element("a:noAutofit"));
    // Resolved values make inherited continuation geometry and spacing explicit.
    const output: Xml[] = [];
    let current: Xml | undefined;
    let previous: Line | undefined;
    for (const line of page) {
      if (!current || line.paragraph !== previous?.paragraph || line.start) {
        const pPr = clone(line.paragraph.properties);
        pPr["a:pPr"] = children(pPr).filter(
          (entry) => !["a:lnSpc", "a:spcBef", "a:spcAft"].includes(tag(entry)),
        );
        children(pPr).unshift(
          element("a:lnSpc", [
            element("a:spcPts", [], {
              val: String(Math.round(line.height * 100)),
            }),
          ]),
          element("a:spcBef", [
            element("a:spcPts", [], {
              val: String(Math.round(line.before * 100)),
            }),
          ]),
          element("a:spcAft", [element("a:spcPts", [], { val: "0" })]),
        );
        pPr["a:pPr"] = children(pPr).sort(
          (a, b) =>
            paragraphPropertyOrder.indexOf(tag(a)) -
            paragraphPropertyOrder.indexOf(tag(b)),
        );
        current = element("a:p", [pPr]);
        output.push(current);
      } else
        children(current).push(
          element("a:br", [clone(line.spans[0].properties)]),
        );
      for (const span of line.spans)
        children(current).push(
          element("a:r", [
            clone(span.properties),
            element("a:t", [{ "#text": span.value }], {
              "xml:space": "preserve",
            }),
          ]),
        );
      const pPr = child(current, "a:pPr")!;
      const aft = child(child(pPr, "a:spcAft"), "a:spcPts")!;
      attrs(aft)["@_val"] = String(Math.round(line.after * 100));
      if (line.end && line.paragraph.end)
        children(current).push(clone(line.paragraph.end));
      previous = line;
    }
    if (!output.length) output.push(element("a:p"));
    tx["p:txBody"] = [
      body,
      clone(child(tx, "a:lstStyle") ?? element("a:lstStyle")),
      ...output,
    ];
    return builder.build([node]) as string;
  }
  return {
    paginate(slideFileName, slideXml, token, replacements, options = {}) {
      const file = slideFileName.startsWith("ppt/")
        ? slideFileName
        : `ppt/slides/${slideFileName}`;
      for (const xml of [
        slideXml,
        ...(options.continuationXml ? [options.continuationXml] : []),
      ]) {
        const groups = descendants(parse(xml), "p:grpSp");
        if (
          groups.some((group) =>
            descendants([group], "a:p").some((p) =>
              placeholder(token).test(text(p)),
            ),
          )
        ) {
          throw new Error(
            `${file}: a dynamic text shape is grouped. Ungroup it in the English template so its actual slide-space extents can be measured; grouped transforms are not supported.`,
          );
        }
      }
      const sourceShapes = shapes(slideXml),
        nextShapes = shapes(options.continuationXml ?? slideXml);
      let replacementIndex = 0;
      const touched = sourceShapes.filter((shape) =>
        children(child(shape.node, "p:txBody")).some(
          (p) => tag(p) === "a:p" && placeholder(token).test(text(p)),
        ),
      );
      if (!touched.length)
        throw new Error(
          `${file}: placeholder (${token}) was not found in a text paragraph.`,
        );
      const layouts = touched.map((shape, touchedIndex) => {
        const tx = child(shape.node, "p:txBody")!;
        const flow: Paragraph[] = [];
        const templateSpacers = new Set<Paragraph>();
        for (const source of children(tx).filter(
          (entry) => tag(entry) === "a:p",
        )) {
          const sourceText = text(source),
            matches = [...sourceText.matchAll(placeholder(token))];
          const paragraphModel = paragraph(file, shape.node, source);
          const firstOnly =
            options.firstOnlyTexts?.includes(sourceText) ?? false;
          const lastOnly = options.lastOnlyTexts?.includes(sourceText) ?? false;
          if (!matches.length) {
            if (!sourceText.trim()) {
              if (options.preserveTemplateSpacers) {
                templateSpacers.add(paragraphModel);
                flow.push(paragraphModel);
              }
              continue;
            }
            paragraphModel.policy = firstOnly
              ? "first"
              : lastOnly
                ? "last"
                : "flow";
            flow.push(paragraphModel);
            continue;
          }
          if (
            matches.some(
              (_, index) =>
                typeof replacements[replacementIndex + index] === "object",
            )
          ) {
            let current: Paragraph = { ...paragraphModel, spans: [] };
            let offset = 0;
            const flush = () => {
              if (current.spans.length) flow.push(current);
              current = { ...paragraphModel, spans: [] };
            };
            for (const match of matches) {
              const start = match.index!;
              current.spans.push(
                ...slices(paragraphModel.spans, offset, start),
              );
              const base =
                slices(paragraphModel.spans, start, start + 1)[0]?.properties ??
                paragraphModel.spans[0].properties;
              const replacement = replacements[replacementIndex++] ?? "";
              if (typeof replacement === "string") {
                current.spans.push({
                  value: normalize(replacement),
                  properties: base,
                });
              } else {
                replacement.paragraphs.forEach((part, index) => {
                  if (index) flush();
                  current.policy = part.policy ?? replacement.policy ?? "flow";
                  current.dynamicLast = current.policy === "last";
                  const styled = (run: ReplacementRun): Span => {
                    const properties = clone(base);
                    const bold = run.bold ?? part.bold;
                    const italic = run.italic ?? part.italic;
                    if (bold !== undefined)
                      attrs(properties)["@_b"] = bold ? "1" : "0";
                    if (italic !== undefined)
                      attrs(properties)["@_i"] = italic ? "1" : "0";
                    return { value: run.text, properties };
                  };
                  // Keep legacy text normalization byte-for-byte; richer input
                  // normalizes across seams while retaining its style slices.
                  const spans = part.runs
                    ? normalizeRuns(part.runs.map(styled))
                    : [
                        styled({
                          text: part.text.split("\t").map(normalize).join("\t"),
                        }),
                      ];
                  const value = spanText(spans);
                  let position = 0;
                  value.split("\n\n").forEach((segment, partIndex) => {
                    if (partIndex) {
                      const policy = current.policy;
                      const dynamicLast = current.dynamicLast;
                      flush();
                      current.policy = policy;
                      current.dynamicLast = dynamicLast;
                    }
                    const segmentSpans = slices(
                      spans,
                      position,
                      position + segment.length,
                    );
                    current.spans.push(
                      ...(segmentSpans.length
                        ? segmentSpans
                        : [styled({ text: "" })]),
                    );
                    position += segment.length + 2;
                  });
                });
              }
              offset = start + match[0].length;
            }
            current.spans.push(
              ...slices(paragraphModel.spans, offset, sourceText.length),
            );
            flush();
            continue;
          }
          const replaced: Span[] = [];
          let offset = 0;
          for (const match of matches) {
            const start = match.index!;
            replaced.push(...slices(paragraphModel.spans, offset, start));
            const replacementStyle =
              slices(paragraphModel.spans, start, start + 1)[0]?.properties ??
              paragraphModel.spans[0].properties;
            replaced.push({
              value: normalize(
                (replacements[replacementIndex++] as string) ?? "",
              ),
              properties: replacementStyle,
            });
            offset = start + match[0].length;
          }
          replaced.push(
            ...slices(paragraphModel.spans, offset, sourceText.length),
          );
          // firstOnly on a token paragraph applies to its fixed prefix/suffix,
          // never to the replacement body. Keep that prefix in semantic flow once.
          const value = spanText(replaced);
          let position = 0;
          for (const part of value.split("\n\n")) {
            const spans = slices(replaced, position, position + part.length);
            if (!spans.length)
              spans.push({
                value: "",
                properties:
                  replaced[0]?.properties ?? paragraphModel.spans[0].properties,
              });
            flow.push({ ...paragraphModel, spans });
            position += part.length + 2;
          }
        }
        const nextShape =
          nextShapes.filter((candidate) =>
            children(child(candidate.node, "p:txBody")).some(
              (p) => tag(p) === "a:p" && placeholder(token).test(text(p)),
            ),
          )[touchedIndex] ??
          nextShapes[shape.index] ??
          nextShapes[touchedIndex];
        if (!nextShape)
          throw new Error(
            `${file}: continuation is missing a compatible text shape.`,
          );
        const firstBox = geometry(file, shape.node),
          nextBox = geometry(file, nextShape.node);
        if (templateSpacers.size) {
          let requiredHeight = flow.reduce(
            (sum, p) =>
              sum +
              lines(p, firstBox).reduce(
                (height, line) => height + cost(line),
                0,
              ),
            0,
          );
          while (
            requiredHeight > firstBox.height + 0.001 &&
            templateSpacers.size
          ) {
            // Reduce double/triple gaps before removing the only gap between
            // two sections. Only template blanks are disposable, never content.
            let index = flow.findIndex(
              (p, i) =>
                templateSpacers.has(p) &&
                (templateSpacers.has(flow[i - 1]) ||
                  templateSpacers.has(flow[i + 1])),
            );
            if (index < 0)
              index = flow.findIndex((p) => templateSpacers.has(p));
            const spacer = flow[index];
            requiredHeight -= lines(spacer, firstBox).reduce(
              (sum, line) => sum + cost(line),
              0,
            );
            templateSpacers.delete(spacer);
            flow.splice(index, 1);
          }
        }
        const dynamicFooters = flow.filter((p) => p.dynamicLast);
        const dynamicFooterHeight = dynamicFooters.reduce(
          (sum, p) =>
            sum +
            lines(p, nextBox).reduce((height, line) => height + cost(line), 0),
          0,
        );
        if (dynamicFooterHeight > nextBox.height + 0.001) {
          // An oversized response is ordered content, not an indivisible static
          // footer. Reuse normal paragraph pagination after the complete body.
          for (const p of dynamicFooters) {
            flow.splice(flow.indexOf(p), 1);
            p.policy = "flow";
          }
          flow.push(...dynamicFooters);
        }
        const pages: Line[][] = [];
        // Each shape has one shared height budget, including every placeholder,
        // fixed heading, response, paragraph gap and first/last-only paragraph.
        let pending = flow.filter((p) => p.policy !== "last");
        const last = flow.filter((p) => p.policy === "last");
        let pageIndex = 0;
        while (pending.length || !pages.length) {
          const box = pageIndex === 0 ? firstBox : nextBox;
          const page: Line[] = [];
          let used = 0;
          let consumed = 0;
          for (const p of pending) {
            if (p.policy === "first" && pageIndex > 0)
              throw new Error(
                `${file}: first-only static content cannot fit on the first slide. Shorten it or enlarge the template text shape.`,
              );
            const wrapped = lines(p, box);
            let count = 0;
            for (const line of wrapped) {
              if (used + cost(line) > box.height + 0.001) break;
              page.push(line);
              used += cost(line);
              count++;
            }
            if (count === wrapped.length) {
              consumed++;
              continue;
            }
            if (!count && !page.length)
              throw new Error(
                `${file}: one text line plus its paragraph spacing cannot fit in ${box.height.toFixed(2)}pt. Increase the template shape height; clipping and shrinking are disabled.`,
              );
            if (p.policy === "first")
              throw new Error(
                `${file}: first-only static content exceeds the first slide capacity.`,
              );
            const remaining = slices(
              p.spans,
              wrapped[count].from,
              spanText(p.spans).length,
            );
            if (!remaining.length)
              remaining.push({
                value: "",
                properties: p.spans[p.spans.length - 1].properties,
              });
            // Reflow original text, not earlier soft-wrap boundaries, when the
            // continuation shape has a different width.
            const properties = clone(p.properties);
            if (count) {
              properties["a:pPr"] = children(properties).filter(
                (entry) => tag(entry) !== "a:spcBef",
              );
              children(properties).push(
                element("a:spcBef", [element("a:spcPts", [], { val: "0" })]),
              );
              attrs(properties)["@_indent"] = "0";
            }
            pending = [
              { ...p, properties, spans: remaining },
              ...pending.slice(consumed + 1),
            ];
            consumed = 0;
            break;
          }
          if (consumed) pending = pending.slice(consumed);
          const footers = last.flatMap((p) => lines(p, box));
          const footerCost = footers.reduce((sum, line) => sum + cost(line), 0);
          if (
            footerCost > box.height + 0.001 &&
            last.some((p) => !p.dynamicLast)
          )
            throw new Error(
              `${file}: last-only static text exceeds the entire shape capacity.`,
            );
          if (!pending.length && used + footerCost <= box.height + 0.001)
            page.push(...footers);
          else if (!pending.length && footers.length) {
            const nextFooters = last.flatMap((p) => lines(p, nextBox));
            if (
              nextFooters.reduce((sum, line) => sum + cost(line), 0) >
              nextBox.height + 0.001
            ) {
              throw new Error(
                `${file}: last-only static text exceeds the continuation shape capacity.`,
              );
            }
            // Move the original tail back through normal pagination: its line
            // may wrap differently in the continuation and need several pages.
            const tail = page[page.length - 1];
            if (page.length > 1 && tail.paragraph.policy !== "first") {
              page.pop();
              const properties = clone(tail.paragraph.properties);
              if (!tail.start) {
                properties["a:pPr"] = children(properties).filter(
                  (entry) => tag(entry) !== "a:spcBef",
                );
                children(properties).push(
                  element("a:spcBef", [element("a:spcPts", [], { val: "0" })]),
                );
                attrs(properties)["@_indent"] = "0";
              }
              pending = [
                {
                  ...tail.paragraph,
                  properties,
                  spans: slices(
                    tail.paragraph.spans,
                    tail.from,
                    spanText(tail.paragraph.spans).length,
                  ),
                },
              ];
            } else {
              // If even one body line cannot share a page with the complete
              // response, a response-only final slide is unavoidable.
              if (page.length) pages.push(page);
              pages.push(nextFooters);
              break;
            }
          }
          pages.push(page);
          pageIndex++;
        }
        for (let index = 0; index < pages.length; index++) {
          const box = index ? nextBox : firstBox;
          if (
            pages[index].reduce((sum, line) => sum + cost(line), 0) >
            box.height + 0.001
          )
            throw new Error(
              `${file}: emitted text exceeds the measured shape capacity.`,
            );
        }
        return { shape, nextShape, firstBox, nextBox, pages };
      });
      if (replacementIndex < replacements.length)
        throw new Error(
          `${file}: ${replacements.length} replacements were supplied for ${replacementIndex} matching placeholder occurrences.`,
        );
      const count = Math.max(
        1,
        ...layouts.map((layout) => layout.pages.length),
      );
      return Array.from({ length: count }, (_, index) => {
        let xml = index ? (options.continuationXml ?? slideXml) : slideXml;
        const replacementsByShape = new Map<number, string>();
        for (const layout of layouts) {
          const shape = index ? layout.nextShape : layout.shape;
          if (replacementsByShape.has(shape.index))
            throw new Error(
              `${file}: multiple source shapes map to one continuation shape.`,
            );
          replacementsByShape.set(
            shape.index,
            render(
              shape,
              layout.pages[index] ?? [],
              index ? layout.nextBox : layout.firstBox,
            ),
          );
        }
        let shapeIndex = 0;
        xml = xml.replace(
          /<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g,
          (raw) => replacementsByShape.get(shapeIndex++) ?? raw,
        );
        return xml;
      });
    },
  };
}
