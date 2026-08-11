import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

export interface LiturgyRun {
  text: string;
  bold: boolean;
  underline: boolean;
  italic: boolean;
}

export interface LiturgyParagraph {
  rawText: string;
  runs: LiturgyRun[];
  isBold: boolean;
  isUnderline: boolean;
  isItalic: boolean;
}

export interface ParsedLiturgySection {
  id: string;
  title: string;
  type:
    | "major_header"
    | "sub_header"
    | "reading"
    | "psalm"
    | "prayer"
    | "doa_umat"
    | "generic";
  epigraph?: string;
  citation?: string;
  body?: string[];
  notesBody?: string;
  intentions?: { title: string; body: string; response?: string }[];
  paragraphs: LiturgyParagraph[];
}

export interface DocxParseResult {
  fileName: string;
  totalParagraphs: number;
  sections: ParsedLiturgySection[];
  rawParagraphs: LiturgyParagraph[];
}

export interface MappedReading {
  epigraph?: string;
  citation?: string;
  notesBody?: string;
}

export interface MappedGospel {
  gospelName?: string;
  epigraph?: string;
  citation?: string;
  notesBody?: string;
}

export interface MappedDoaUmat {
  intro?: string;
  intentions: { title: string; body: string; response?: string }[];
  closing?: string;
}

export interface MappedLiturgyData {
  seruanTobat?: string[];
  doaKolekta?: string;
  bacaanPertama?: MappedReading;
  mazmurTanggapan?: { verses: string[] };
  bacaanKedua?: MappedReading;
  bacaanInjil?: MappedGospel;
  doaUmat?: MappedDoaUmat;
  doaAtasPersembahan?: string;
  doaSesudahKomuni?: string;
  unmatchedSections: string[];
}

// ponytail: helper to ensure fast-xml-parser outputs are always arrays
function toArray<T>(val: T | T[] | undefined | null): T[] {
  if (!val) return [];
  return Array.isArray(val) ? val : [val];
}

// ponytail: extract heading title text from runs that are bold AND NOT italic
function getHeadingTitleText(p: LiturgyParagraph): string | null {
  const titleRuns = p.runs.filter((r) => r.bold && !r.italic);
  if (titleRuns.length === 0) return null;
  const titleText = titleRuns
    .map((r) => r.text)
    .join("")
    .trim();
  return titleText.length > 0 ? titleText : null;
}

// ponytail: helper to clean formula lines, posture titles, and strip speaker prefixes from prayer bodies
function cleanPrayerBody(lines: string[], sectionTitle?: string): string {
  // Find index of congregation response (U. Amin / Amin)
  let aminIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (/^U[\s.:]*Amin/i.test(trimmed) || /^Amin\.?$/i.test(trimmed)) {
      aminIdx = i;
      break;
    }
  }

  // If U. Amin was found, ignore all lines after U. Amin
  const targetLines = aminIdx !== -1 ? lines.slice(0, aminIdx) : lines;

  const filtered = targetLines
    .filter((line) => {
      const trimmed = line.trim();
      if (!trimmed) return false;

      // Filter out section title if present in body
      if (sectionTitle && trimmed.toLowerCase() === sectionTitle.toLowerCase())
        return false;
      if (
        /^(?:DOA\s+)?(KOLEKTA|ATAS\s+PERSEMBAHAN|SESUDAH\s+KOMUNI)\s*(\(.*\))?$/i.test(
          trimmed,
        )
      )
        return false;

      // Filter out posture annotations alone e.g. "(umat berdiri)"
      if (/^\(?umat\s*(berdiri|duduk|berlutut)\)?$/i.test(trimmed))
        return false;

      // Filter out intro prompt lines: Marilah kita berdoa / Mari kita berdoa / Mari kita berdoa bersama / hening sejenak
      if (
        /^(?:[LI][.:\s\t]*|(?:Imam|Lektor)[.:\s\t]*)?Mari(?:lah)?\b.*berdoa/i.test(
          trimmed,
        )
      )
        return false;
      if (/^\(?hening\s+sejenak\)?$/i.test(trimmed)) return false;

      // Filter out U. Amin / Amin
      if (/^U[\s.:]*Amin/i.test(trimmed)) return false;
      if (/^Amin\.?$/i.test(trimmed)) return false;
      return true;
    })
    .map((line) => {
      let cleaned = line.trim();
      cleaned = cleaned.replace(/^(?:I|Imam|L|Lektor)[.:\s\t]+/i, "");
      cleaned = cleaned.replace(/^(?:I|Imam|L|Lektor)[.:]/i, "");
      return cleaned.trim();
    });

  return filtered.join("\n\n").trim();
}

// ponytail: helper to check if a paragraph is a liturgical dialogue greeting line (e.g. "I. Tuhan bersamamu", "U. Dan bersama roh-Mu", "U. Dimuliakanlah Tuhan")
function isGospelDialogueLine(text: string): boolean {
  const norm = text.trim();
  if (/tuhan\s+bersamamu|tuhan\s+serta-mu/i.test(norm)) return true;
  if (/dan\s+bersama\s+roh-?mu/i.test(norm)) return true;
  if (/dimuliakanlah\s+tuhan/i.test(norm)) return true;
  return false;
}

// ponytail: filter out Mazmur section headers, musical notations (Do=A; 4/4), refren titles, and posture annotations ((umat duduk))
function isMazmurHeaderOrMetadata(text: string): boolean {
  const norm = text.trim();
  if (!norm) return true;
  if (/^mazmur\s+tanggapan/i.test(norm)) return true;
  if (/^bacaan\s*mazmur/i.test(norm)) return true;
  if (/^\(?umat\s*(duduk|berdiri|berlutut)\)?$/i.test(norm)) return true;
  if (
    /\b(Do|Re|Mi|Fa|Sol|La|Si|[1-7])\s*=\s*[A-G][#b]?/i.test(norm) ||
    /\b\d\/\d\b/.test(norm)
  )
    return true;
  if (/^(refren|ulangan|ref)[.:]?/i.test(norm)) return true;
  return false;
}

// ponytail: split Mazmur paragraphs into distinct verses directly by paragraph breaks (Enters)
function splitMazmurVerses(cleanParagraphs: LiturgyParagraph[]): string[] {
  return cleanParagraphs
    .map((p) => p.rawText.trim())
    .filter((text) => !isMazmurHeaderOrMetadata(text));
}

// ponytail: normalize all speaker prefix combinations (e.g. "L", "L.", "L. ", "L.\t", "L\t", "L:", "Lektor:", "1.", "1.\t", fused "LBagi")
function normalizeSpeakerPrefix(text: string): string {
  if (!text) return "";
  let norm = text.trim();
  norm = norm.replace(
    /^(Lektor|Imam|Umat)[.:\s\t]*/i,
    (match) => `${match.charAt(0).toUpperCase()}. `,
  );
  norm = norm.replace(/^([ILU])([A-Z\u00C0-\u024F])/, "$1. $2");
  norm = norm.replace(/^([ILU])\.([A-Z\u00C0-\u024F])/i, "$1. $2");
  norm = norm.replace(/^([ILU])[.:\s\t]+/i, "$1. ");
  return norm.trim();
}

// ponytail: helper to strip speaker prefixes (I., L., U., I , L , U , Imam:, Lektor:, Umat:, 1., 1.\t, 1. )
function stripSpeakerPrefix(text: string): string {
  if (!text) return "";
  let norm = normalizeSpeakerPrefix(text);
  norm = norm.replace(/^(Imam|Lektor|Umat|[ILU])[.:\s\t]*/i, "").trim();
  norm = norm.replace(/^\d+[.:\s\t]*/, "").trim();
  return norm;
}

// ponytail: check if a paragraph is a posture annotation like (umat berdiri)
function isPostureAnnotation(text: string): boolean {
  const norm = text.trim();
  if (/^\(?umat\s*(duduk|berdiri|berlutut)\)?$/i.test(norm)) return true;
  if (/^doa\s+umat\s*\(?umat\s*(duduk|berdiri|berlutut)\)?$/i.test(norm))
    return true;
  return false;
}

// ponytail: check if a text is a section title for Doa Umat (e.g. "Doa Umat", "Doa Umat (umat berdiri)")
function isDoaUmatTitle(text: string): boolean {
  const norm = text.trim();
  if (/^doa\s+umat/i.test(norm)) return true;
  return false;
}

// ponytail: check if a line is a Doa Ujub / Doa Pribadi marker
function isDoaUjubMarker(text: string): boolean {
  return /doa\s*(ujub|pribadi)|hening\s*sejenak/i.test(text.trim());
}

// ponytail: check if a line is a congregation response line matching U or Umat speaker prefix ONLY (zero hardcoded response words)
function isCongregationResponse(text: string): boolean {
  const norm = normalizeSpeakerPrefix(text);
  if (/^U[.:\s\t]/i.test(norm) || /^Umat/i.test(norm)) return true;
  return false;
}

// ponytail: check if a heading title is a known liturgy section header that can close Doa Umat
function isKnownLiturgySectionHeader(title: string): boolean {
  const upper = title.toUpperCase();
  return (
    upper.includes("LITURGI") ||
    upper.includes("PERSEMBAHAN") ||
    upper.includes("BACAAN") ||
    upper.includes("INJIL") ||
    upper.includes("MAZMUR") ||
    upper.includes("KOLEKTA") ||
    upper.includes("TOBAT") ||
    upper.includes("BAPA KAMI") ||
    upper.includes("PENGUMUMAN")
  );
}

// ponytail: helper to parse Doa Umat intro, intentions, and true priest conclusion outro
function parseDoaUmatData(section: ParsedLiturgySection): MappedDoaUmat {
  const validParagraphs = section.paragraphs.filter((p) => {
    const text = p.rawText.trim();
    if (text === section.title.trim()) return false;
    if (isDoaUmatTitle(text)) return false;
    if (isPostureAnnotation(text)) return false;
    return true;
  });

  if (validParagraphs.length === 0) {
    return { intro: "", intentions: [], closing: "" };
  }

  // 1. Intro: First text starting with "I" / "Imam" or validParagraphs[0]
  const introP =
    validParagraphs.find((p) => {
      const norm = normalizeSpeakerPrefix(p.rawText);
      return /^I[.:\s\t]/i.test(norm) || /^Imam/i.test(norm);
    }) || validParagraphs[0];
  const intro = introP ? stripSpeakerPrefix(introP.rawText) : "";

  // 2. Intentions & Outro: Process all valid paragraphs after introP
  const intentions: { title: string; body: string; response?: string }[] = [];
  const outroLines: string[] = [];
  let foundUjub = false;

  for (const p of validParagraphs) {
    if (p === introP) continue;
    const text = p.rawText.trim();
    if (!text) continue;

    if (isPostureAnnotation(text)) continue;

    // Check if Doa Ujub / Pribadi marker reached
    if (isDoaUjubMarker(text)) {
      foundUjub = true;
      continue;
    }

    if (foundUjub) {
      const norm = normalizeSpeakerPrefix(text);
      const isPriestLine = /^I[.:\s\t]/i.test(norm) || /^Imam/i.test(norm);

      if (isPriestLine) {
        const cleanOutro = stripSpeakerPrefix(text);
        // Exclude formula prompt lines like "Kami mohon..."
        if (cleanOutro && !/^kami\s+mohon/i.test(cleanOutro)) {
          outroLines.push(cleanOutro);
          break; // Captured the priest conclusion prayer!
        }
      }
      continue;
    }

    // Skip congregation response lines (matching U / Umat speaker prefix)
    if (isCongregationResponse(text)) {
      const cleanResponse = stripSpeakerPrefix(text);
      if (intentions.length > 0) {
        intentions[intentions.length - 1].response = cleanResponse;
      }
      continue;
    }

    const cleanText = stripSpeakerPrefix(text);
    if (!cleanText) continue;

    const firstDotIdx = cleanText.indexOf(".");
    if (firstDotIdx > 0) {
      intentions.push({
        title: cleanText.slice(0, firstDotIdx + 1).trim(),
        body: cleanText.slice(firstDotIdx + 1).trim(),
        response: "",
      });
    } else {
      intentions.push({ title: cleanText, body: cleanText, response: "" });
    }
  }

  return {
    intro,
    intentions,
    closing: outroLines.join("\n\n").trim(),
  };
}

// ponytail: helper to strip trailing invocation ending phrases like "Tuhan, kasihanilah kami" or "Kristus, kasihanilah kami"
function stripTobatInvocationEnding(text: string): string {
  if (!text) return "";
  let cleaned = text.trim();
  cleaned = cleaned
    .replace(/[\s,.-]*(?:Tuhan|Kristus)[\s,.-]*kasihanilah\s*kami[\s.]*$/gi, "")
    .trim();
  return cleaned;
}

// ponytail: helper to parse Seruan Tobat items (capturing at most 3 priest "I" prayers, stripping trailing kasihanilah kami)
function parseSeruanTobatLines(lines: string[]): string[] {
  const tobatItems: string[] = [];
  let currentItemLines: string[] = [];

  for (const rawLine of lines) {
    const trimmed = rawLine.trim();
    if (!trimmed) continue;

    // Filter out section title and posture annotations
    if (
      /^tobat|^seruan\s*tobat/i.test(trimmed) &&
      (/\(\s*umat/i.test(trimmed) || trimmed.split(/\s+/).length <= 3)
    )
      continue;
    if (/^\(\s*umat\s*(?:berdiri|duduk|berlutut)\s*\)$/i.test(trimmed))
      continue;

    // Stop at Priest Absolusi conclusion prayer (e.g. "Semoga Allah yang Mahakuasa...")
    if (
      /^semoga\s+allah/i.test(trimmed) ||
      /mengampuni\s+dosa/i.test(trimmed)
    ) {
      break;
    }

    // Check if line is Congregation response (U. ...)
    if (isCongregationResponse(trimmed)) {
      if (currentItemLines.length > 0) {
        const fullItemText = currentItemLines.join(" ").trim();
        const cleanedText = stripTobatInvocationEnding(fullItemText);
        if (cleanedText) {
          tobatItems.push(cleanedText);
        }
        currentItemLines = [];
      }
      if (tobatItems.length >= 3) break;
      continue;
    }

    // Check if line starts with Priest speaker prefix (I. / I / Imam)
    const isPriestPrefix = /^(?:[I][.:\s\t]*|Imam[.:\s\t]*)/i.test(trimmed);
    if (isPriestPrefix && currentItemLines.length > 0) {
      const fullItemText = currentItemLines.join(" ").trim();
      const cleanedText = stripTobatInvocationEnding(fullItemText);
      if (cleanedText) {
        tobatItems.push(cleanedText);
      }
      currentItemLines = [];
    }

    if (tobatItems.length >= 3) break;

    // Line spoken by Priest (I. ...) or continuation line before U
    const cleanLine = stripSpeakerPrefix(trimmed);
    if (cleanLine) {
      currentItemLines.push(cleanLine);
    }
  }

  // Trailing item if any
  if (currentItemLines.length > 0 && tobatItems.length < 3) {
    const fullItemText = currentItemLines.join(" ").trim();
    const cleanedText = stripTobatInvocationEnding(fullItemText);
    if (cleanedText) {
      tobatItems.push(cleanedText);
    }
  }

  return tobatItems.slice(0, 3);
}

// ponytail: helper to clean reading notes body text (excluding header, epigraph, citation, dialogue greetings, Demikianlah sabda Tuhan, U Syukur kepada Allah)
function cleanReadingNotesBody(lines: string[], epigraph?: string): string {
  const normEpigraph = epigraph
    ? epigraph
        .replace(/^[*"“']|["”'*]$/g, "")
        .trim()
        .toLowerCase()
    : "";

  const filtered = lines.filter((line) => {
    const trimmed = line.trim();
    if (!trimmed) return false;

    // Filter out section headers like "Bacaan I(umat duduk)", "Bacaan II", "Injil"
    if (/^bacaan\s*(I|II|pertama|kedua|injil)/i.test(trimmed)) return false;
    if (/^injil/i.test(trimmed)) return false;
    if (isPostureAnnotation(trimmed)) return false;

    // Filter out Gospel dialogue greetings
    if (isGospelDialogueLine(trimmed)) return false;

    // Filter out standalone epigraph lines (exact string match regardless of length)
    if (normEpigraph && trimmed.toLowerCase() === normEpigraph) {
      return false;
    }

    // Filter out closing responses
    if (/^demikianlah\s*(sabda|injil)/i.test(trimmed)) return false;
    if (/^U[\s.]*(syukur|terpujilah)/i.test(trimmed)) return false;
    if (/^(syukur kepada allah|terpujilah kristus)\.?$/i.test(trimmed))
      return false;

    return true;
  });
  return filtered.join("\n\n").trim();
}

// ponytail: helper to check if paragraph is a scripture reading citation line (defining source)
function isReadingCitationLine(
  p: LiturgyParagraph,
  headingTitle: string | null,
): boolean {
  const text = (headingTitle || p.rawText).trim();

  // If it's a greeting dialogue line (e.g. "I. Tuhan bersamamu"), it is NOT a citation line!
  if (isGospelDialogueLine(text)) return false;

  // Scripture source keywords (e.g. "Inilah Injil", "Bacaan dari", "Matius", "Markus", "Lukas", "Yohanes", "Kitab", "Surat")
  if (
    /bacaan dari|inilah injil|injil suci|kitab|surat|kisah|warta|nubuat|petikan/i.test(
      text,
    )
  )
    return true;
  if (
    /^[LI][.:\s\t]*(bacaan|inilah|kitab|surat|kisah|warta|nubuat|petikan)/i.test(
      text,
    )
  )
    return true;

  return false;
}

// ponytail: extract text and run formatting from word/document.xml
export async function parseDocxFile(file: File): Promise<DocxParseResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);

  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) {
    throw new Error(
      "Berkas .docx tidak valid: word/document.xml tidak ditemukan.",
    );
  }

  const xmlContent = await documentXmlFile.async("text");

  // ponytail: trimValues: false is critical to preserve spaces in text runs!
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    parseTagValue: false,
    trimValues: false,
  });

  const parsedXml = parser.parse(xmlContent);

  const body = parsedXml?.["w:document"]?.["w:body"];
  if (!body) {
    throw new Error("Struktur XML word/document.xml tidak dapat dibaca.");
  }

  const rawParagraphs: LiturgyParagraph[] = [];
  const pNodes = toArray(body["w:p"]);

  for (const pNode of pNodes) {
    const runs: LiturgyRun[] = [];
    const rNodes = toArray(pNode["w:r"]);

    for (const rNode of rNodes) {
      const rPr = rNode["w:rPr"];

      // Check formatting tags
      const hasBold = Boolean(
        rPr &&
        ("w:b" in rPr ||
          rPr["w:b"]?.["@_w:val"] === "true" ||
          rPr["w:b"]?.["@_w:val"] === "1" ||
          rPr["w:b"] === ""),
      );

      // Underline detection logic
      const uVal = rPr?.["w:u"]?.["@_w:val"];
      const hasUnderline = Boolean(rPr && "w:u" in rPr) && uVal !== "none";

      const hasItalic = Boolean(
        rPr &&
        ("w:i" in rPr ||
          rPr["w:i"]?.["@_w:val"] === "true" ||
          rPr["w:i"]?.["@_w:val"] === "1" ||
          rPr["w:i"] === ""),
      );

      // Extract text
      let text = "";
      if (rNode["w:t"]) {
        if (typeof rNode["w:t"] === "string") {
          text = rNode["w:t"];
        } else if (
          typeof rNode["w:t"] === "object" &&
          "#text" in rNode["w:t"]
        ) {
          text = rNode["w:t"]["#text"];
        }
      } else if (rNode["w:tab"]) {
        text = "\t";
      } else if (rNode["w:br"]) {
        text = "\n";
      }

      if (text) {
        runs.push({
          text,
          bold: hasBold,
          underline: hasUnderline,
          italic: hasItalic,
        });
      }
    }

    let rawText = runs.map((r) => r.text).join("");

    // ponytail: check if paragraph uses Word auto-numbering (w:numPr)
    const pPr = pNode["w:pPr"];
    const hasNumPr = Boolean(
      pPr && ("w:numPr" in pPr || "w:outlineLvl" in pPr),
    );

    // ponytail: strip Word auto-numbering prefixes if numPr present or if fused like "1.HARI"
    if (hasNumPr || /^\d+\.[A-Za-z\u00C0-\u024F]/.test(rawText.trim())) {
      rawText = rawText.trim().replace(/^\d+\.\s*/, "");
    } else {
      rawText = rawText.trim();
    }

    // ponytail: normalize fused speaker labels (e.g. "ITuhan" -> "I. Tuhan", "LBagi" -> "L. Bagi", "UKabulkanlah" -> "U. Kabulkanlah")
    rawText = normalizeSpeakerPrefix(rawText);
    rawText = rawText.replace(/^([LI])\.\s*([A-Za-z])/i, "$1. $2");
    // ponytail: fix fused words like "dariKitab" -> "dari Kitab"
    rawText = rawText.replace(/([a-z])Kitab/g, "$1 Kitab");

    if (rawText.length > 0) {
      const pStyle = pPr?.["w:pStyle"]?.["@_w:val"] || "";
      const isHeading = pStyle.toLowerCase().includes("heading");

      const isAllBold = runs.length > 0 && runs.every((r) => r.bold);
      const hasBold = runs.some((r) => r.bold) || isHeading;
      const hasUnderline = runs.some((r) => r.underline);
      const hasItalic = runs.some((r) => r.italic);

      rawParagraphs.push({
        rawText,
        runs,
        isBold: hasBold || isAllBold,
        isUnderline: hasUnderline,
        isItalic: hasItalic,
      });
    }
  }

  // ponytail: group raw paragraphs into structured sections
  const sections = groupIntoSections(rawParagraphs);

  return {
    fileName: file.name,
    totalParagraphs: rawParagraphs.length,
    sections,
    rawParagraphs,
  };
}

// ponytail: section classifier using run-level bold-and-not-italic title extraction
function groupIntoSections(
  paragraphs: LiturgyParagraph[],
): ParsedLiturgySection[] {
  const sections: ParsedLiturgySection[] = [];
  let currentSection: ParsedLiturgySection | null = null;

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];

    // 1. Major Section Header (Bold + Underline)
    if (p.isBold && p.isUnderline) {
      if (currentSection) sections.push(currentSection);
      const cleanMajorTitle = getHeadingTitleText(p) || p.rawText;
      currentSection = {
        id: `section-${sections.length + 1}`,
        title: cleanMajorTitle,
        type: "major_header",
        paragraphs: [p],
        body: [],
      };
      continue;
    }

    const headingTitle = getHeadingTitleText(p);

    // ponytail: check if headingTitle is a bare speaker label (e.g. "U", "L.", "I:")
    const normTitle = headingTitle
      ? normalizeSpeakerPrefix(headingTitle)
      : null;
    const isJustSpeakerLabel =
      normTitle !== null && /^[ULI][.:]?$/i.test(normTitle.trim());

    // ponytail: check if paragraph is a speaker line (I, L, U, Imam, Lektor, Umat)
    const isSpeakerLabelLine =
      normTitle !== null &&
      (/^[ILU][.:\s\t]/i.test(normTitle.trim()) ||
        /^(Imam|Lektor|Umat)[.:\s\t]/i.test(normTitle.trim()));

    // ponytail: check if current section is Doa Umat
    const isDoaUmatSection =
      currentSection !== null &&
      (currentSection.type === "doa_umat" ||
        currentSection.title.toUpperCase().includes("DOA UMAT"));

    // ponytail: check if paragraph is a scripture reading citation line (defining source)
    const isCitationLine = isReadingCitationLine(p, headingTitle);

    // 2. Sub-Section Header condition
    const shouldOpenNewSection =
      headingTitle !== null &&
      !isJustSpeakerLabel &&
      !isSpeakerLabelLine &&
      !(isDoaUmatSection && !isKnownLiturgySectionHeader(headingTitle)) &&
      !(currentSection?.type === "reading" && isCitationLine);

    if (shouldOpenNewSection) {
      if (currentSection) sections.push(currentSection);

      const titleUpper = headingTitle.toUpperCase();
      let type: ParsedLiturgySection["type"] = "sub_header";
      if (
        titleUpper.includes("BAIT PENGANTAR") ||
        titleUpper.includes("ALLELUIA")
      ) {
        type = "sub_header";
      } else if (
        titleUpper.includes("BACAAN") ||
        titleUpper.includes("INJIL")
      ) {
        type = "reading";
      } else if (titleUpper.includes("MAZMUR")) {
        type = "psalm";
      } else if (titleUpper.includes("DOA UMAT")) {
        type = "doa_umat";
      } else if (titleUpper.includes("DOA")) {
        type = "prayer";
      }

      currentSection = {
        id: `section-${sections.length + 1}`,
        title: headingTitle,
        type,
        paragraphs: [p],
        body: [],
      };
      continue;
    }

    // Append to current section or fallback generic section
    if (!currentSection) {
      currentSection = {
        id: "section-0",
        title: "Pendahuluan",
        type: "generic",
        paragraphs: [],
        body: [],
      };
    }

    currentSection.paragraphs.push(p);
    if (currentSection.body) {
      currentSection.body.push(p.rawText);
    }
  }

  if (currentSection) {
    sections.push(currentSection);
  }

  // Refine section details (epigraph, citation, intentions)
  return sections.map((sec) => processSectionDetails(sec));
}

// ponytail: extract reading epigraphs (multi-paragraph), citations, and Doa Umat intentions
function processSectionDetails(
  section: ParsedLiturgySection,
): ParsedLiturgySection {
  const result = { ...section };

  if (section.type === "reading") {
    // ponytail: collect ALL leading italic/quoted paragraphs for epigraph (skipping section header title paragraph at index 0)
    const epigraphLines: string[] = [];
    const epigraphParagraphs: LiturgyParagraph[] = [];

    const nonTitleParagraphs = section.paragraphs.slice(1);

    for (const p of nonTitleParagraphs) {
      // If we hit the true bold citation line (e.g. "I. Inilah Injil Suci menurut Matius..."), break!
      if (p.isBold && isReadingCitationLine(p, getHeadingTitleText(p))) {
        break;
      }

      // Skip Gospel dialogue greetings (e.g. "I. Tuhan bersamamu", "U. Dan bersama roh-Mu")
      if (isGospelDialogueLine(p.rawText)) {
        continue;
      }

      if (
        p.isItalic ||
        p.rawText.startsWith('"') ||
        p.rawText.startsWith("“") ||
        p.rawText.startsWith('*"') ||
        p.rawText.startsWith("*“") ||
        p.rawText.endsWith('"') ||
        p.rawText.endsWith("”")
      ) {
        epigraphLines.push(p.rawText.replace(/^[*"“']|["”'*]$/g, "").trim());
        epigraphParagraphs.push(p);
      }
    }

    if (epigraphLines.length > 0) {
      result.epigraph = epigraphLines.join(" ");
    }

    // Citation line (bold line defining scripture source)
    const citationP = section.paragraphs.find(
      (p) =>
        p.isBold &&
        (isReadingCitationLine(p, getHeadingTitleText(p)) ||
          p.rawText.includes("Bacaan dari") ||
          p.rawText.includes("Inilah Injil")),
    );
    if (citationP) {
      let rawCitation = citationP.rawText.trim();
      rawCitation = rawCitation.replace(/^([LI][.:\s-]*)?Inilah\s+/i, "");
      rawCitation = rawCitation.replace(/^nilah\s+/i, "");
      rawCitation = rawCitation.replace(/^[LI][.:\s-]+/i, "");
      rawCitation = rawCitation.replace(/^[LI](?=Bacaan|Inilah)/i, "");
      result.citation = rawCitation.trim();
    }

    // Notes body is full text excluding header, epigraph paragraphs, citation line, dialogue greetings, and Demikianlah sabda Tuhan / U Syukur kepada Allah
    const rawBodyLines = section.paragraphs
      .filter(
        (p) =>
          !epigraphParagraphs.includes(p) &&
          p !== citationP &&
          p.rawText !== section.title,
      )
      .map((p) => p.rawText);
    result.notesBody = cleanReadingNotesBody(rawBodyLines, result.epigraph);
  }

  if (
    section.type === "doa_umat" ||
    section.title.toUpperCase().includes("DOA UMAT")
  ) {
    const data = parseDoaUmatData(section);
    result.intentions = data.intentions;
  }

  return result;
}

// ponytail: select & map parsed sections to MVP target placeholders, tracking unmatched sections
export function selectRelevantSections(
  parsed: DocxParseResult,
): MappedLiturgyData {
  const mappedSectionIds = new Set<string>();
  const mappedData: MappedLiturgyData = {
    unmatchedSections: [],
  };

  for (const sec of parsed.sections) {
    const titleNorm = sec.title.toUpperCase();

    // 1. Seruan Tobat
    if (
      titleNorm.includes("SERUAN TOBAT") ||
      (titleNorm.includes("TOBAT") && !mappedData.seruanTobat)
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      const parsedTobat = parseSeruanTobatLines(lines);
      mappedData.seruanTobat =
        parsedTobat.length > 0 ? parsedTobat : lines.slice(0, 3);
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 2. Doa Kolekta (clean formula lines)
    if (
      titleNorm.includes("DOA KOLEKTA") ||
      (titleNorm.includes("KOLEKTA") && !mappedData.doaKolekta)
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaKolekta = cleanPrayerBody(lines, sec.title);
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 3. Bacaan Pertama
    if (
      (titleNorm.includes("BACAAN PERTAMA") ||
        titleNorm.includes("BACAAN I") ||
        (titleNorm.includes("BACAAN") &&
          (titleNorm.includes("PERTAMA") || titleNorm.includes(" 1")))) &&
      !titleNorm.includes("KEDUA") &&
      !titleNorm.includes(" II") &&
      !mappedData.bacaanPertama
    ) {
      mappedData.bacaanPertama = {
        epigraph: sec.epigraph,
        citation: sec.citation,
        notesBody: sec.notesBody,
      };
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 4. Mazmur Tanggapan (split distinct verses directly by paragraph breaks/Enters)
    if (titleNorm.includes("MAZMUR") && !mappedData.mazmurTanggapan) {
      const verses = splitMazmurVerses(sec.paragraphs);

      mappedData.mazmurTanggapan = {
        verses: verses.length > 0 ? verses : sec.body || [],
      };
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 5. Bacaan Kedua
    if (
      (titleNorm.includes("BACAAN KEDUA") ||
        titleNorm.includes("BACAAN II") ||
        (titleNorm.includes("BACAAN") &&
          (titleNorm.includes("KEDUA") || titleNorm.includes(" 2")))) &&
      !mappedData.bacaanKedua
    ) {
      mappedData.bacaanKedua = {
        epigraph: sec.epigraph,
        citation: sec.citation,
        notesBody: sec.notesBody,
      };
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 6. Bacaan Injil (excluding Bait Pengantar Injil / Alleluia)
    if (
      (titleNorm.includes("INJIL") || titleNorm.includes("BACAAN INJIL")) &&
      !titleNorm.includes("BAIT PENGANTAR") &&
      !titleNorm.includes("PENGANTAR") &&
      !mappedData.bacaanInjil
    ) {
      let gospelName = "";
      const match = sec.citation?.match(
        /(Matius|Markus|Lukas|Yohanes)\s+\d+:[0-9a-z.-]+/i,
      );
      if (match) {
        gospelName = match[0];
      } else if (sec.citation) {
        gospelName = sec.citation.replace(/^L\s*\*\*|\*\*/g, "").trim();
      }

      mappedData.bacaanInjil = {
        gospelName,
        epigraph: sec.epigraph,
        citation: sec.citation,
        notesBody: sec.notesBody,
      };
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 7. Doa Umat (Universal parseDoaUmatData regardless of section header formatting)
    if (titleNorm.includes("DOA UMAT") && !mappedData.doaUmat) {
      mappedData.doaUmat = parseDoaUmatData(sec);
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 8. Doa Atas Persembahan (clean formula lines)
    if (
      (titleNorm.includes("DOA ATAS PERSEMBAHAN") ||
        titleNorm.includes("ATAS PERSEMBAHAN")) &&
      !mappedData.doaAtasPersembahan
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaAtasPersembahan = cleanPrayerBody(lines, sec.title);
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 9. Doa Sesudah Komuni (clean formula lines)
    if (
      (titleNorm.includes("DOA SESUDAH KOMUNI") ||
        titleNorm.includes("SESUDAH KOMUNI")) &&
      !mappedData.doaSesudahKomuni
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaSesudahKomuni = cleanPrayerBody(lines, sec.title);
      mappedSectionIds.add(sec.id);
      continue;
    }
  }

  // Record unmapped section titles
  for (const sec of parsed.sections) {
    if (!mappedSectionIds.has(sec.id)) {
      mappedData.unmatchedSections.push(sec.title);
    }
  }

  return mappedData;
}
