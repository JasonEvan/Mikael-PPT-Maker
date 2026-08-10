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
  type: "major_header" | "sub_header" | "reading" | "psalm" | "prayer" | "doa_umat" | "generic";
  epigraph?: string;
  citation?: string;
  body?: string[];
  notesBody?: string;
  intentions?: { title: string; body: string }[];
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
  intentions: { title: string; body: string }[];
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
  const titleText = titleRuns.map((r) => r.text).join("").trim();
  return titleText.length > 0 ? titleText : null;
}

// ponytail: helper to clean formula lines (Marilah kita berdoa, U. Amin) from prayer bodies
function cleanPrayerBody(lines: string[]): string {
  const filtered = lines.filter((line) => {
    const trimmed = line.trim();
    if (!trimmed) return false;
    if (/^I[\s.]*Marilah/i.test(trimmed)) return false;
    if (/^Marilah kita berdoa/i.test(trimmed)) return false;
    if (/^\(?hening sejenak\)?$/i.test(trimmed)) return false;
    if (/^U[\s.]*Amin/i.test(trimmed)) return false;
    if (/^Amin\.?$/i.test(trimmed)) return false;
    return true;
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

// ponytail: helper to clean reading notes body text (excluding header, epigraph, citation, dialogue greetings, Demikianlah sabda Tuhan, U Syukur kepada Allah)
function cleanReadingNotesBody(lines: string[], epigraph?: string): string {
  const normEpigraph = epigraph ? epigraph.replace(/^[*"“']|["”'*]$/g, "").trim().toLowerCase() : "";

  const filtered = lines.filter((line) => {
    const trimmed = line.trim();
    if (!trimmed) return false;

    // Filter out section headers like "Bacaan I(umat duduk)", "Bacaan II", "Injil"
    if (/^bacaan\s*(I|II|pertama|kedua|injil)/i.test(trimmed)) return false;
    if (/^injil/i.test(trimmed)) return false;
    if (/^\(?umat\s*(duduk|berdiri|berlutut)\)?$/i.test(trimmed)) return false;

    // Filter out Gospel dialogue greetings
    if (isGospelDialogueLine(trimmed)) return false;

    // Filter out epigraph matching lines
    if (normEpigraph && (trimmed.toLowerCase().includes(normEpigraph) || normEpigraph.includes(trimmed.toLowerCase()))) {
      return false;
    }

    // Filter out closing responses
    if (/^demikianlah\s*(sabda|injil)/i.test(trimmed)) return false;
    if (/^U[\s.]*(syukur|terpujilah)/i.test(trimmed)) return false;
    if (/^(syukur kepada allah|terpujilah kristus)\.?$/i.test(trimmed)) return false;

    return true;
  });
  return filtered.join("\n\n").trim();
}

// ponytail: helper to check if paragraph is a scripture reading citation line (defining source)
function isReadingCitationLine(p: LiturgyParagraph, headingTitle: string | null): boolean {
  const text = (headingTitle || p.rawText).trim();

  // If it's a greeting dialogue line (e.g. "I. Tuhan bersamamu"), it is NOT a citation line!
  if (isGospelDialogueLine(text)) return false;

  // Scripture source keywords (e.g. "Inilah Injil", "Bacaan dari", "Matius", "Markus", "Lukas", "Yohanes", "Kitab", "Surat")
  if (/bacaan dari|inilah injil|injil suci|kitab|surat|kisah|warta|nubuat|petikan/i.test(text)) return true;
  if (/^[LI][.:\s-]*(bacaan|inilah|kitab|surat|kisah|warta|nubuat|petikan)/i.test(text)) return true;

  return false;
}

// ponytail: extract text and run formatting from word/document.xml
export async function parseDocxFile(file: File): Promise<DocxParseResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);

  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) {
    throw new Error("Berkas .docx tidak valid: word/document.xml tidak ditemukan.");
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
            rPr["w:b"] === "")
      );

      // Underline detection logic
      const uVal = rPr?.["w:u"]?.["@_w:val"];
      const hasUnderline = Boolean(rPr && "w:u" in rPr) && uVal !== "none";

      const hasItalic = Boolean(
        rPr &&
          ("w:i" in rPr ||
            rPr["w:i"]?.["@_w:val"] === "true" ||
            rPr["w:i"]?.["@_w:val"] === "1" ||
            rPr["w:i"] === "")
      );

      // Extract text
      let text = "";
      if (rNode["w:t"]) {
        if (typeof rNode["w:t"] === "string") {
          text = rNode["w:t"];
        } else if (typeof rNode["w:t"] === "object" && "#text" in rNode["w:t"]) {
          text = rNode["w:t"]["#text"];
        }
      } else if (rNode["w:tab"]) {
        text = " ";
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
    const hasNumPr = Boolean(pPr && ("w:numPr" in pPr || "w:outlineLvl" in pPr));

    // ponytail: strip Word auto-numbering prefixes if numPr present or if fused like "1.HARI"
    if (hasNumPr || /^\d+\.[A-Za-z\u00C0-\u024F]/.test(rawText.trim())) {
      rawText = rawText.trim().replace(/^\d+\.\s*/, "");
    } else {
      rawText = rawText.trim();
    }

    // ponytail: normalize fused speaker labels (e.g. "L.Bacaan" -> "L. Bacaan", "I.Inilah" -> "I. Inilah")
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
    const isJustSpeakerLabel =
      headingTitle !== null && /^[ULI][.:]?$/i.test(headingTitle.trim());

    // ponytail: check if paragraph is a scripture reading citation line (defining source)
    const isCitationLine = isReadingCitationLine(p, headingTitle);

    // 2. Sub-Section Header condition
    const shouldOpenNewSection =
      headingTitle !== null &&
      headingTitle.length < 80 &&
      !isJustSpeakerLabel &&
      !(currentSection?.type === "reading" && isCitationLine);

    if (shouldOpenNewSection) {
      if (currentSection) sections.push(currentSection);

      const titleUpper = headingTitle.toUpperCase();
      let type: ParsedLiturgySection["type"] = "sub_header";
      if (titleUpper.includes("BAIT PENGANTAR") || titleUpper.includes("ALLELUIA")) {
        type = "sub_header";
      } else if (titleUpper.includes("BACAAN") || titleUpper.includes("INJIL")) {
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
function processSectionDetails(section: ParsedLiturgySection): ParsedLiturgySection {
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
        p.rawText.startsWith('“') ||
        p.rawText.startsWith('*"') ||
        p.rawText.startsWith('*“') ||
        p.rawText.endsWith('"') ||
        p.rawText.endsWith('”')
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
          p.rawText.includes("Inilah Injil"))
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
          p.rawText !== section.title
      )
      .map((p) => p.rawText);
    result.notesBody = cleanReadingNotesBody(rawBodyLines, result.epigraph);
  }

  if (section.type === "doa_umat") {
    const intentions: { title: string; body: string }[] = [];
    let currentTitle = "";
    let currentBody: string[] = [];

    for (const p of section.paragraphs) {
      if (p.isBold && /^L[\s.]/i.test(p.rawText)) {
        if (currentTitle) {
          intentions.push({ title: currentTitle, body: currentBody.join(" ") });
        }
        currentTitle = p.rawText;
        currentBody = [];
      } else if (currentTitle) {
        currentBody.push(p.rawText);
      }
    }
    if (currentTitle) {
      intentions.push({ title: currentTitle, body: currentBody.join(" ") });
    }
    if (intentions.length > 0) {
      result.intentions = intentions;
    }
  }

  return result;
}

// ponytail: select & map parsed sections to MVP target placeholders, tracking unmatched sections
export function selectRelevantSections(parsed: DocxParseResult): MappedLiturgyData {
  const mappedSectionIds = new Set<string>();
  const mappedData: MappedLiturgyData = {
    unmatchedSections: [],
  };

  for (const sec of parsed.sections) {
    const titleNorm = sec.title.toUpperCase();

    // 1. Seruan Tobat
    if (titleNorm.includes("SERUAN TOBAT") || (titleNorm.includes("TOBAT") && !mappedData.seruanTobat)) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.seruanTobat = lines.slice(0, 3);
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 2. Doa Kolekta (clean formula lines)
    if (titleNorm.includes("DOA KOLEKTA") || (titleNorm.includes("KOLEKTA") && !mappedData.doaKolekta)) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaKolekta = cleanPrayerBody(lines);
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 3. Bacaan Pertama
    if (
      (titleNorm.includes("BACAAN PERTAMA") ||
        titleNorm.includes("BACAAN I") ||
        (titleNorm.includes("BACAAN") && (titleNorm.includes("PERTAMA") || titleNorm.includes(" 1")))) &&
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

    // 4. Mazmur Tanggapan
    if (titleNorm.includes("MAZMUR") && !mappedData.mazmurTanggapan) {
      const verses: string[] = [];
      const nonHeaderParagraphs = sec.paragraphs.filter((p) => p.rawText !== sec.title);
      let currentVerse = "";

      for (const p of nonHeaderParagraphs) {
        if (/^\d+[.\s]/.test(p.rawText) || (p.isBold && verses.length < 4)) {
          if (currentVerse) verses.push(currentVerse);
          currentVerse = p.rawText;
        } else if (currentVerse) {
          currentVerse += "\n" + p.rawText;
        } else {
          currentVerse = p.rawText;
        }
      }
      if (currentVerse) verses.push(currentVerse);

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
        (titleNorm.includes("BACAAN") && (titleNorm.includes("KEDUA") || titleNorm.includes(" 2")))) &&
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
      const match = sec.citation?.match(/(Matius|Markus|Lukas|Yohanes)\s+\d+:[0-9a-z.-]+/i);
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

    // 7. Doa Umat
    if (titleNorm.includes("DOA UMAT") && !mappedData.doaUmat) {
      const nonHeaderParagraphs = sec.paragraphs.filter((p) => p.rawText !== sec.title);
      const introP = nonHeaderParagraphs.find((p) => !p.isBold || !/^L[\s.]/i.test(p.rawText));
      const closingP = nonHeaderParagraphs.slice(-1).find((p) => !/^L[\s.]/i.test(p.rawText) && p !== introP);

      mappedData.doaUmat = {
        intro: introP?.rawText,
        intentions: sec.intentions || [],
        closing: closingP?.rawText,
      };
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 8. Doa Atas Persembahan (clean formula lines)
    if (
      (titleNorm.includes("DOA ATAS PERSEMBAHAN") || titleNorm.includes("ATAS PERSEMBAHAN")) &&
      !mappedData.doaAtasPersembahan
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaAtasPersembahan = cleanPrayerBody(lines);
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 9. Doa Sesudah Komuni (clean formula lines)
    if (
      (titleNorm.includes("DOA SESUDAH KOMUNI") || titleNorm.includes("SESUDAH KOMUNI")) &&
      !mappedData.doaSesudahKomuni
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaSesudahKomuni = cleanPrayerBody(lines);
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
