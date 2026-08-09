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

// ponytail: extract text and run formatting from word/document.xml
export async function parseDocxFile(file: File): Promise<DocxParseResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);

  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) {
    throw new Error("Berkas .docx tidak valid: word/document.xml tidak ditemukan.");
  }

  const xmlContent = await documentXmlFile.async("text");

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    parseTagValue: false,
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

      // ponytail: fixed underline detection logic
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

    const rawText = runs.map((r) => r.text).join("").trim();
    if (rawText.length > 0) {
      // Paragraph properties fallback
      const pPr = pNode["w:pPr"];
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

// ponytail: section classifier based on bold/underline headers
function groupIntoSections(paragraphs: LiturgyParagraph[]): ParsedLiturgySection[] {
  const sections: ParsedLiturgySection[] = [];
  let currentSection: ParsedLiturgySection | null = null;

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const textUpper = p.rawText.toUpperCase();

    // 1. Major Section Header (Bold + Underline)
    if (p.isBold && p.isUnderline) {
      if (currentSection) sections.push(currentSection);
      currentSection = {
        id: `section-${sections.length + 1}`,
        title: p.rawText,
        type: "major_header",
        paragraphs: [p],
        body: [],
      };
      continue;
    }

    // 2. Sub-Section Header (Bold only, not underlined, not italic)
    // ponytail: added !p.isItalic to avoid false positives on short bold+italic responses
    if (p.isBold && !p.isUnderline && !p.isItalic && p.rawText.length < 80) {
      if (currentSection) sections.push(currentSection);

      let type: ParsedLiturgySection["type"] = "sub_header";
      if (textUpper.includes("BACAAN") || textUpper.includes("INJIL")) {
        type = "reading";
      } else if (textUpper.includes("MAZMUR")) {
        type = "psalm";
      } else if (textUpper.includes("DOA UMAT")) {
        type = "doa_umat";
      } else if (textUpper.includes("DOA")) {
        type = "prayer";
      }

      currentSection = {
        id: `section-${sections.length + 1}`,
        title: p.rawText,
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

  // Refine section data (epigraph, citation, intentions)
  return sections.map((sec) => processSectionDetails(sec));
}

// ponytail: extract reading epigraphs (multi-paragraph), citations, and Doa Umat intentions
function processSectionDetails(section: ParsedLiturgySection): ParsedLiturgySection {
  const result = { ...section };

  if (section.type === "reading") {
    // ponytail: collect ALL leading italic/quoted paragraphs for epigraph before bold citation line is hit
    const epigraphLines: string[] = [];
    const epigraphParagraphs: LiturgyParagraph[] = [];

    for (const p of section.paragraphs) {
      if (p.isBold) break; // Citation starts with bold line
      if (
        p.isItalic ||
        p.rawText.startsWith('"') ||
        p.rawText.startsWith('“') ||
        p.rawText.startsWith('*"') ||
        p.rawText.startsWith('*“')
      ) {
        epigraphLines.push(p.rawText.replace(/^[*"“']|["”'*]$/g, "").trim());
        epigraphParagraphs.push(p);
      }
    }

    if (epigraphLines.length > 0) {
      result.epigraph = epigraphLines.join(" ");
    }

    // Citation line (bold line with Bacaan dari / Inilah Injil / speaker label L)
    const citationP = section.paragraphs.find(
      (p) =>
        p.isBold &&
        (p.rawText.includes("Bacaan dari") ||
          p.rawText.includes("Inilah Injil") ||
          /^L[\s.]/.test(p.rawText))
    );
    if (citationP) {
      result.citation = citationP.rawText;
    }

    // Notes body is full text excluding header, epigraph paragraphs, and citation line
    const bodyLines = section.paragraphs
      .filter(
        (p) =>
          !p.isBold &&
          !epigraphParagraphs.includes(p) &&
          p !== citationP &&
          p.rawText !== section.title
      )
      .map((p) => p.rawText);
    result.notesBody = bodyLines.join("\n\n");
  }

  if (section.type === "doa_umat") {
    const intentions: { title: string; body: string }[] = [];
    let currentTitle = "";
    let currentBody: string[] = [];

    for (const p of section.paragraphs) {
      // ponytail: regex /^L[\s.]/ for robust intention title matching
      if (p.isBold && /^L[\s.]/.test(p.rawText)) {
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

    // 2. Doa Kolekta
    if (titleNorm.includes("DOA KOLEKTA") || (titleNorm.includes("KOLEKTA") && !mappedData.doaKolekta)) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaKolekta = lines.join("\n\n");
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
      // Collect numbered verses (bait)
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

    // 6. Bacaan Injil
    if ((titleNorm.includes("INJIL") || titleNorm.includes("BACAAN INJIL")) && !mappedData.bacaanInjil) {
      // Extract gospel name / book citation (e.g. Matius 14:22-33)
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
      const introP = nonHeaderParagraphs.find((p) => !p.isBold || !/^L[\s.]/.test(p.rawText));
      const closingP = nonHeaderParagraphs.slice(-1).find((p) => !/^L[\s.]/.test(p.rawText) && p !== introP);

      mappedData.doaUmat = {
        intro: introP?.rawText,
        intentions: sec.intentions || [],
        closing: closingP?.rawText,
      };
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 8. Doa Atas Persembahan
    if (
      (titleNorm.includes("DOA ATAS PERSEMBAHAN") || titleNorm.includes("ATAS PERSEMBAHAN")) &&
      !mappedData.doaAtasPersembahan
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaAtasPersembahan = lines.join("\n\n");
      mappedSectionIds.add(sec.id);
      continue;
    }

    // 9. Doa Sesudah Komuni
    if (
      (titleNorm.includes("DOA SESUDAH KOMUNI") || titleNorm.includes("SESUDAH KOMUNI")) &&
      !mappedData.doaSesudahKomuni
    ) {
      const lines = sec.body?.filter((b) => b !== sec.title) || [];
      mappedData.doaSesudahKomuni = lines.join("\n\n");
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
