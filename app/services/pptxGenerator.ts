import fs from "fs";
import path from "path";
import JSZip from "jszip";
import { MappedLiturgyData } from "../utils/docxParser";

// ponytail: helper to escape XML special characters
function escapeXml(str: string): string {
  if (!str) return "";
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// ponytail: format epigraph wrapped in quotation marks “...”
function formatEpigraph(epigraph?: string): string {
  if (!epigraph || !epigraph.trim()) return "";
  const trimmed = epigraph.trim();
  if (trimmed.startsWith("“") || trimmed.startsWith('"')) {
    return trimmed;
  }
  return `“${trimmed}”`;
}

// ponytail: helper to strip leading L / I / L. / I. / Inilah / nilah prefix from reading citations
function cleanCitation(citation?: string): string {
  if (!citation) return "";
  let cleaned = citation.trim();

  // Strip "Inilah " or "I. Inilah " or "nilah " from Injil citations
  cleaned = cleaned.replace(/^([LI][.:\s-]*)?Inilah\s+/i, "");
  cleaned = cleaned.replace(/^nilah\s+/i, "");

  // Strip speaker prefix (L. / I. / L / I) if followed by punctuation, space, OR uppercase word like Bacaan/Inilah
  cleaned = cleaned.replace(/^[LI][.:\s-]+/i, "");
  cleaned = cleaned.replace(/^[LI](?=Bacaan|Inilah)/i, "");

  return cleaned.trim();
}

// ponytail: helper to check if a slide XML contains a (teks) placeholder
function slideHasPlaceholder(slideXml: string): boolean {
  const pMatches = slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
  for (const p of pMatches) {
    const tMatches = p.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
    const text = tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");
    if (/\(\s*teks\s*\)/i.test(text)) {
      return true;
    }
  }
  return false;
}

// ponytail: helper to count (teks) placeholders in a slide XML
function getPlaceholderCount(slideXml: string): number {
  const pMatches = slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
  let count = 0;
  for (const p of pMatches) {
    const tMatches = p.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
    const text = tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");
    if (/\(\s*teks\s*\)/i.test(text)) {
      count++;
    }
  }
  return count;
}

// ponytail: replace placeholders ( (teks) or (teks) ) in slide XML, handling run-splitting
function replacePlaceholdersInXml(
  slideXml: string,
  replacements: string[],
): string {
  let replacementIndex = 0;

  return slideXml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (pBlock) => {
    // Extract text from all <a:t> tags in this paragraph
    const tMatches = pBlock.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
    const fullText = tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");

    if (/\(\s*teks\s*\)/i.test(fullText)) {
      const valToInsert =
        replacementIndex < replacements.length
          ? replacements[replacementIndex]
          : "";
      replacementIndex++;

      const newText = fullText.replace(/\(\s*teks\s*\)/gi, valToInsert);

      let isFirstT = true;
      return pBlock.replace(/<a:t([^>]*)>(.*?)<\/a:t>/g, (tBlock, attrs) => {
        if (isFirstT) {
          isFirstT = false;
          return `<a:t${attrs}>${escapeXml(newText)}</a:t>`;
        }
        return `<a:t${attrs}></a:t>`;
      });
    }

    return pBlock;
  });
}

// ponytail: helper to fill Doa Umat intention slide (Placeholder 1 = Bold Title + Unbold Body separated by \n, Placeholder 2 = Yellow Italic Response)
function replaceDoaUmatIntentionInXml(
  slideXml: string,
  title: string,
  body: string,
  response?: string,
): string {
  let replacementIndex = 0;

  return slideXml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (pBlock) => {
    const tMatches = pBlock.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
    const fullText = tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");

    if (/\(\s*teks\s*\)/i.test(fullText)) {
      if (replacementIndex === 0) {
        replacementIndex++;

        const pPrMatch = pBlock.match(/<a:pPr\b[\s\S]*?<\/a:pPr>/i);
        const pPrXml = pPrMatch ? pPrMatch[0] : "";

        const rPrMatch = pBlock.match(/<a:rPr\b[\s\S]*?<\/a:rPr>/i);
        let rPrBoldXml = rPrMatch
          ? rPrMatch[0]
          : '<a:rPr lang="en-US" sz="4800" b="1"/>';

        if (!/b="1"/i.test(rPrBoldXml)) {
          rPrBoldXml = rPrBoldXml.replace(/<a:rPr/i, '<a:rPr b="1"');
        }

        const rPrUnboldXml = rPrBoldXml
          .replace(/\bb="1"/gi, 'b="0"')
          .replace(/\bb="true"/gi, 'b="0"');

        if (title && body) {
          const titleParagraph = `<a:p>${pPrXml}<a:r>${rPrBoldXml}<a:t>${escapeXml(title)}</a:t></a:r></a:p>`;
          const bodyParagraph = `<a:p>${pPrXml}<a:r>${rPrUnboldXml}<a:t>${escapeXml(body)}</a:t></a:r></a:p>`;
          return `${titleParagraph}${bodyParagraph}`;
        } else {
          const singleText = title || body || "";
          const rPrStyle = title ? rPrBoldXml : rPrUnboldXml;
          return `<a:p>${pPrXml}<a:r>${rPrStyle}<a:t>${escapeXml(singleText)}</a:t></a:r></a:p>`;
        }
      } else if (replacementIndex === 1) {
        replacementIndex++;
        const responseText = response ? `U\t${response}` : "";

        let isFirstT = true;
        return pBlock.replace(/<a:t([^>]*)>(.*?)<\/a:t>/g, (tBlock, attrs) => {
          if (isFirstT) {
            isFirstT = false;
            return `<a:t${attrs}>${escapeXml(responseText)}</a:t>`;
          }
          return `<a:t${attrs}></a:t>`;
        });
      }
    }

    return pBlock;
  });
}

// ponytail: helper to structurally remove all intro paragraphs appearing BEFORE the (teks) placeholder in txBody
function stripIntroParagraphs(slideXml: string): string {
  return slideXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/g, (txBodyBlock) => {
    const pBlocks = txBodyBlock.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
    let placeholderIdx = -1;

    for (let i = 0; i < pBlocks.length; i++) {
      const tMatches = pBlocks[i].match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
      const fullText = tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");
      if (/\(\s*teks\s*\)/i.test(fullText)) {
        placeholderIdx = i;
        break;
      }
    }

    if (placeholderIdx <= 0) {
      return txBodyBlock;
    }

    const remainingP = pBlocks.slice(placeholderIdx).join("");
    return txBodyBlock.replace(/<a:p>[\s\S]*?<\/a:p>/g, (match, offset) => {
      if (offset === txBodyBlock.indexOf("<a:p>")) {
        return remainingP;
      }
      return "";
    });
  });
}

// ponytail: helper to structurally remove all closing/response paragraphs appearing AFTER the (teks) placeholder in txBody
function stripClosingParagraphs(slideXml: string): string {
  return slideXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/g, (txBodyBlock) => {
    const pBlocks = txBodyBlock.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
    let placeholderIdx = -1;

    for (let i = 0; i < pBlocks.length; i++) {
      const tMatches = pBlocks[i].match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
      const fullText = tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");
      if (/\(\s*teks\s*\)/i.test(fullText)) {
        placeholderIdx = i;
        break;
      }
    }

    if (placeholderIdx < 0 || placeholderIdx >= pBlocks.length - 1) {
      return txBodyBlock;
    }

    const remainingP = pBlocks.slice(0, placeholderIdx + 1).join("");
    return txBodyBlock.replace(/<a:p>[\s\S]*?<\/a:p>/g, (match, offset) => {
      if (offset === txBodyBlock.indexOf("<a:p>")) {
        return remainingP;
      }
      return "";
    });
  });
}

// ponytail: update speaker notes in ppt/notesSlides/notesSlideX.xml, completely replacing template notes
function updateNotesSlideXml(notesXml: string, notesText: string): string {
  if (!notesText) return notesXml;
  const escapedNotes = escapeXml(notesText);

  const paragraphs = escapedNotes
    .split(/\n\n+/)
    .map((para) => `<a:p><a:r><a:t>${para}</a:t></a:r></a:p>`)
    .join("");

  // Target the txBody inside the shape that contains type="body" placeholder
  return notesXml.replace(
    /(<p:sp>[\s\S]*?<p:ph[^>]*type="body"[\s\S]*?<p:txBody>)([\s\S]*?)(<\/p:txBody>[\s\S]*?<\/p:sp>)/i,
    (match, prefix) => {
      return `${prefix}<a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
    },
  );
}

// ponytail: dynamically calculate overflow cut points for any text based on character limit & punctuation boundaries (. , ! ? ;)
function splitTextIntoSlideChunks(
  text: string,
  maxCharsPerSlide: number = 320,
): string[] {
  if (!text || text.trim().length === 0) return [""];
  const trimmed = text.trim();

  // Return single chunk if text fits within maxCharsPerSlide
  if (trimmed.length <= maxCharsPerSlide) {
    return [trimmed];
  }

  const chunks: string[] = [];
  let remaining = trimmed;

  while (remaining.length > maxCharsPerSlide) {
    const slice = remaining.slice(0, maxCharsPerSlide);

    let lastPunctuationEnd = -1;
    const punctuationRegex = /[.,!?;]\s+/g;
    let match: RegExpExecArray | null;

    while ((match = punctuationRegex.exec(slice)) !== null) {
      lastPunctuationEnd = match.index + match[0].length - 1;
    }

    let cutIndex = -1;

    // Prefer cutting at punctuation boundary if it occurs past 35% of maxCharsPerSlide
    if (lastPunctuationEnd > Math.floor(maxCharsPerSlide * 0.35)) {
      cutIndex = lastPunctuationEnd + 1; // Include punctuation mark
    } else {
      // Fallback: cut at last space (word boundary)
      const lastSpaceIndex = slice.lastIndexOf(" ");
      if (lastSpaceIndex > Math.floor(maxCharsPerSlide * 0.35)) {
        cutIndex = lastSpaceIndex;
      } else {
        // Fallback to exact limit
        cutIndex = maxCharsPerSlide;
      }
    }

    const chunk = remaining.slice(0, cutIndex).trim();
    if (chunk.length > 0) {
      chunks.push(chunk);
    }
    remaining = remaining.slice(cutIndex).trim();
  }

  if (remaining.length > 0) {
    chunks.push(remaining);
  }

  return chunks.length > 0 ? chunks : [trimmed];
}

// ponytail: helper to generically split off the last sentence by punctuation boundary to accompany the congregation response on an overflow slide (ZERO hardcoded words)
function splitLastSentence(bodyText: string): {
  mainBody: string;
  lastSentence: string;
} {
  const trimmed = bodyText.trim();
  if (trimmed.length === 0) return { mainBody: "", lastSentence: "" };

  // Strip trailing punctuation when searching for the sentence break point
  const searchSlice = trimmed.replace(/[.,!?;…]+$/, "").trim();

  let lastPunctIdx = -1;
  const delimiters = [". ", "! ", "? ", "; ", "… ", "…", "\n"];

  for (const delim of delimiters) {
    const idx = searchSlice.lastIndexOf(delim);
    if (idx > lastPunctIdx) {
      lastPunctIdx = idx;
    }
  }

  if (lastPunctIdx > 0 && lastPunctIdx > Math.floor(trimmed.length * 0.35)) {
    const cutIndex = lastPunctIdx + 1;
    const mainBody = trimmed.slice(0, cutIndex).trim();
    const lastSentence = trimmed.slice(cutIndex).trim();

    if (mainBody.length > 0 && lastSentence.length > 0) {
      return { mainBody, lastSentence };
    }
  }

  return { mainBody: trimmed, lastSentence: "" };
}

// ponytail: generate filled PPTX from template and MappedLiturgyData
export async function generatePPTXBuffer(
  mappedData: MappedLiturgyData,
): Promise<Buffer> {
  const templatePath = path.join(
    process.cwd(),
    "app/template/Template_Mingguan.pptx",
  );
  if (!fs.existsSync(templatePath)) {
    throw new Error(
      "Template PowerPoint (Template_Mingguan.pptx) tidak ditemukan.",
    );
  }

  const templateBuffer = await fs.promises.readFile(templatePath);
  const zip = await JSZip.loadAsync(templateBuffer);

  // 1. Read presentation structure and slide mapping
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const presentationRels = await zip
    .file("ppt/_rels/presentation.xml.rels")
    ?.async("text");

  if (!presentationXml || !presentationRels) {
    throw new Error("Struktur berkas PowerPoint tidak valid.");
  }

  const rIdMap: Record<string, string> = {};
  const relMatches = presentationRels.matchAll(
    /<Relationship [^>]*Id="([^"]+)" [^>]*Target="([^"]+)"/g,
  );
  for (const match of relMatches) {
    rIdMap[match[1]] = match[2];
  }

  const sldIdMatches = presentationXml.matchAll(
    /<p:sldId [^>]*r:id="([^"]+)"[^>]*>/g,
  );
  const slideOrder: { rId: string; fileName: string }[] = [];
  for (const match of sldIdMatches) {
    const rId = match[1];
    const target = rIdMap[rId];
    if (target && target.startsWith("slides/slide")) {
      slideOrder.push({ rId, fileName: `ppt/${target}` });
    }
  }

  // Extract maximum numeric rId from presentation.xml.rels
  let maxRIdNum = 0;
  for (const rIdStr of Object.keys(rIdMap)) {
    const numMatch = rIdStr.match(/^rId(\d+)$/);
    if (numMatch) {
      const val = parseInt(numMatch[1], 10);
      if (val > maxRIdNum) maxRIdNum = val;
    }
  }

  // Extract max sldId id from presentation.xml
  let maxSldId = 255;
  const allSldIdMatches = presentationXml.matchAll(/<p:sldId [^>]*id="(\d+)"/g);
  for (const match of allSldIdMatches) {
    const val = parseInt(match[1], 10);
    if (val > maxSldId) maxSldId = val;
  }

  // ponytail: helper to dynamically duplicate a slide file, update Relationships & presentation.xml
  const duplicateSlide = async (
    sourceFileName: string,
    targetRId: string,
    newSlideXml: string,
  ): Promise<{ newSlideNum: number; newFileName: string; newRId: string }> => {
    const slideFiles = Object.keys(zip.files).filter((f) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(f),
    );
    const existingNums = slideFiles.map((f) => {
      const match = f.match(/slide(\d+)\.xml/);
      return match ? parseInt(match[1], 10) : 0;
    });
    const newSlideNum = Math.max(...existingNums, 0) + 1;
    const newFileName = `ppt/slides/slide${newSlideNum}.xml`;

    maxRIdNum += 1;
    const newRId = `rId${maxRIdNum}`;
    rIdMap[newRId] = `slides/slide${newSlideNum}.xml`;

    maxSldId += 1;
    const newSldId = maxSldId;

    // Write duplicated slide XML
    zip.file(newFileName, newSlideXml);

    // Duplicate source slide's .rels file if present
    const sourceSlideNumMatch = sourceFileName.match(/slide(\d+)\.xml/);
    const sourceSlideNum = sourceSlideNumMatch ? sourceSlideNumMatch[1] : "1";
    const relsFileName = `slide${sourceSlideNum}.xml.rels`;
    const sourceRels = `ppt/slides/_rels/${relsFileName}`;
    const sourceRelsFile = zip.file(sourceRels);

    if (sourceRelsFile) {
      let relsXml = await sourceRelsFile.async("text");
      // Strip notesSlide relationship from duplicated slide to prevent OpenXML schema collision
      relsXml = relsXml.replace(
        /<Relationship [^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesSlide"[^>]*\/>/g,
        "",
      );
      zip.file(`ppt/slides/_rels/slide${newSlideNum}.xml.rels`, relsXml);
    }

    // Update [Content_Types].xml
    let contentTypes = await zip.file("[Content_Types].xml")?.async("text");
    if (contentTypes) {
      const overrideTag = `<Override PartName="/ppt/slides/slide${newSlideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
      contentTypes = contentTypes.replace("</Types>", `${overrideTag}</Types>`);
      zip.file("[Content_Types].xml", contentTypes);
    }

    // Update ppt/_rels/presentation.xml.rels using valid numeric rId (e.g. rId125)
    let presRels = await zip
      .file("ppt/_rels/presentation.xml.rels")
      ?.async("text");
    if (presRels) {
      const relTag = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNum}.xml"/>`;
      presRels = presRels.replace(
        "</Relationships>",
        `${relTag}</Relationships>`,
      );
      zip.file("ppt/_rels/presentation.xml.rels", presRels);
    }

    // Update ppt/presentation.xml - insert IMMEDIATELY AFTER targetRId's <p:sldId>
    let presXml = await zip.file("ppt/presentation.xml")?.async("text");
    if (presXml) {
      const newSldTag = `<p:sldId id="${newSldId}" r:id="${newRId}"/>`;
      const targetRegex = new RegExp(
        `(<p:sldId [^>]*r:id="${targetRId}"[^>]*\\/?>)`,
      );
      if (targetRegex.test(presXml)) {
        presXml = presXml.replace(targetRegex, `$1${newSldTag}`);
      } else {
        presXml = presXml.replace("</p:sldIdLst>", `${newSldTag}</p:sldIdLst>`);
      }
      zip.file("ppt/presentation.xml", presXml);
    }

    return { newSlideNum, newFileName, newRId };
  };

  // Tracking section states dynamically across content slides
  let seruanCount = 0;
  let hasHandledMazmur = false;
  let hasPassedBacaanPertama = false;
  let hasPassedBacaanKedua = false;
  let hasPassedDoaUmatTitle = false;
  let hasHandledDoaUmatIntro = false;
  let hasHandledDoaUmatIntentions = false;
  let hasHandledDoaUmatOutro = false;
  let hasPassedDoaKomuniTitle = false;
  let hasHandledDoaKomuni = false;

  // ponytail: process each slide placeholder using Content Slide filtering + Content Pattern Anchors (ZERO slide index dependency)
  for (const slide of slideOrder) {
    const slideZip = zip.file(slide.fileName);
    if (!slideZip) continue;

    let slideXml = await slideZip.async("text");

    // Extract visible slide text to anchor slide type by internal text/placeholder content
    const pMatches = slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
    const slideTextUpper = pMatches
      .map((p) => {
        const tMatches = p.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
        return tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");
      })
      .join(" ")
      .toUpperCase();

    // Check if slide is Doa Umat Title slide (Slide 62: DOA UMAT (umat berdiri))
    if (
      !hasPassedDoaUmatTitle &&
      slideTextUpper.includes("DOA UMAT") &&
      !slideTextUpper.includes("AMIN")
    ) {
      hasPassedDoaUmatTitle = true;
    }

    // Check if slide is Doa Sesudah Komuni Title slide (Slide 91: DOA SESUDAH KOMUNI)
    if (
      !hasPassedDoaKomuniTitle &&
      (slideTextUpper.includes("DOA SESUDAH KOMUNI") ||
        slideTextUpper.includes("SESUDAH KOMUNI"))
    ) {
      hasPassedDoaKomuniTitle = true;
    }

    // Track Bacaan Pertama & Bacaan Kedua titles
    if (
      !hasPassedBacaanPertama &&
      (slideTextUpper.includes("BACAAN PERTAMA") ||
        slideTextUpper.includes("BACAAN I "))
    ) {
      hasPassedBacaanPertama = true;
    }

    if (
      !hasPassedBacaanKedua &&
      (slideTextUpper.includes("BACAAN KEDUA") ||
        slideTextUpper.includes("BACAAN II "))
    ) {
      hasPassedBacaanKedua = true;
    }

    // Only process content slides that actually contain a (teks) placeholder
    const hasPlaceholder = slideHasPlaceholder(slideXml);
    if (!hasPlaceholder) continue;

    // Skip song lyrics and announcement slides by content pattern
    if (
      slideTextUpper.includes("PENGUMUMAN") ||
      slideTextUpper.includes("LIRIK LAGU")
    ) {
      continue;
    }

    // 1. Seruan Tobat Content Slides
    if (slideTextUpper.includes("KASIHANILAH KAMI")) {
      const idx =
        seruanCount < (mappedData.seruanTobat?.length || 0) ? seruanCount : 0;
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.seruanTobat?.[idx] || "",
      ]);
      zip.file(slide.fileName, slideXml);
      seruanCount++;
    }

    // 2. Doa Kolekta Content Slide
    else if (
      slideTextUpper.includes("MARILAH KITA BERDOA") &&
      !hasPassedBacaanPertama
    ) {
      const originalSlideXml = slideXml;
      const kolektaChunks = splitTextIntoSlideChunks(
        mappedData.doaKolekta || "",
        260,
      );

      if (kolektaChunks.length === 1) {
        const filledSlideXml = replacePlaceholdersInXml(originalSlideXml, [
          kolektaChunks[0] || "",
        ]);
        zip.file(slide.fileName, filledSlideXml);
      } else {
        let slide1Xml = stripClosingParagraphs(originalSlideXml);
        slide1Xml = replacePlaceholdersInXml(slide1Xml, [kolektaChunks[0]]);
        zip.file(slide.fileName, slide1Xml);

        let lastRId = slide.rId;
        for (let i = 1; i < kolektaChunks.length; i++) {
          const isLast = i === kolektaChunks.length - 1;
          let dupSlideXml = originalSlideXml;

          dupSlideXml = stripIntroParagraphs(dupSlideXml);
          if (!isLast) {
            dupSlideXml = stripClosingParagraphs(dupSlideXml);
          }

          const filledChunkXml = replacePlaceholdersInXml(dupSlideXml, [
            kolektaChunks[i],
          ]);
          const { newRId } = await duplicateSlide(
            slide.fileName,
            lastRId,
            filledChunkXml,
          );
          lastRId = newRId;
        }
      }
    }

    // 3. Bacaan Pertama Content Slide (+ notesSlide2.xml)
    else if (
      slideTextUpper.includes("BACAAN PERTAMA") ||
      slideTextUpper.includes("BACAAN I ")
    ) {
      const epigraph = formatEpigraph(mappedData.bacaanPertama?.epigraph);
      const citation = cleanCitation(mappedData.bacaanPertama?.citation);
      slideXml = replacePlaceholdersInXml(slideXml, [epigraph, citation]);
      zip.file(slide.fileName, slideXml);

      if (mappedData.bacaanPertama?.notesBody) {
        const notesFile = zip.file("ppt/notesSlides/notesSlide2.xml");
        if (notesFile) {
          const notesXml = await notesFile.async("text");
          const updatedNotes = updateNotesSlideXml(
            notesXml,
            mappedData.bacaanPertama.notesBody,
          );
          zip.file("ppt/notesSlides/notesSlide2.xml", updatedNotes);
        }
      }
    }

    // 4. Mazmur Tanggapan Content Slide (duplication for all verses sequentially)
    else if (
      (slideTextUpper.includes("MAZMUR") ||
        slideTextUpper.includes("TANGGAPAN") ||
        (hasPassedBacaanPertama && !hasPassedBacaanKedua)) &&
      !hasHandledMazmur &&
      !slideTextUpper.includes("BACAAN KEDUA") &&
      !slideTextUpper.includes("BACAAN PERTAMA")
    ) {
      const verses = mappedData.mazmurTanggapan?.verses || [];
      const originalSlideXml = slideXml;

      if (verses.length > 0) {
        // Fill 1st verse into 1st Mazmur placeholder slide
        const filledSlideXml = replacePlaceholdersInXml(originalSlideXml, [
          verses[0] || "",
        ]);
        zip.file(slide.fileName, filledSlideXml);

        // Duplicate slide for all remaining verses (verses[1], verses[2], etc.)
        if (verses.length > 1) {
          let lastRId = slide.rId;
          for (let i = 1; i < verses.length; i++) {
            const filledVerseXml = replacePlaceholdersInXml(originalSlideXml, [
              verses[i],
            ]);
            const { newRId } = await duplicateSlide(
              slide.fileName,
              lastRId,
              filledVerseXml,
            );
            lastRId = newRId;
          }
        }
      }
      hasHandledMazmur = true;
    }

    // 5. Bacaan Kedua Content Slide (+ notesSlide3.xml)
    else if (
      slideTextUpper.includes("BACAAN KEDUA") ||
      slideTextUpper.includes("BACAAN II ")
    ) {
      const epigraph = formatEpigraph(mappedData.bacaanKedua?.epigraph);
      const citation = cleanCitation(mappedData.bacaanKedua?.citation);
      slideXml = replacePlaceholdersInXml(slideXml, [epigraph, citation]);
      zip.file(slide.fileName, slideXml);

      if (mappedData.bacaanKedua?.notesBody) {
        const notesFile = zip.file("ppt/notesSlides/notesSlide3.xml");
        if (notesFile) {
          const notesXml = await notesFile.async("text");
          const updatedNotes = updateNotesSlideXml(
            notesXml,
            mappedData.bacaanKedua.notesBody,
          );
          zip.file("ppt/notesSlides/notesSlide3.xml", updatedNotes);
        }
      }
    }

    // 6. Bacaan Injil - Gospel Citation Intro Slide
    else if (slideTextUpper.includes("INILAH INJIL SUCI MENURUT")) {
      const gospelName = mappedData.bacaanInjil?.gospelName || "";
      slideXml = replacePlaceholdersInXml(slideXml, [gospelName]);
      zip.file(slide.fileName, slideXml);
    }

    // 7. Bacaan Injil - Content Slide (+ notesSlide6.xml)
    else if (
      slideTextUpper.includes("INJIL") &&
      !slideTextUpper.includes("PENGANTAR")
    ) {
      const epigraph = formatEpigraph(mappedData.bacaanInjil?.epigraph);
      const citation = cleanCitation(mappedData.bacaanInjil?.citation);
      slideXml = replacePlaceholdersInXml(slideXml, [epigraph, citation]);
      zip.file(slide.fileName, slideXml);

      if (mappedData.bacaanInjil?.notesBody) {
        const notesFile = zip.file("ppt/notesSlides/notesSlide6.xml");
        if (notesFile) {
          const notesXml = await notesFile.async("text");
          const updatedNotes = updateNotesSlideXml(
            notesXml,
            mappedData.bacaanInjil.notesBody,
          );
          zip.file("ppt/notesSlides/notesSlide6.xml", updatedNotes);
        }
      }
    }

    // 8. Doa Umat Intro Slide (Slide 63: right after DOA UMAT title slide with 1 placeholder, supporting overflow splitting)
    else if (hasPassedDoaUmatTitle && !hasHandledDoaUmatIntro) {
      const count = getPlaceholderCount(slideXml);
      if (count === 1) {
        const originalSlideXml = slideXml;
        const introChunks = splitTextIntoSlideChunks(
          mappedData.doaUmat?.intro || "",
          260,
        );

        const filledIntroXml = replacePlaceholdersInXml(originalSlideXml, [
          introChunks[0] || "",
        ]);
        zip.file(slide.fileName, filledIntroXml);

        if (introChunks.length > 1) {
          let lastRId = slide.rId;
          for (let i = 1; i < introChunks.length; i++) {
            const filledChunkXml = replacePlaceholdersInXml(originalSlideXml, [
              introChunks[i],
            ]);
            const { newRId } = await duplicateSlide(
              slide.fileName,
              lastRId,
              filledChunkXml,
            );
            lastRId = newRId;
          }
        }
        hasHandledDoaUmatIntro = true;
      }
    }

    // 9. Doa Umat Intentions Content Slide (Slide 64: 2 placeholders - Placeholder 1 = Bold Title & Unbold Body (\n), Placeholder 2 = Yellow Italic Response, with overflow splitting)
    else if (
      hasPassedDoaUmatTitle &&
      hasHandledDoaUmatIntro &&
      !hasHandledDoaUmatIntentions
    ) {
      const count = getPlaceholderCount(slideXml);
      if (count >= 2) {
        const originalSlideXml = slideXml;
        const intentions = mappedData.doaUmat?.intentions || [];

        if (intentions.length > 0) {
          let isFirstSlide = true;
          let lastRId = slide.rId;

          const maxCombinedLimit = 250;

          for (let i = 0; i < intentions.length; i++) {
            const item = intentions[i];
            const bodyChunks = splitTextIntoSlideChunks(item.body || "", 260);

            for (let c = 0; c < bodyChunks.length; c++) {
              const bodyChunk = bodyChunks[c];
              const isLastBodyChunk = c === bodyChunks.length - 1;
              const intentionTitle = c === 0 ? item.title : "";
              const responseText = item.response ? item.response : "";

              if (isLastBodyChunk) {
                // Check if bodyChunk + responseText exceeds visual height (maxCombinedLimit)
                if (
                  responseText &&
                  bodyChunk.length + responseText.length > maxCombinedLimit
                ) {
                  const { mainBody, lastSentence } =
                    splitLastSentence(bodyChunk);

                  // Slide 1: Main body chunk (without last sentence / Kami mohon)
                  const filledIntentionXml1 = replaceDoaUmatIntentionInXml(
                    originalSlideXml,
                    intentionTitle,
                    mainBody,
                    "",
                  );

                  if (isFirstSlide) {
                    zip.file(slide.fileName, filledIntentionXml1);
                    isFirstSlide = false;
                  } else {
                    const { newRId } = await duplicateSlide(
                      slide.fileName,
                      lastRId,
                      filledIntentionXml1,
                    );
                    lastRId = newRId;
                  }

                  // Slide 2: Duplicated slide for Last Sentence ("Kami mohon ...") + Response!
                  const filledResponseXml = replaceDoaUmatIntentionInXml(
                    originalSlideXml,
                    "",
                    lastSentence,
                    responseText,
                  );
                  const { newRId } = await duplicateSlide(
                    slide.fileName,
                    lastRId,
                    filledResponseXml,
                  );
                  lastRId = newRId;
                } else {
                  // Both bodyChunk & responseText fit on current slide
                  const filledIntentionXml = replaceDoaUmatIntentionInXml(
                    originalSlideXml,
                    intentionTitle,
                    bodyChunk,
                    responseText,
                  );

                  if (isFirstSlide) {
                    zip.file(slide.fileName, filledIntentionXml);
                    isFirstSlide = false;
                  } else {
                    const { newRId } = await duplicateSlide(
                      slide.fileName,
                      lastRId,
                      filledIntentionXml,
                    );
                    lastRId = newRId;
                  }
                }
              } else {
                // Non-last body chunk (no response)
                const filledIntentionXml = replaceDoaUmatIntentionInXml(
                  originalSlideXml,
                  intentionTitle,
                  bodyChunk,
                  "",
                );

                if (isFirstSlide) {
                  zip.file(slide.fileName, filledIntentionXml);
                  isFirstSlide = false;
                } else {
                  const { newRId } = await duplicateSlide(
                    slide.fileName,
                    lastRId,
                    filledIntentionXml,
                  );
                  lastRId = newRId;
                }
              }
            }
          }
        } else {
          const clearedSlideXml = replacePlaceholdersInXml(originalSlideXml, [
            "",
            "",
          ]);
          zip.file(slide.fileName, clearedSlideXml);
        }
        hasHandledDoaUmatIntentions = true;
      }
    }

    // 10. Doa Umat Outro / Penutup Slide (Slide 66: U Amin. & 1 placeholder, supporting overflow splitting)
    else if (
      hasPassedDoaUmatTitle &&
      hasHandledDoaUmatIntentions &&
      !hasHandledDoaUmatOutro
    ) {
      const originalSlideXml = slideXml;
      const outroChunks = splitTextIntoSlideChunks(
        mappedData.doaUmat?.closing || "",
        260,
      );

      const filledOutroXml = replacePlaceholdersInXml(originalSlideXml, [
        outroChunks[0] || "",
      ]);
      zip.file(slide.fileName, filledOutroXml);

      if (outroChunks.length > 1) {
        let lastRId = slide.rId;
        for (let i = 1; i < outroChunks.length; i++) {
          const filledChunkXml = replacePlaceholdersInXml(originalSlideXml, [
            outroChunks[i],
          ]);
          const { newRId } = await duplicateSlide(
            slide.fileName,
            lastRId,
            filledChunkXml,
          );
          lastRId = newRId;
        }
      }
      hasHandledDoaUmatOutro = true;
      hasPassedDoaUmatTitle = false;
    }

    // 11. Doa Atas Persembahan Content Slide
    else if (
      slideTextUpper.includes("DOA ATAS PERSEMBAHAN") ||
      slideTextUpper.includes("ATAS PERSEMBAHAN")
    ) {
      const originalSlideXml = slideXml;
      const persembahanChunks = splitTextIntoSlideChunks(
        mappedData.doaAtasPersembahan || "",
        240,
      );

      if (persembahanChunks.length === 1) {
        const filledSlide77Xml = replacePlaceholdersInXml(originalSlideXml, [
          persembahanChunks[0] || "",
        ]);
        zip.file(slide.fileName, filledSlide77Xml);
      } else {
        let slide1Xml = stripClosingParagraphs(originalSlideXml);
        slide1Xml = replacePlaceholdersInXml(slide1Xml, [persembahanChunks[0]]);
        zip.file(slide.fileName, slide1Xml);

        let lastRId = slide.rId;
        for (let i = 1; i < persembahanChunks.length; i++) {
          const isLast = i === persembahanChunks.length - 1;
          let dupSlideXml = originalSlideXml;

          if (!isLast) {
            dupSlideXml = stripClosingParagraphs(dupSlideXml);
          }

          const filledChunkXml = replacePlaceholdersInXml(dupSlideXml, [
            persembahanChunks[i],
          ]);
          const { newRId } = await duplicateSlide(
            slide.fileName,
            lastRId,
            filledChunkXml,
          );
          lastRId = newRId;
        }
      }
    }

    // 12. Doa Sesudah Komuni Content Slide (Slide 92: right after DOA SESUDAH KOMUNI title slide)
    else if (
      (hasPassedDoaKomuniTitle && !hasHandledDoaKomuni) ||
      slideTextUpper.includes("DOA SESUDAH KOMUNI") ||
      slideTextUpper.includes("SESUDAH KOMUNI")
    ) {
      const originalSlideXml = slideXml;
      const komuniChunks = splitTextIntoSlideChunks(
        mappedData.doaSesudahKomuni || "",
        280,
      );

      if (komuniChunks.length === 1) {
        const filledSlide97Xml = replacePlaceholdersInXml(originalSlideXml, [
          komuniChunks[0] || "",
        ]);
        zip.file(slide.fileName, filledSlide97Xml);
      } else {
        const slide1Xml = replacePlaceholdersInXml(originalSlideXml, [
          komuniChunks[0],
        ]);
        zip.file(slide.fileName, slide1Xml);

        let lastRId = slide.rId;
        for (let i = 1; i < komuniChunks.length; i++) {
          let dupSlideXml = originalSlideXml;
          dupSlideXml = stripIntroParagraphs(dupSlideXml);

          const filledChunkXml = replacePlaceholdersInXml(dupSlideXml, [
            komuniChunks[i],
          ]);
          const { newRId } = await duplicateSlide(
            slide.fileName,
            lastRId,
            filledChunkXml,
          );
          lastRId = newRId;
        }
      }
      hasHandledDoaKomuni = true;
      hasPassedDoaKomuniTitle = false;
    }
  }

  // ponytail: re-zip into nodebuffer binary
  const pptxBuffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return pptxBuffer;
}
