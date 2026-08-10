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

// ponytail: helper to strip leading L / I / L. / I. / L- / L.  / LBacaan prefix from reading citations
function cleanCitation(citation?: string): string {
  if (!citation) return "";
  let cleaned = citation.trim();
  cleaned = cleaned.replace(/^[LI][.:\s-]*/i, "").trim();
  return cleaned;
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
  maxCharsPerSlide: number = 260,
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

    // Look for last punctuation boundary (. , ! ? ;) in upper portion of slice
    const periodIndex = slice.lastIndexOf(". ");
    const commaIndex = slice.lastIndexOf(", ");
    const exclamIndex = slice.lastIndexOf("! ");
    const questIndex = slice.lastIndexOf("? ");
    const semiIndex = slice.lastIndexOf("; ");
    const lastPunctuationEnd = Math.max(
      periodIndex,
      commaIndex,
      exclamIndex,
      questIndex,
      semiIndex,
    );

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
  for (const m of relMatches) {
    rIdMap[m[1]] = m[2];
  }

  const sldMatches = presentationXml.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g);
  const slideOrder: { rId: string; fileName: string }[] = [];

  for (const m of sldMatches) {
    const rId = m[1];
    const target = rIdMap[rId];
    if (target) {
      const fileName = target.startsWith("ppt/") ? target : `ppt/${target}`;
      slideOrder.push({ rId, fileName });
    }
  }

  // Track max numeric rId in presentation.xml.rels for valid OpenXML IDs (e.g. rId125)
  const rIdNums = [...presentationRels.matchAll(/Id="rId(\d+)"/g)].map((m) =>
    parseInt(m[1], 10),
  );
  let maxRId = Math.max(1, ...rIdNums);

  // Find max slide file number and max sldId for duplication
  const allSlideFiles = Object.keys(zip.files).filter((f) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(f),
  );
  let maxSlideNum = 117;
  allSlideFiles.forEach((f) => {
    const num = parseInt(f.match(/\d+/)?.[0] || "0", 10);
    if (num > maxSlideNum) maxSlideNum = num;
  });

  const sldIdNums = [...presentationXml.matchAll(/<p:sldId id="(\d+)"/g)].map(
    (m) => parseInt(m[1], 10),
  );
  let maxSldId = Math.max(256, ...sldIdNums);

  // ponytail: helper to duplicate a slide cleanly by exact sourceFileName, inserting IMMEDIATELY AFTER targetRId
  const duplicateSlide = async (
    sourceFileName: string,
    targetRId: string,
    customXml?: string,
  ): Promise<{ newSlideNum: number; newFileName: string; newRId: string }> => {
    maxSlideNum++;
    maxSldId++;
    maxRId++;
    const newSlideNum = maxSlideNum;
    const newFileName = `ppt/slides/slide${newSlideNum}.xml`;
    const newRId = `rId${maxRId}`;

    const xmlToWrite =
      customXml || (await zip.file(sourceFileName)?.async("text"));
    if (xmlToWrite) zip.file(newFileName, xmlToWrite);

    // Compute sourceRels from sourceFileName
    const relsFileName = sourceFileName
      .replace(/^ppt\/slides\//, "")
      .replace(/\.xml$/, ".xml.rels");
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
      const newSldTag = `<p:sldId id="${maxSldId}" r:id="${newRId}"/>`;
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
  let mazmurCount = 0;
  let hasPassedBacaanPertama = false;
  let hasPassedBacaanKedua = false;

  // ponytail: process each slide placeholder using Content Slide filtering + Content Pattern Anchors (ZERO slide index dependency)
  for (const slide of slideOrder) {
    const slideZip = zip.file(slide.fileName);
    if (!slideZip) continue;

    let slideXml = await slideZip.async("text");

    // Only process content slides that actually contain a (teks) placeholder
    const hasPlaceholder = slideHasPlaceholder(slideXml);
    if (!hasPlaceholder) continue;

    // Extract visible slide text to anchor slide type by internal text/placeholder content
    const pMatches = slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
    const slideTextUpper = pMatches
      .map((p) => {
        const tMatches = p.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
        return tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");
      })
      .join(" ")
      .toUpperCase();

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
      hasPassedBacaanPertama = true;
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

    // 4. Mazmur Tanggapan Content Slides (occurs between Bacaan Pertama and Bacaan Kedua)
    else if (
      hasPassedBacaanPertama &&
      !hasPassedBacaanKedua &&
      !slideTextUpper.includes("BACAAN KEDUA")
    ) {
      const verses = mappedData.mazmurTanggapan?.verses || [];
      if (mazmurCount < 2) {
        slideXml = replacePlaceholdersInXml(slideXml, [
          verses[mazmurCount] || "",
        ]);
        zip.file(slide.fileName, slideXml);
        mazmurCount++;
      } else if (mazmurCount === 2) {
        const originalSlideXml = slideXml;
        const filledSlide43Xml = replacePlaceholdersInXml(originalSlideXml, [
          verses[2] || "",
        ]);
        zip.file(slide.fileName, filledSlide43Xml);

        if (verses.length > 3) {
          let lastRId = slide.rId;
          for (let i = 3; i < verses.length; i++) {
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
        mazmurCount++;
      }
    }

    // 5. Bacaan Kedua Content Slide (+ notesSlide3.xml)
    else if (
      slideTextUpper.includes("BACAAN KEDUA") ||
      slideTextUpper.includes("BACAAN II ")
    ) {
      hasPassedBacaanKedua = true;
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

    // 8. Doa Umat Intentions Content Slide (Slide with 2 placeholders)
    else if (
      pMatches.length >= 2 &&
      pMatches.filter((p) => /\(\s*teks\s*\)/i.test(p)).length >= 2
    ) {
      const originalSlideXml = slideXml;
      const intentions = mappedData.doaUmat?.intentions || [];
      const title0 = intentions[0]?.title || "";
      const body0 = intentions[0]?.body || "";
      const filledSlide69Xml = replacePlaceholdersInXml(originalSlideXml, [
        title0,
        body0,
      ]);
      zip.file(slide.fileName, filledSlide69Xml);

      if (intentions.length > 1) {
        let lastRId = slide.rId;
        for (let i = 1; i < intentions.length; i++) {
          const filledIntentionXml = replacePlaceholdersInXml(
            originalSlideXml,
            [intentions[i].title, intentions[i].body],
          );
          const { newRId } = await duplicateSlide(
            slide.fileName,
            lastRId,
            filledIntentionXml,
          );
          lastRId = newRId;
        }
      }
    }

    // 9. Doa Umat Intro
    else if (
      slideTextUpper.includes("DOA UMAT") &&
      !slideTextUpper.includes("AMIN")
    ) {
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.doaUmat?.intro || "",
      ]);
      zip.file(slide.fileName, slideXml);
    }

    // 10. Doa Umat Closing
    else if (
      slideTextUpper.includes("DOA UMAT") &&
      slideTextUpper.includes("AMIN")
    ) {
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.doaUmat?.closing || "",
      ]);
      zip.file(slide.fileName, slideXml);
    }

    // 11. Doa Atas Persembahan Content Slide
    else if (
      slideTextUpper.includes("DOA ATAS PERSEMBAHAN") ||
      slideTextUpper.includes("ATAS PERSEMBAHAN")
    ) {
      const originalSlideXml = slideXml;
      const persembahanChunks = splitTextIntoSlideChunks(
        mappedData.doaAtasPersembahan || "",
        260,
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

    // 12. Doa Sesudah Komuni Content Slide
    else if (
      slideTextUpper.includes("DOA SESUDAH KOMUNI") ||
      slideTextUpper.includes("SESUDAH KOMUNI")
    ) {
      const originalSlideXml = slideXml;
      const komuniChunks = splitTextIntoSlideChunks(
        mappedData.doaSesudahKomuni || "",
        260,
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
