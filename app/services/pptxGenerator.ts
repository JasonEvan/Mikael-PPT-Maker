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

// ponytail: update speaker notes in ppt/notesSlides/notesSlideX.xml
function updateNotesSlideXml(notesXml: string, notesText: string): string {
  if (!notesText) return notesXml;
  const escapedNotes = escapeXml(notesText);

  // Preserve (jangan diblank) reminder if present
  if (/(\(\s*jangan diblank\s*\))/i.test(notesXml)) {
    return notesXml.replace(
      /(\(\s*jangan diblank\s*\))/i,
      `$1\n\n${escapedNotes}`,
    );
  } else {
    return notesXml.replace(
      /(<\/p:txBody>)/,
      `<a:p><a:r><a:t>${escapedNotes}</a:t></a:r></a:p>$1`,
    );
  }
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
  const slideOrder: { index: number; rId: string; fileName: string }[] = [];
  let idx = 1;

  for (const m of sldMatches) {
    const rId = m[1];
    const target = rIdMap[rId];
    if (target) {
      const fileName = target.startsWith("ppt/") ? target : `ppt/${target}`;
      slideOrder.push({ index: idx, rId, fileName });
      idx++;
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

  // ponytail: helper to duplicate a slide cleanly and insert it IMMEDIATELY AFTER targetRId
  const duplicateSlide = async (
    sourceSlideNum: number,
    targetRId: string,
    customXml?: string,
  ): Promise<{ newSlideNum: number; newFileName: string; newRId: string }> => {
    maxSlideNum++;
    maxSldId++;
    maxRId++;
    const newSlideNum = maxSlideNum;
    const newFileName = `ppt/slides/slide${newSlideNum}.xml`;
    const sourceFileName = `ppt/slides/slide${sourceSlideNum}.xml`;
    const newRId = `rId${maxRId}`;

    const xmlToWrite =
      customXml || (await zip.file(sourceFileName)?.async("text"));
    if (xmlToWrite) zip.file(newFileName, xmlToWrite);

    const sourceRels = `ppt/slides/_rels/slide${sourceSlideNum}.xml.rels`;
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

  // ponytail: process each slide placeholder according to MVP scope
  for (const slide of slideOrder) {
    const sNum = slide.index;

    // Skip slides 74, 95, 117 (song lyrics) & slides 102–112 (announcements)
    if (
      sNum === 74 ||
      sNum === 95 ||
      sNum === 117 ||
      (sNum >= 102 && sNum <= 112)
    ) {
      continue;
    }

    const slideZip = zip.file(slide.fileName);
    if (!slideZip) continue;

    let slideXml = await slideZip.async("text");

    // 1. Seruan Tobat (Slides 22, 23, 24)
    if (sNum === 22) {
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.seruanTobat?.[0] || "",
      ]);
      zip.file(slide.fileName, slideXml);
    } else if (sNum === 23) {
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.seruanTobat?.[1] || "",
      ]);
      zip.file(slide.fileName, slideXml);
    } else if (sNum === 24) {
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.seruanTobat?.[2] || "",
      ]);
      zip.file(slide.fileName, slideXml);
    }

    // 2. Doa Kolekta (Slide 33)
    else if (sNum === 33) {
      const originalSlideXml = slideXml;
      const kolektaChunks = splitTextIntoSlideChunks(
        mappedData.doaKolekta || "",
        260,
      );

      if (kolektaChunks.length === 1) {
        // Single slide: keep template intact (has intro and closing)
        const filledSlide33Xml = replacePlaceholdersInXml(originalSlideXml, [
          kolektaChunks[0] || "",
        ]);
        zip.file(slide.fileName, filledSlide33Xml);
      } else {
        // Multi-slide duplication:
        // Slide 1: Keep intro paragraphs before (teks), strip closing paragraphs after (teks)
        let slide1Xml = stripClosingParagraphs(originalSlideXml);
        slide1Xml = replacePlaceholdersInXml(slide1Xml, [kolektaChunks[0]]);
        zip.file(slide.fileName, slide1Xml);

        let lastRId = slide.rId;
        for (let i = 1; i < kolektaChunks.length; i++) {
          const isLast = i === kolektaChunks.length - 1;
          let dupSlideXml = originalSlideXml;

          // Structurally strip intro paragraphs before (teks) on duplicated slides
          dupSlideXml = stripIntroParagraphs(dupSlideXml);

          // Structurally strip closing paragraphs on middle slides (keep closing on final slide)
          if (!isLast) {
            dupSlideXml = stripClosingParagraphs(dupSlideXml);
          }

          const filledChunkXml = replacePlaceholdersInXml(dupSlideXml, [
            kolektaChunks[i],
          ]);
          const { newRId } = await duplicateSlide(33, lastRId, filledChunkXml);
          lastRId = newRId;
        }
      }
    }

    // 3. Bacaan Pertama (Slide 35 + notesSlide2.xml)
    else if (sNum === 35) {
      const epigraph = mappedData.bacaanPertama?.epigraph || "";
      const citation = mappedData.bacaanPertama?.citation || "";
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

    // 4. Mazmur Tanggapan (Slides 39, 41, 43)
    else if (sNum === 39) {
      const verses = mappedData.mazmurTanggapan?.verses || [];
      slideXml = replacePlaceholdersInXml(slideXml, [verses[0] || ""]);
      zip.file(slide.fileName, slideXml);
    } else if (sNum === 41) {
      const verses = mappedData.mazmurTanggapan?.verses || [];
      slideXml = replacePlaceholdersInXml(slideXml, [verses[1] || ""]);
      zip.file(slide.fileName, slideXml);
    } else if (sNum === 43) {
      const originalSlideXml = slideXml;
      const verses = mappedData.mazmurTanggapan?.verses || [];
      const filledSlide43Xml = replacePlaceholdersInXml(originalSlideXml, [
        verses[2] || "",
      ]);
      zip.file(slide.fileName, filledSlide43Xml);

      // Duplicate slide 43 right after if there are > 3 verses
      if (verses.length > 3) {
        let lastRId = slide.rId;
        for (let i = 3; i < verses.length; i++) {
          const filledVerseXml = replacePlaceholdersInXml(originalSlideXml, [
            verses[i],
          ]);
          const { newRId } = await duplicateSlide(43, lastRId, filledVerseXml);
          lastRId = newRId;
        }
      }
    }

    // 5. Bacaan Kedua (Slide 46 + notesSlide3.xml)
    else if (sNum === 46) {
      const epigraph = mappedData.bacaanKedua?.epigraph || "";
      const citation = mappedData.bacaanKedua?.citation || "";
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

    // 6. Bacaan Injil - Gospel Citation (Slide 50)
    else if (sNum === 50) {
      const gospelName = mappedData.bacaanInjil?.gospelName || "";
      slideXml = replacePlaceholdersInXml(slideXml, [gospelName]);
      zip.file(slide.fileName, slideXml);
    }

    // 7. Bacaan Injil - Epigraph & Citation (Slide 51 + notesSlide6.xml)
    else if (sNum === 51) {
      const epigraph = mappedData.bacaanInjil?.epigraph || "";
      const citation = mappedData.bacaanInjil?.citation || "";
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

    // 8. Doa Umat Intro (Slide 68)
    else if (sNum === 68) {
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.doaUmat?.intro || "",
      ]);
      zip.file(slide.fileName, slideXml);
    }

    // 9. Doa Umat Intentions (Slide 69 + Duplications right after if > 1 intentions)
    else if (sNum === 69) {
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
            69,
            lastRId,
            filledIntentionXml,
          );
          lastRId = newRId;
        }
      }
    }

    // 10. Doa Umat Closing (Slide 71)
    else if (sNum === 71) {
      slideXml = replacePlaceholdersInXml(slideXml, [
        mappedData.doaUmat?.closing || "",
      ]);
      zip.file(slide.fileName, slideXml);
    }

    // 11. Doa Atas Persembahan (Slide 77)
    else if (sNum === 77) {
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
          const { newRId } = await duplicateSlide(77, lastRId, filledChunkXml);
          lastRId = newRId;
        }
      }
    }

    // 12. Doa Sesudah Komuni (Slide 97)
    else if (sNum === 97) {
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
          const { newRId } = await duplicateSlide(97, lastRId, filledChunkXml);
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
