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
function replacePlaceholdersInXml(slideXml: string, replacements: string[]): string {
  let replacementIndex = 0;

  return slideXml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (pBlock) => {
    // Extract text from all <a:t> tags in this paragraph
    const tMatches = pBlock.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
    const fullText = tMatches.map((t) => t.replace(/<[^>]+>/g, "")).join("");

    if (/\(\s*teks\s*\)/i.test(fullText)) {
      const valToInsert = replacementIndex < replacements.length ? replacements[replacementIndex] : "";
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

// ponytail: update speaker notes in ppt/notesSlides/notesSlideX.xml
function updateNotesSlideXml(notesXml: string, notesText: string): string {
  if (!notesText) return notesXml;
  const escapedNotes = escapeXml(notesText);

  // Preserve (jangan diblank) reminder if present
  if (/(\(\s*jangan diblank\s*\))/i.test(notesXml)) {
    return notesXml.replace(/(\(\s*jangan diblank\s*\))/i, `$1\n\n${escapedNotes}`);
  } else {
    return notesXml.replace(/(<\/p:txBody>)/, `<a:p><a:r><a:t>${escapedNotes}</a:t></a:r></a:p>$1`);
  }
}

// ponytail: generate filled PPTX from template and MappedLiturgyData
export async function generatePPTXBuffer(mappedData: MappedLiturgyData): Promise<Buffer> {
  const templatePath = path.join(process.cwd(), "app/template/Template_Mingguan.pptx");
  if (!fs.existsSync(templatePath)) {
    throw new Error("Template PowerPoint (Template_Mingguan.pptx) tidak ditemukan.");
  }

  const templateBuffer = await fs.promises.readFile(templatePath);
  const zip = await JSZip.loadAsync(templateBuffer);

  // 1. Read presentation structure and slide mapping
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const presentationRels = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");

  if (!presentationXml || !presentationRels) {
    throw new Error("Struktur berkas PowerPoint tidak valid.");
  }

  const rIdMap: Record<string, string> = {};
  const relMatches = presentationRels.matchAll(/<Relationship [^>]*Id="([^"]+)" [^>]*Target="([^"]+)"/g);
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

  // Find max slide file number and max sldId for duplication
  const allSlideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  let maxSlideNum = 117;
  allSlideFiles.forEach((f) => {
    const num = parseInt(f.match(/\d+/)?.[0] || "0", 10);
    if (num > maxSlideNum) maxSlideNum = num;
  });

  const sldIdNums = [...presentationXml.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => parseInt(m[1], 10));
  let maxSldId = Math.max(256, ...sldIdNums);

  // ponytail: helper to duplicate a slide in the PPTX archive
  const duplicateSlide = async (sourceSlideNum: number): Promise<{ newSlideNum: number; newFileName: string }> => {
    maxSlideNum++;
    maxSldId++;
    const newSlideNum = maxSlideNum;
    const newFileName = `ppt/slides/slide${newSlideNum}.xml`;
    const sourceFileName = `ppt/slides/slide${sourceSlideNum}.xml`;

    const sourceXml = await zip.file(sourceFileName)?.async("text");
    if (sourceXml) zip.file(newFileName, sourceXml);

    const sourceRels = `ppt/slides/_rels/slide${sourceSlideNum}.xml.rels`;
    const sourceRelsFile = zip.file(sourceRels);
    if (sourceRelsFile) {
      const relsXml = await sourceRelsFile.async("text");
      zip.file(`ppt/slides/_rels/slide${newSlideNum}.xml.rels`, relsXml);
    }

    // Update [Content_Types].xml
    let contentTypes = await zip.file("[Content_Types].xml")?.async("text");
    if (contentTypes) {
      const overrideTag = `<Override PartName="/ppt/slides/slide${newSlideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
      contentTypes = contentTypes.replace("</Types>", `${overrideTag}</Types>`);
      zip.file("[Content_Types].xml", contentTypes);
    }

    // Update ppt/_rels/presentation.xml.rels
    let presRels = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");
    const newRId = `rIdSld${newSlideNum}`;
    if (presRels) {
      const relTag = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNum}.xml"/>`;
      presRels = presRels.replace("</Relationships>", `${relTag}</Relationships>`);
      zip.file("ppt/_rels/presentation.xml.rels", presRels);
    }

    // Update ppt/presentation.xml
    let presXml = await zip.file("ppt/presentation.xml")?.async("text");
    if (presXml) {
      const newSldTag = `<p:sldId id="${maxSldId}" r:id="${newRId}"/>`;
      presXml = presXml.replace("</p:sldIdLst>", `${newSldTag}</p:sldIdLst>`);
      zip.file("ppt/presentation.xml", presXml);
    }

    return { newSlideNum, newFileName };
  };

  // ponytail: process each slide placeholder according to MVP scope
  for (const slide of slideOrder) {
    const sNum = slide.index;

    // Skip slides 74, 95, 117 (song lyrics) & slides 102–112 (announcements)
    if (sNum === 74 || sNum === 95 || sNum === 117 || (sNum >= 102 && sNum <= 112)) {
      continue;
    }

    const slideZip = zip.file(slide.fileName);
    if (!slideZip) continue;

    let slideXml = await slideZip.async("text");

    // 1. Seruan Tobat (Slides 22, 23, 24)
    if (sNum === 22) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.seruanTobat?.[0] || ""]);
      zip.file(slide.fileName, slideXml);
    } else if (sNum === 23) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.seruanTobat?.[1] || ""]);
      zip.file(slide.fileName, slideXml);
    } else if (sNum === 24) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.seruanTobat?.[2] || ""]);
      zip.file(slide.fileName, slideXml);
    }

    // 2. Doa Kolekta (Slide 33)
    else if (sNum === 33) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.doaKolekta || ""]);
      zip.file(slide.fileName, slideXml);
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
          const updatedNotes = updateNotesSlideXml(notesXml, mappedData.bacaanPertama.notesBody);
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
      const verses = mappedData.mazmurTanggapan?.verses || [];
      slideXml = replacePlaceholdersInXml(slideXml, [verses[2] || ""]);
      zip.file(slide.fileName, slideXml);

      // Duplicate slide 43 if there are > 3 verses
      if (verses.length > 3) {
        for (let i = 3; i < verses.length; i++) {
          const { newFileName } = await duplicateSlide(43);
          const dupXml = await zip.file(newFileName)?.async("text");
          if (dupXml) {
            const filledDupXml = replacePlaceholdersInXml(dupXml, [verses[i]]);
            zip.file(newFileName, filledDupXml);
          }
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
          const updatedNotes = updateNotesSlideXml(notesXml, mappedData.bacaanKedua.notesBody);
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
          const updatedNotes = updateNotesSlideXml(notesXml, mappedData.bacaanInjil.notesBody);
          zip.file("ppt/notesSlides/notesSlide6.xml", updatedNotes);
        }
      }
    }

    // 8. Doa Umat Intro (Slide 68)
    else if (sNum === 68) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.doaUmat?.intro || ""]);
      zip.file(slide.fileName, slideXml);
    }

    // 9. Doa Umat Intentions (Slide 69 + Duplications if > 1 intentions)
    else if (sNum === 69) {
      const intentions = mappedData.doaUmat?.intentions || [];
      const title0 = intentions[0]?.title || "";
      const body0 = intentions[0]?.body || "";
      slideXml = replacePlaceholdersInXml(slideXml, [title0, body0]);
      zip.file(slide.fileName, slideXml);

      if (intentions.length > 1) {
        for (let i = 1; i < intentions.length; i++) {
          const { newFileName } = await duplicateSlide(69);
          const dupXml = await zip.file(newFileName)?.async("text");
          if (dupXml) {
            const filledDupXml = replacePlaceholdersInXml(dupXml, [intentions[i].title, intentions[i].body]);
            zip.file(newFileName, filledDupXml);
          }
        }
      }
    }

    // 10. Doa Umat Closing (Slide 71)
    else if (sNum === 71) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.doaUmat?.closing || ""]);
      zip.file(slide.fileName, slideXml);
    }

    // 11. Doa Atas Persembahan (Slide 77)
    else if (sNum === 77) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.doaAtasPersembahan || ""]);
      zip.file(slide.fileName, slideXml);
    }

    // 12. Doa Sesudah Komuni (Slide 97)
    else if (sNum === 97) {
      slideXml = replacePlaceholdersInXml(slideXml, [mappedData.doaSesudahKomuni || ""]);
      zip.file(slide.fileName, slideXml);
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
