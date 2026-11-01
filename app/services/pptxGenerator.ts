import fs from "fs";
import path from "path";
import JSZip from "jszip";
import { MappedLiturgyData } from "../utils/docxParser";
import { createPptxTextLayout, PptxTextReplacement } from "./pptxTextLayout";

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

// ponytail: format epigraph wrapped cleanly in quotation marks “...”
function formatEpigraph(epigraph?: string): string {
  if (!epigraph || !epigraph.trim()) return "";
  let cleaned = epigraph.trim();

  while (/^[*\s"“'”‘’`«»]/.test(cleaned) || /[*\s"“'”‘’`«»]$/.test(cleaned)) {
    cleaned = cleaned.replace(/^[*\s"“'”‘’`«»]+|[*\s"“'”‘’`«»]+$/g, "").trim();
  }

  if (!cleaned) return "";
  return `“${cleaned}”`;
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

function cleanGospelCitation(
  citation?: string,
  omitGospelPrefix = false,
): string {
  const cleaned = cleanCitation(citation);
  return omitGospelPrefix
    ? cleaned.replace(/^Injil\s+Suci\s+menurut\s+/i, "").trim()
    : cleaned;
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

// OBS has one intention token: the initial sentence/title is emphasized inline,
// while its response follows the entire body in that same measured shape.
function obsIntentionReplacement(
  item: NonNullable<MappedLiturgyData["doaUmat"]>["intentions"][number],
): PptxTextReplacement {
  const body = [item.title, item.body].filter(Boolean).join("\n");
  const lineBreak = body.indexOf("\n");
  const sentence = body.match(/[.!?…](?=\s|$)/);
  const sentenceEnd = sentence ? sentence.index! + sentence[0].length : -1;
  const splitAt =
    sentenceEnd >= 0 && (lineBreak < 0 || sentenceEnd <= lineBreak)
      ? sentenceEnd
      : lineBreak >= 0
        ? lineBreak
        : body.length;
  return {
    paragraphs: [
      {
        runs: [
          { text: body.slice(0, splitAt), bold: true },
          { text: body.slice(splitAt), bold: false },
        ],
      },
      ...(item.response
        ? [
            {
              text: `U\t${item.response}`,
              bold: false,
              policy: "last" as const,
            },
          ]
        : []),
    ],
  };
}

export interface PptxGeneratorOptions {
  usesPlaceholderOnlyInjilContentSlide?: boolean;
  omitsGospelCitationPrefix?: boolean;
  usesSinglePlaceholderDoaUmatIntentions?: boolean;
  usesPlaceholderOnlySeruanTobatSlides?: boolean;
}

// ponytail: generate filled PPTX from template and MappedLiturgyData
export async function generatePPTXBufferForTemplate(
  mappedData: MappedLiturgyData,
  templateFileName: string,
  options: PptxGeneratorOptions = {},
): Promise<Buffer> {
  const templatePath = path.join(
    process.cwd(),
    "app/template",
    templateFileName,
  );
  if (!fs.existsSync(templatePath)) {
    throw new Error(
      `Template PowerPoint (${templateFileName}) tidak ditemukan.`,
    );
  }

  const templateBuffer = await fs.promises.readFile(templatePath);
  const zip = await JSZip.loadAsync(templateBuffer);
  const textLayout = await createPptxTextLayout(zip, {
    fontProfile: "indonesian",
  });

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

  // Resolve a slide's notes through its own relationship so each template can
  // use its natural notes-slide numbering.
  const updateSlideNotes = async (slideFileName: string, notesText: string) => {
    if (!notesText) return;

    const slideNum = slideFileName.match(/slide(\d+)\.xml$/)?.[1];
    if (!slideNum) return;

    const relsFile = zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`);
    if (!relsFile) return;

    const relsXml = await relsFile.async("text");
    const target = relsXml.match(
      /<Relationship [^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesSlide"[^>]*Target="([^\"]+)"/i,
    )?.[1];
    if (!target) return;

    const notesFileName = path.posix.normalize(
      path.posix.join("ppt/slides", target),
    );
    const notesFile = zip.file(notesFileName);
    if (!notesFile) return;

    const notesXml = await notesFile.async("text");
    zip.file(notesFileName, updateNotesSlideXml(notesXml, notesText));
  };

  // Tracking section states dynamically across content slides
  let seruanCount = 0;
  let hasPassedSeruanTobatTitle = false;
  let hasHandledMazmur = false;
  let hasPassedBacaanPertama = false;
  let hasPassedBacaanKedua = false;
  let hasPendingInjilContent = false;
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

    const slideXml = await slideZip.async("text");

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

    // OBS Hijau places each Seruan Tobat invocation on a placeholder-only
    // slide, followed by a separate, fixed congregation-response slide.
    if (
      options.usesPlaceholderOnlySeruanTobatSlides &&
      slideTextUpper.includes("SERUAN TOBAT")
    ) {
      hasPassedSeruanTobatTitle = true;
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

    // In OBS Hijau, the content slide after this citation contains only the
    // two placeholders, so it cannot be identified from its own slide text.
    if (
      options.usesPlaceholderOnlyInjilContentSlide &&
      slideTextUpper.includes("INILAH INJIL SUCI MENURUT")
    ) {
      hasPendingInjilContent = true;
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
    let sequences: PptxTextReplacement[][] | undefined;
    let notesText = "";

    if (hasPassedSeruanTobatTitle && seruanCount < 3) {
      sequences = [[mappedData.seruanTobat?.[seruanCount] || ""]];
      seruanCount++;
      if (seruanCount === 3) hasPassedSeruanTobatTitle = false;
    } else if (slideTextUpper.includes("KASIHANILAH KAMI")) {
      const idx =
        seruanCount < (mappedData.seruanTobat?.length || 0) ? seruanCount : 0;
      sequences = [[mappedData.seruanTobat?.[idx] || ""]];
      seruanCount++;
    } else if (
      slideTextUpper.includes("MARILAH KITA BERDOA") &&
      !hasPassedBacaanPertama
    ) {
      sequences = [[mappedData.doaKolekta || ""]];
    } else if (
      slideTextUpper.includes("BACAAN PERTAMA") ||
      slideTextUpper.includes("BACAAN I ")
    ) {
      sequences = [
        [
          formatEpigraph(mappedData.bacaanPertama?.epigraph),
          cleanCitation(mappedData.bacaanPertama?.citation),
        ],
      ];
      notesText = mappedData.bacaanPertama?.notesBody || "";
    } else if (
      (slideTextUpper.includes("MAZMUR") ||
        slideTextUpper.includes("TANGGAPAN") ||
        (hasPassedBacaanPertama && !hasPassedBacaanKedua)) &&
      !hasHandledMazmur &&
      !slideTextUpper.includes("BACAAN KEDUA") &&
      !slideTextUpper.includes("BACAAN PERTAMA")
    ) {
      // A verse is a separate sequence, even when another verse would fit.
      // With zero verses the original placeholder is intentionally untouched.
      sequences = (mappedData.mazmurTanggapan?.verses || []).map((verse) => [
        verse,
      ]);
      hasHandledMazmur = true;
    } else if (
      slideTextUpper.includes("BACAAN KEDUA") ||
      slideTextUpper.includes("BACAAN II ")
    ) {
      sequences = [
        [
          formatEpigraph(mappedData.bacaanKedua?.epigraph),
          cleanCitation(mappedData.bacaanKedua?.citation),
        ],
      ];
      notesText = mappedData.bacaanKedua?.notesBody || "";
    } else if (slideTextUpper.includes("INILAH INJIL SUCI MENURUT")) {
      sequences = [[mappedData.bacaanInjil?.gospelName || ""]];
    } else if (
      (hasPendingInjilContent && getPlaceholderCount(slideXml) === 2) ||
      (slideTextUpper.includes("INJIL") &&
        !slideTextUpper.includes("PENGANTAR"))
    ) {
      sequences = [
        [
          formatEpigraph(mappedData.bacaanInjil?.epigraph),
          cleanGospelCitation(
            mappedData.bacaanInjil?.citation,
            options.omitsGospelCitationPrefix,
          ),
        ],
      ];
      notesText = mappedData.bacaanInjil?.notesBody || "";
      hasPendingInjilContent = false;
    } else if (hasPassedDoaUmatTitle && !hasHandledDoaUmatIntro) {
      if (getPlaceholderCount(slideXml) === 1) {
        sequences = [[mappedData.doaUmat?.intro || ""]];
        hasHandledDoaUmatIntro = true;
      }
    } else if (
      hasPassedDoaUmatTitle &&
      hasHandledDoaUmatIntro &&
      !hasHandledDoaUmatIntentions
    ) {
      const count = getPlaceholderCount(slideXml);
      const intentions = mappedData.doaUmat?.intentions || [];
      if (options.usesSinglePlaceholderDoaUmatIntentions && count === 1) {
        sequences = intentions.length
          ? intentions.map((item) => [obsIntentionReplacement(item)])
          : [[""]];
        hasHandledDoaUmatIntentions = true;
      } else if (count >= 2) {
        sequences = intentions.length
          ? intentions.map<PptxTextReplacement[]>((item) => [
              {
                paragraphs: [
                  ...(item.title ? [{ text: item.title, bold: true }] : []),
                  ...(item.body ? [{ text: item.body, bold: false }] : []),
                ],
              },
              {
                paragraphs: item.response
                  ? [{ text: `U\t${item.response}` }]
                  : [],
                policy: "last",
              },
            ])
          : [["", ""]];
        hasHandledDoaUmatIntentions = true;
      }
    } else if (
      hasPassedDoaUmatTitle &&
      hasHandledDoaUmatIntentions &&
      !hasHandledDoaUmatOutro
    ) {
      sequences = [[mappedData.doaUmat?.closing || ""]];
      hasHandledDoaUmatOutro = true;
      hasPassedDoaUmatTitle = false;
    } else if (
      slideTextUpper.includes("DOA ATAS PERSEMBAHAN") ||
      slideTextUpper.includes("ATAS PERSEMBAHAN")
    ) {
      sequences = [[mappedData.doaAtasPersembahan || ""]];
    } else if (
      (hasPassedDoaKomuniTitle && !hasHandledDoaKomuni) ||
      slideTextUpper.includes("DOA SESUDAH KOMUNI") ||
      slideTextUpper.includes("SESUDAH KOMUNI")
    ) {
      sequences = [[mappedData.doaSesudahKomuni || ""]];
      hasHandledDoaKomuni = true;
      hasPassedDoaKomuniTitle = false;
    }

    if (sequences?.length) {
      // Fixed prompts and responses share the measured shape budget with
      // its tokens. Keep their exact template text/style, once at each end.
      const paragraphTexts = pMatches.map((paragraph) =>
        (paragraph.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [])
          .map((run) => run.replace(/<[^>]+>/g, ""))
          .join(""),
      );
      const firstToken = paragraphTexts.findIndex((text) =>
        /\(\s*teks\s*\)/i.test(text),
      );
      const lastToken = paragraphTexts.reduce(
        (last, text, index) => (/\(\s*teks\s*\)/i.test(text) ? index : last),
        -1,
      );
      const firstOnlyTexts = paragraphTexts
        .slice(0, firstToken)
        .filter((text) => text.trim());
      const lastOnlyTexts = paragraphTexts
        .slice(lastToken + 1)
        .filter((text) => text.trim());
      let wroteFirstPage = false;
      let lastRId = slide.rId;
      for (const replacements of sequences) {
        const pages = textLayout.paginate(
          slide.fileName,
          slideXml,
          "teks",
          replacements,
          { firstOnlyTexts, lastOnlyTexts, preserveTemplateSpacers: true },
        );
        for (const pageXml of pages) {
          if (!wroteFirstPage) {
            zip.file(slide.fileName, pageXml);
            wroteFirstPage = true;
          } else {
            const duplicate = await duplicateSlide(
              slide.fileName,
              lastRId,
              pageXml,
            );
            lastRId = duplicate.newRId;
          }
        }
      }
      await updateSlideNotes(slide.fileName, notesText);
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

// Default generator retained for the original weekly template.
export function generatePPTXBuffer(
  mappedData: MappedLiturgyData,
): Promise<Buffer> {
  return generatePPTXBufferForTemplate(mappedData, "Template_Mingguan.pptx");
}
