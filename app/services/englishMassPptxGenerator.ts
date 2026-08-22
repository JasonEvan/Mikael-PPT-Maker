import fs from "fs";
import path from "path";
import JSZip from "jszip";
import {
  EnglishMassBlock,
  EnglishMassParseResult,
} from "../utils/englishMassParser";

const ENGLISH_TEMPLATE_FILE = "EM_Template.pptx";

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getParagraphText(paragraphXml: string): string {
  return (paragraphXml.match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g) ?? [])
    .map((textXml) => textXml.replace(/<[^>]+>/g, ""))
    .join("");
}

function getSlideText(slideXml: string): string {
  return (slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? [])
    .map(getParagraphText)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function countTokenPlaceholders(slideXml: string, token: string): number {
  const tokenPattern = new RegExp(`\\(\\s*${token}\\s*\\)`, "i");
  return (slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? []).filter(
    (paragraphXml) => tokenPattern.test(getParagraphText(paragraphXml)),
  ).length;
}

function replaceTokenPlaceholders(
  slideXml: string,
  token: string,
  replacements: string[],
): string {
  const tokenPattern = new RegExp(`\\(\\s*${token}\\s*\\)`, "gi");
  let replacementIndex = 0;

  return slideXml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (paragraphXml) => {
    const fullText = getParagraphText(paragraphXml);
    if (!tokenPattern.test(fullText)) return paragraphXml;

    tokenPattern.lastIndex = 0;
    const replacement = replacements[replacementIndex++] ?? "";
    const nextText = fullText.replace(tokenPattern, replacement);
    let isFirstTextRun = true;

    return paragraphXml.replace(
      /<a:t([^>]*)>[\s\S]*?<\/a:t>/g,
      (_match, attributes) => {
        if (!isFirstTextRun) return `<a:t${attributes}></a:t>`;
        isFirstTextRun = false;
        return `<a:t${attributes}>${escapeXml(nextText)}</a:t>`;
      },
    );
  });
}

function normalizeTitle(title: string): string {
  return title.trim().replace(/\s+/g, " ").toUpperCase();
}

function findBlock(
  result: EnglishMassParseResult,
  title: string,
): EnglishMassBlock | undefined {
  const normalizedTitle = normalizeTitle(title);
  return result.rites
    .flatMap((rite) => rite.blocks)
    .find((block) => normalizeTitle(block.title) === normalizedTitle);
}

function findBlockByType(
  result: EnglishMassParseResult,
  type: EnglishMassBlock["type"],
): EnglishMassBlock | undefined {
  return result.rites
    .flatMap((rite) => rite.blocks)
    .find((block) => block.type === type);
}

function findSongBlock(
  result: EnglishMassParseResult,
  songType: "ENTRANCE" | "OFFERTORY" | "COMMUNION" | "CLOSING",
): EnglishMassBlock | undefined {
  return result.rites
    .flatMap((rite) => rite.blocks)
    .find(
      (block) =>
        block.type === "song" &&
        normalizeTitle(block.title).startsWith(`${songType} SONG`),
    );
}

function getBlockBody(block?: EnglishMassBlock): string {
  if (!block) return "";
  if (block.body?.length) return block.body.join("\n\n");
  return block.paragraphs
    .slice(1)
    .map((paragraph) => paragraph.rawText.trim())
    .filter(Boolean)
    .join("\n\n");
}

function getGospelAcclamationBody(block?: EnglishMassBlock): string {
  return getBlockBody(block)
    .replace(
      /^(?:\s*(?:R\s*:\s*)?alleluia[!.,;:]*\s*){3}/i,
      "",
    )
    .trimStart();
}

function getMcPlaceholderBody(block?: EnglishMassBlock): string {
  return getBlockBody(block)
    .replace(
      /(^|\n\s*\n)\s*(?:MC|Main Celebrant \(MC\))\s*:\s*/gi,
      "$1",
    )
    .trim();
}

function isPrayerOfFaithfulTemplateClosing(line: string): boolean {
  const content = line
    .replace(/^\s*(?:MC|Main Celebrant \(MC\))\s*:\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();
  return /^We pray for those intentions offered for the Mass/i.test(content);
}

function getPrayerOfFaithfulOutro(lines: string[], closingIndex: number): string {
  if (closingIndex === -1) return "";

  const templateClosingLines = [
    /^We pray for those intentions offered for the Mass/i,
    /^We also pray for our personal intentions/i,
    /^Let us pray to the Lord/i,
    /^Lord, graciously hear us/i,
    /^Amen\.?$/i,
  ];

  return lines
    .slice(closingIndex + 1)
    .map((line) =>
      line
        .replace(
          /^\s*(?:MC|Main Celebrant \(MC\)|Cs|Celebrants \(Cs\))\s*:\s*/i,
          "",
        )
        .trim(),
    )
    .filter(
      (line) =>
        line && !templateClosingLines.some((templateLine) => templateLine.test(line)),
    )
    .join("\n\n");
}

function getSongTitle(block?: EnglishMassBlock): string {
  if (!block) return "";

  const inlineTitle = block.title
    .replace(/^(?:ENTRANCE|OFFERTORY|COMMUNION|CLOSING) SONG\s*(?::|[-–—])\s*/i, "")
    .trim();
  if (inlineTitle && !/^\(?\s*(?:stand|sit|kneel)\s*\)?$/i.test(inlineTitle)) {
    return inlineTitle;
  }

  return (
    block.paragraphs
      .slice(1)
      .map((paragraph) => paragraph.rawText.trim())
      .find(Boolean) ?? ""
  );
}

function cleanReadingSource(citation?: string): string {
  return (citation ?? "")
    .replace(/^A reading from(?: the)?(?: Holy Gospel)?\s*/i, "")
    .trim();
}

function splitTextIntoSlideChunks(
  text: string,
  maxCharsPerSlide = 260,
): string[] {
  const remainingText = text.trim();
  if (!remainingText) return [""];
  if (remainingText.length <= maxCharsPerSlide) return [remainingText];

  const chunks: string[] = [];
  let remaining = remainingText;

  while (remaining.length > maxCharsPerSlide) {
    const slice = remaining.slice(0, maxCharsPerSlide);
    const punctuationMatches = [...slice.matchAll(/[.,!?;]\s+/g)];
    const punctuationCut = punctuationMatches.at(-1);
    const cutAtPunctuation = punctuationCut
      ? punctuationCut.index! + punctuationCut[0].length
      : -1;
    const cutAtSpace = slice.lastIndexOf(" ");
    const minimumUsefulChunk = Math.floor(maxCharsPerSlide * 0.35);
    const cutIndex =
      cutAtPunctuation > minimumUsefulChunk
        ? cutAtPunctuation
        : cutAtSpace > minimumUsefulChunk
          ? cutAtSpace
          : maxCharsPerSlide;

    chunks.push(remaining.slice(0, cutIndex).trim());
    remaining = remaining.slice(cutIndex).trim();
  }

  if (remaining) chunks.push(remaining);
  return chunks;
}

function clipTextOverflow(slideXml: string): string {
  return slideXml.replace(/<a:bodyPr\b([^>]*)>/g, (bodyPrXml, attributes) => {
    if (/\bvertOverflow=/.test(bodyPrXml)) return bodyPrXml;
    return `<a:bodyPr${attributes} vertOverflow="clip">`;
  });
}

function removeParagraphWithText(slideXml: string, text: string): string {
  const normalizedText = normalizeTitle(text);
  return slideXml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (paragraphXml) =>
    normalizeTitle(getParagraphText(paragraphXml)) === normalizedText
      ? ""
      : paragraphXml,
  );
}

function getSlidesInPresentationOrder(
  zip: JSZip,
  presentationXml: string,
  relsXml: string,
) {
  const targetByRelationshipId: Record<string, string> = {};
  for (const match of relsXml.matchAll(
    /<Relationship [^>]*Id="([^"]+)" [^>]*Target="([^"]+)"/g,
  )) {
    targetByRelationshipId[match[1]] = match[2];
  }

  return [...presentationXml.matchAll(/<p:sldId [^>]*r:id="([^"]+)"[^>]*>/g)]
    .map((match) => ({
      rId: match[1],
      target: targetByRelationshipId[match[1]],
    }))
    .filter((slide): slide is { rId: string; target: string } =>
      Boolean(slide.target?.startsWith("slides/slide")),
    )
    .map((slide) => ({ rId: slide.rId, fileName: `ppt/${slide.target}` }))
    .filter((slide) => Boolean(zip.file(slide.fileName)));
}

export async function generateEnglishMassPptxBuffer(
  parseResult: EnglishMassParseResult,
): Promise<Buffer> {
  const templatePath = path.join(
    process.cwd(),
    "app/template",
    ENGLISH_TEMPLATE_FILE,
  );
  const zip = await JSZip.loadAsync(await fs.promises.readFile(templatePath));
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const presentationRelsXml = await zip
    .file("ppt/_rels/presentation.xml.rels")
    ?.async("text");
  if (!presentationXml || !presentationRelsXml)
    throw new Error("English Mass template is invalid.");

  const entranceSong = findSongBlock(parseResult, "ENTRANCE");
  const offertorySong = findSongBlock(parseResult, "OFFERTORY");
  const communionSong = findSongBlock(parseResult, "COMMUNION");
  const closingSong = findSongBlock(parseResult, "CLOSING");
  const firstReading = findBlock(parseResult, "FIRST READING");
  const secondReading = findBlock(parseResult, "SECOND READING");
  const gospel = findBlock(parseResult, "GOSPEL");
  const psalm = findBlockByType(parseResult, "responsorial_psalm");
  const gospelAcclamation = findBlockByType(
    parseResult,
    "gospel_acclamation",
  );
  const intercessions = findBlockByType(parseResult, "intercessions");
  const kyrie = findBlock(parseResult, "LORD HAVE MERCY");
  const kyrieLines =
    kyrie?.dialogues
      .filter((line) => line.speaker === "MC")
      .map((line) => line.text) ?? [];
  const parsedIntercessionLines =
    intercessions?.paragraphs
      .slice(1)
      .map((paragraph) => paragraph.rawText.trim())
      .filter(Boolean) ?? [];
  const intercessionClosingIndex = parsedIntercessionLines.findIndex(
    isPrayerOfFaithfulTemplateClosing,
  );
  const intercessionLines =
    intercessionClosingIndex === -1
      ? parsedIntercessionLines
      : parsedIntercessionLines.slice(0, intercessionClosingIndex);
  let pendingIntercessionOutro = getPrayerOfFaithfulOutro(
    parsedIntercessionLines,
    intercessionClosingIndex,
  );

  let pendingSongTitle: string | null = null;
  let pendingReadingBody: EnglishMassBlock | undefined;
  let pendingReadingUsesContinuationSlides = false;
  let pendingPsalmVerses: string[] = [];
  let pendingPsalmRefrain = "";
  let hasReachedIntercessions = false;
  let kyrieIndex = 0;

  const slideOrder = getSlidesInPresentationOrder(
    zip,
    presentationXml,
    presentationRelsXml,
  );
  const relationshipIds = [
    ...presentationRelsXml.matchAll(/<Relationship [^>]*Id="([^"]+)"/g),
  ]
    .map((match) => Number(match[1].replace(/^rId/, "")))
    .filter(Number.isFinite);
  let maxRelationshipId = Math.max(0, ...relationshipIds);
  const slideIds = [...presentationXml.matchAll(/<p:sldId [^>]*id="(\d+)"/g)]
    .map((match) => Number(match[1]))
    .filter(Number.isFinite);
  let maxSlideId = Math.max(255, ...slideIds);

  const duplicateSlide = async (
    sourceFileName: string,
    afterRelationshipId: string,
    newSlideXml: string,
  ) => {
    const existingSlideNumbers = Object.keys(zip.files)
      .map((fileName) => fileName.match(/^ppt\/slides\/slide(\d+)\.xml$/)?.[1])
      .filter((slideNumber): slideNumber is string => Boolean(slideNumber))
      .map(Number);
    const newSlideNumber = Math.max(0, ...existingSlideNumbers) + 1;
    const newFileName = `ppt/slides/slide${newSlideNumber}.xml`;
    const newRelationshipId = `rId${++maxRelationshipId}`;
    const newSlideId = ++maxSlideId;
    zip.file(newFileName, newSlideXml);

    const sourceSlideNumber = sourceFileName.match(/slide(\d+)\.xml$/)?.[1];
    const sourceRelsFile = sourceSlideNumber
      ? zip.file(`ppt/slides/_rels/slide${sourceSlideNumber}.xml.rels`)
      : null;
    if (sourceRelsFile) {
      const sourceRelsXml = await sourceRelsFile.async("text");
      zip.file(
        `ppt/slides/_rels/slide${newSlideNumber}.xml.rels`,
        sourceRelsXml.replace(
          /<Relationship [^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesSlide"[^>]*\/>/g,
          "",
        ),
      );
    }

    const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
    if (contentTypes) {
      zip.file(
        "[Content_Types].xml",
        contentTypes.replace(
          "</Types>",
          `<Override PartName="/ppt/slides/slide${newSlideNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
        ),
      );
    }

    const presentationRels = await zip
      .file("ppt/_rels/presentation.xml.rels")
      ?.async("text");
    if (presentationRels) {
      zip.file(
        "ppt/_rels/presentation.xml.rels",
        presentationRels.replace(
          "</Relationships>",
          `<Relationship Id="${newRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNumber}.xml"/></Relationships>`,
        ),
      );
    }

    const currentPresentation = await zip
      .file("ppt/presentation.xml")
      ?.async("text");
    if (currentPresentation) {
      const newSlideTag = `<p:sldId id="${newSlideId}" r:id="${newRelationshipId}"/>`;
      const afterSlidePattern = new RegExp(
        `(<p:sldId [^>]*r:id="${afterRelationshipId}"[^>]*\/?>)`,
      );
      zip.file(
        "ppt/presentation.xml",
        afterSlidePattern.test(currentPresentation)
          ? currentPresentation.replace(afterSlidePattern, `$1${newSlideTag}`)
          : currentPresentation.replace(
              "</p:sldIdLst>",
              `${newSlideTag}</p:sldIdLst>`,
            ),
      );
    }

    return { newFileName, newRelationshipId };
  };

  for (const slide of slideOrder) {
    const slideFile = zip.file(slide.fileName);
    if (!slideFile) continue;
    let slideXml = await slideFile.async("text");
    const slideText = getSlideText(slideXml);
    const textPlaceholderCount = countTokenPlaceholders(slideXml, "teks");

    if (slideText.includes("ENTRANCE SONG"))
      pendingSongTitle = getSongTitle(entranceSong);
    if (slideText.includes("OFFERTORY SONG"))
      pendingSongTitle = getSongTitle(offertorySong);
    if (slideText.includes("COMMUNION SONG"))
      pendingSongTitle = getSongTitle(communionSong);
    if (slideText.includes("CLOSING SONG"))
      pendingSongTitle = getSongTitle(closingSong);

    // Song lyrics remain as the template's (teks) placeholder by design.
    if (
      pendingSongTitle !== null &&
      countTokenPlaceholders(slideXml, "judul") > 0
    ) {
      slideXml = replaceTokenPlaceholders(slideXml, "judul", [
        pendingSongTitle,
      ]);
      pendingSongTitle = null;
    } else if (slideText.includes("OPENING PRAYER") && textPlaceholderCount) {
      const prayerChunks = splitTextIntoSlideChunks(
        getBlockBody(findBlock(parseResult, "OPENING PRAYER")),
      );
      zip.file(
        slide.fileName,
        clipTextOverflow(
          replaceTokenPlaceholders(slideXml, "teks", [prayerChunks[0]]),
        ),
      );

      let previousRelationshipId = slide.rId;
      for (const prayerChunk of prayerChunks.slice(1)) {
        const continuationSlide = clipTextOverflow(
          removeParagraphWithText(
            removeParagraphWithText(
              replaceTokenPlaceholders(slideXml, "teks", [prayerChunk]),
              "OPENING PRAYER",
            ),
            "MC: Let us pray.",
          ),
        );
        const duplicate = await duplicateSlide(
          slide.fileName,
          previousRelationshipId,
          continuationSlide,
        );
        previousRelationshipId = duplicate.newRelationshipId;
      }
      continue;
    } else if (slideText.includes("FIRST READING") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        firstReading?.epigraph ?? "",
        cleanReadingSource(firstReading?.citation),
      ]);
      pendingReadingBody = firstReading;
      pendingReadingUsesContinuationSlides = true;
    } else if (slideText.includes("SECOND READING") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        secondReading?.epigraph ?? "",
        cleanReadingSource(secondReading?.citation),
      ]);
      pendingReadingBody = secondReading;
      pendingReadingUsesContinuationSlides = true;
    } else if (
      slideText.includes("RESPONSORIAL PSALM") &&
      textPlaceholderCount
    ) {
      const verses = psalm?.body ?? [];
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        psalm?.refrain ?? "",
        verses[0] ?? "",
      ]);
      pendingPsalmVerses = verses.slice(1);
      pendingPsalmRefrain = psalm?.refrain ?? "";
    } else if (
      slideText.includes("GOSPEL ACCLAMATION") &&
      textPlaceholderCount
    ) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        getGospelAcclamationBody(gospelAcclamation),
      ]);
    } else if (slideText.includes("GOSPEL") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        cleanReadingSource(gospel?.citation),
        gospel?.epigraph ?? "",
      ]);
      pendingReadingBody = gospel;
      pendingReadingUsesContinuationSlides = true;
    } else if (
      slideText.includes("PRAYER OVER THE OFFERINGS") &&
      textPlaceholderCount
    ) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        getMcPlaceholderBody(
          findBlock(parseResult, "PRAYER OVER THE OFFERINGS"),
        ),
      ]);
    } else if (
      slideText.includes("PRAYER AFTER COMMUNION") &&
      textPlaceholderCount
    ) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        getBlockBody(findBlock(parseResult, "PRAYER AFTER COMMUNION")),
      ]);
    } else if (
      hasReachedIntercessions &&
      pendingIntercessionOutro &&
      /CS\s*:\s*AMEN/.test(slideText) &&
      textPlaceholderCount
    ) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        pendingIntercessionOutro,
      ]);
      pendingIntercessionOutro = "";
    } else if (
      hasReachedIntercessions &&
      textPlaceholderCount === 2 &&
      intercessionLines.length
    ) {
      zip.file(
        slide.fileName,
        replaceTokenPlaceholders(slideXml, "teks", [
          intercessionLines.shift() ?? "",
          intercessionLines.shift() ?? "",
        ]),
      );

      let previousRelationshipId = slide.rId;
      while (intercessionLines.length) {
        const duplicate = await duplicateSlide(
          slide.fileName,
          previousRelationshipId,
          replaceTokenPlaceholders(slideXml, "teks", [
            intercessionLines.shift() ?? "",
            intercessionLines.shift() ?? "",
          ]),
        );
        previousRelationshipId = duplicate.newRelationshipId;
      }
      continue;
    } else if (hasReachedIntercessions && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(
        slideXml,
        "teks",
        Array.from(
          { length: textPlaceholderCount },
          () => intercessionLines.shift() ?? "",
        ),
      );
    } else if (pendingReadingBody && textPlaceholderCount) {
      const readingBody = getBlockBody(pendingReadingBody);
      if (pendingReadingUsesContinuationSlides) {
        const readingChunks = splitTextIntoSlideChunks(readingBody, 380);
        zip.file(
          slide.fileName,
          clipTextOverflow(
            replaceTokenPlaceholders(slideXml, "teks", [readingChunks[0]]),
          ),
        );

        let previousRelationshipId = slide.rId;
        for (const readingChunk of readingChunks.slice(1)) {
          const duplicate = await duplicateSlide(
            slide.fileName,
            previousRelationshipId,
            clipTextOverflow(
              replaceTokenPlaceholders(slideXml, "teks", [readingChunk]),
            ),
          );
          previousRelationshipId = duplicate.newRelationshipId;
        }
        pendingReadingBody = undefined;
        pendingReadingUsesContinuationSlides = false;
        continue;
      }

      slideXml = replaceTokenPlaceholders(slideXml, "teks", [readingBody]);
      pendingReadingBody = undefined;
    } else if (pendingPsalmVerses.length && textPlaceholderCount) {
      const firstContinuationVerse = pendingPsalmVerses.shift() ?? "";
      zip.file(
        slide.fileName,
        replaceTokenPlaceholders(slideXml, "teks", [
          firstContinuationVerse,
          pendingPsalmRefrain,
        ]),
      );

      let previousRelationshipId = slide.rId;
      for (const verse of pendingPsalmVerses) {
        const duplicate = await duplicateSlide(
          slide.fileName,
          previousRelationshipId,
          replaceTokenPlaceholders(slideXml, "teks", [
            verse,
            pendingPsalmRefrain,
          ]),
        );
        previousRelationshipId = duplicate.newRelationshipId;
      }
      pendingPsalmVerses = [];
      pendingPsalmRefrain = "";
      continue;
    } else if (
      textPlaceholderCount &&
      /(?:LORD HAVE MERCY|CHRIST HAVE MERCY)/.test(slideText)
    ) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [
        kyrieLines[kyrieIndex++] ?? "",
      ]);
    }

    if (slideText.includes("PRAYER OF THE FAITHFUL"))
      hasReachedIntercessions = true;
    if (slideText.includes("LITURGY OF EUCHARIST"))
      hasReachedIntercessions = false;
    zip.file(slide.fileName, slideXml);
  }

  return zip.generateAsync({ type: "nodebuffer" });
}
