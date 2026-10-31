import fs from "fs";
import path from "path";
import JSZip from "jszip";
import {
  EnglishMassBlock,
  EnglishMassParseResult,
} from "../utils/englishMassParser";
import { createEnglishMassTextLayout } from "./englishMassTextLayout";

const ENGLISH_TEMPLATE_FILE = "EM_Template.pptx";

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
    .replace(/^(?:\s*(?:R\s*:\s*)?alleluia[!.,;:]*\s*){3}/i, "")
    .trimStart();
}

function getMcPlaceholderBody(block?: EnglishMassBlock): string {
  return getBlockBody(block)
    .replace(/(^|\n\s*\n)\s*(?:MC|Main Celebrant \(MC\))\s*:\s*/gi, "$1")
    .trim();
}

function isPrayerOfFaithfulTemplateClosing(line: string): boolean {
  const content = line
    .replace(/^\s*(?:MC|Main Celebrant \(MC\))\s*:\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();
  return /^We pray for those intentions offered for the Mass/i.test(content);
}

function getPrayerOfFaithfulOutro(
  lines: string[],
  closingIndex: number,
): string {
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
        line &&
        !templateClosingLines.some((templateLine) => templateLine.test(line)),
    )
    .join("\n\n");
}

function getSongTitle(block?: EnglishMassBlock): string {
  if (!block) return "";

  const inlineTitle = block.title
    .replace(
      /^(?:ENTRANCE|OFFERTORY|COMMUNION|CLOSING) SONG\s*(?::|[-–—])\s*/i,
      "",
    )
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
  const textLayout = await createEnglishMassTextLayout(zip);
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
  const gospelAcclamation = findBlockByType(parseResult, "gospel_acclamation");
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
  let awaitingReadingBody = false;
  let pendingPsalmVerses: string[] = [];
  let pendingPsalmRefrain = "";
  let awaitingPsalmContinuation = false;
  let hasReachedIntercessions = false;
  let kyrieIndex = 0;
  let awaitingHomilyText = false;
  let homilyInserted = false;
  const homilyText = parseResult.homily?.paragraphs.join("\n\n").trim() ?? "";

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

  const sourceSlides = new Map(
    await Promise.all(
      slideOrder.map(
        async (slide) =>
          [
            slide.fileName,
            await zip.file(slide.fileName)!.async("text"),
          ] as const,
      ),
    ),
  );
  const writeSlidePages = async (
    slide: (typeof slideOrder)[number],
    pages: readonly string[],
  ) => {
    zip.file(slide.fileName, pages[0]);
    let previousRelationshipId = slide.rId;
    for (const page of pages.slice(1)) {
      const duplicate = await duplicateSlide(
        slide.fileName,
        previousRelationshipId,
        page,
      );
      previousRelationshipId = duplicate.newRelationshipId;
    }
  };

  for (const [slideIndex, slide] of slideOrder.entries()) {
    const slideXml = sourceSlides.get(slide.fileName)!;
    const slideText = getSlideText(slideXml);
    const textPlaceholderCount = countTokenPlaceholders(slideXml, "teks");
    const paragraphTexts = (slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? []).map(
      getParagraphText,
    );
    const matchingTexts = (pattern: RegExp) =>
      paragraphTexts.filter((text) => pattern.test(text.trim()));
    const prayerOptions = {
      firstOnlyTexts: matchingTexts(
        /^(?:OPENING PRAYER|PRAYER OVER THE OFFERINGS|PRAYER AFTER COMMUNION|MC\s*:)/i,
      ),
      lastOnlyTexts: matchingTexts(/^Cs\s*:\s*Amen[.!]?$/i),
    };
    const readingOptions = {
      continuationXml: sourceSlides.get(
        slideOrder[slideIndex + 1]?.fileName ?? "",
      ),
      firstOnlyTexts: matchingTexts(
        /^(?:FIRST READING|SECOND READING|GOSPEL$|MC\s*:|A reading from)/i,
      ),
    };
    let pages: string[] | undefined;

    if (/\bHOMILY\b/.test(slideText)) awaitingHomilyText = true;
    if (
      /\b(?:PROFES+SION OF FAITH|NICENE CREED|PRAYER OF THE FAITHFUL)\b/.test(
        slideText,
      )
    )
      awaitingHomilyText = false;
    if (awaitingHomilyText && textPlaceholderCount) {
      await writeSlidePages(
        slide,
        textLayout.paginate(slide.fileName, slideXml, "teks", [homilyText]),
      );
      homilyInserted = true;
      awaitingHomilyText = false;
      continue;
    }

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
      pages = textLayout.paginate(slide.fileName, slideXml, "judul", [
        pendingSongTitle,
      ]);
      pendingSongTitle = null;
    } else if (slideText.includes("OPENING PRAYER") && textPlaceholderCount) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [getBlockBody(findBlock(parseResult, "OPENING PRAYER"))],
        prayerOptions,
      );
    } else if (slideText.includes("FIRST READING") && textPlaceholderCount) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [
          firstReading?.epigraph ?? "",
          cleanReadingSource(firstReading?.citation),
        ],
        readingOptions,
      );
      pendingReadingBody = firstReading;
      awaitingReadingBody = true;
    } else if (slideText.includes("SECOND READING") && textPlaceholderCount) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [
          secondReading?.epigraph ?? "",
          cleanReadingSource(secondReading?.citation),
        ],
        readingOptions,
      );
      pendingReadingBody = secondReading;
      awaitingReadingBody = true;
    } else if (
      slideText.includes("RESPONSORIAL PSALM") &&
      textPlaceholderCount
    ) {
      const verses = psalm?.body ?? [];
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [psalm?.refrain ?? "", verses[0] ?? ""],
        {
          firstOnlyTexts: matchingTexts(/^RESPONSORIAL PSALM/i),
        },
      );
      pendingPsalmVerses = verses.slice(1);
      pendingPsalmRefrain = psalm?.refrain ?? "";
      awaitingPsalmContinuation = true;
    } else if (
      slideText.includes("GOSPEL ACCLAMATION") &&
      textPlaceholderCount
    ) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [getGospelAcclamationBody(gospelAcclamation)],
        {
          firstOnlyTexts: matchingTexts(
            /^(?:GOSPEL ACCLAMATION|Alleluia!|\(\s*stand\s*\))/i,
          ),
        },
      );
    } else if (slideText.includes("GOSPEL") && textPlaceholderCount) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [cleanReadingSource(gospel?.citation), gospel?.epigraph ?? ""],
        readingOptions,
      );
      pendingReadingBody = gospel;
      awaitingReadingBody = true;
    } else if (
      slideText.includes("PRAYER OVER THE OFFERINGS") &&
      textPlaceholderCount
    ) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [
          getMcPlaceholderBody(
            findBlock(parseResult, "PRAYER OVER THE OFFERINGS"),
          ),
        ],
        prayerOptions,
      );
    } else if (
      slideText.includes("PRAYER AFTER COMMUNION") &&
      textPlaceholderCount
    ) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [getBlockBody(findBlock(parseResult, "PRAYER AFTER COMMUNION"))],
        prayerOptions,
      );
    } else if (
      hasReachedIntercessions &&
      /CS\s*:\s*AMEN/.test(slideText) &&
      textPlaceholderCount
    ) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [pendingIntercessionOutro],
        prayerOptions,
      );
      pendingIntercessionOutro = "";
    } else if (hasReachedIntercessions && textPlaceholderCount === 2) {
      pages = [];
      do {
        pages.push(
          ...textLayout.paginate(slide.fileName, slideXml, "teks", [
            intercessionLines.shift() ?? "",
            intercessionLines.shift() ?? "",
          ]),
        );
      } while (intercessionLines.length);
    } else if (hasReachedIntercessions && textPlaceholderCount) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        Array.from(
          { length: textPlaceholderCount },
          () => intercessionLines.shift() ?? "",
        ),
      );
    } else if (awaitingReadingBody && textPlaceholderCount) {
      pages = textLayout.paginate(slide.fileName, slideXml, "teks", [
        getBlockBody(pendingReadingBody),
      ]);
      pendingReadingBody = undefined;
      awaitingReadingBody = false;
    } else if (awaitingPsalmContinuation && textPlaceholderCount) {
      pages = [];
      // Each verse keeps its response after it, even when the pair spills.
      if (pendingPsalmVerses.length) {
        for (const verse of pendingPsalmVerses) {
          pages.push(
            ...textLayout.paginate(slide.fileName, slideXml, "teks", [
              verse,
              pendingPsalmRefrain,
            ]),
          );
        }
      } else {
        pages = textLayout.paginate(slide.fileName, slideXml, "teks", [
          "",
          pendingPsalmRefrain,
        ]);
      }
      pendingPsalmVerses = [];
      pendingPsalmRefrain = "";
      awaitingPsalmContinuation = false;
    } else if (
      textPlaceholderCount &&
      /(?:LORD HAVE MERCY|CHRIST HAVE MERCY)/.test(slideText)
    ) {
      pages = textLayout.paginate(
        slide.fileName,
        slideXml,
        "teks",
        [kyrieLines[kyrieIndex++] ?? ""],
        {
          firstOnlyTexts: matchingTexts(
            /^(?:MC\s*:|LORD HAVE MERCY|\(\s*stand\s*\))/i,
          ),
          lastOnlyTexts: matchingTexts(/^(?:Choir|Cs)\s*:/i),
        },
      );
    }

    if (slideText.includes("PRAYER OF THE FAITHFUL"))
      hasReachedIntercessions = true;
    if (slideText.includes("LITURGY OF EUCHARIST"))
      hasReachedIntercessions = false;
    if (pages) await writeSlidePages(slide, pages);
  }

  if (homilyText && !homilyInserted)
    throw new Error(
      "Placeholder (teks) setelah HOMILY tidak ditemukan di template English Mass.",
    );

  return zip.generateAsync({ type: "nodebuffer" });
}
