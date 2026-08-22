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
  return (slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? []).filter((paragraphXml) =>
    tokenPattern.test(getParagraphText(paragraphXml)),
  ).length;
}

function replaceTokenPlaceholders(slideXml: string, token: string, replacements: string[]): string {
  const tokenPattern = new RegExp(`\\(\\s*${token}\\s*\\)`, "gi");
  let replacementIndex = 0;

  return slideXml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (paragraphXml) => {
    const fullText = getParagraphText(paragraphXml);
    if (!tokenPattern.test(fullText)) return paragraphXml;

    tokenPattern.lastIndex = 0;
    const replacement = replacements[replacementIndex++] ?? "";
    const nextText = fullText.replace(tokenPattern, replacement);
    let isFirstTextRun = true;

    return paragraphXml.replace(/<a:t([^>]*)>[\s\S]*?<\/a:t>/g, (_match, attributes) => {
      if (!isFirstTextRun) return `<a:t${attributes}></a:t>`;
      isFirstTextRun = false;
      return `<a:t${attributes}>${escapeXml(nextText)}</a:t>`;
    });
  });
}

function normalizeTitle(title: string): string {
  return title.trim().replace(/\s+/g, " ").toUpperCase();
}

function findBlock(result: EnglishMassParseResult, title: string): EnglishMassBlock | undefined {
  const normalizedTitle = normalizeTitle(title);
  return result.rites
    .flatMap((rite) => rite.blocks)
    .find((block) => normalizeTitle(block.title) === normalizedTitle);
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

function getSongTitle(block?: EnglishMassBlock): string {
  return block?.paragraphs.slice(1).map((paragraph) => paragraph.rawText.trim()).find(Boolean) ?? "";
}

function cleanReadingSource(citation?: string): string {
  return (citation ?? "")
    .replace(/^A reading from(?: the)?(?: Holy Gospel)?\s*/i, "")
    .trim();
}

function getSlidesInPresentationOrder(zip: JSZip, presentationXml: string, relsXml: string) {
  const targetByRelationshipId: Record<string, string> = {};
  for (const match of relsXml.matchAll(/<Relationship [^>]*Id="([^"]+)" [^>]*Target="([^"]+)"/g)) {
    targetByRelationshipId[match[1]] = match[2];
  }

  return [...presentationXml.matchAll(/<p:sldId [^>]*r:id="([^"]+)"[^>]*>/g)]
    .map((match) => targetByRelationshipId[match[1]])
    .filter((target): target is string => Boolean(target?.startsWith("slides/slide")))
    .map((target) => `ppt/${target}`)
    .filter((fileName) => Boolean(zip.file(fileName)));
}

export async function generateEnglishMassPptxBuffer(parseResult: EnglishMassParseResult): Promise<Buffer> {
  const templatePath = path.join(process.cwd(), "app/template", ENGLISH_TEMPLATE_FILE);
  const zip = await JSZip.loadAsync(await fs.promises.readFile(templatePath));
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const presentationRelsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");
  if (!presentationXml || !presentationRelsXml) throw new Error("English Mass template is invalid.");

  const entranceSong = findBlock(parseResult, "ENTRANCE SONG");
  const offertorySong = findBlock(parseResult, "OFFERTORY SONG");
  const communionSong = findBlock(parseResult, "COMMUNION SONG");
  const closingSong = findBlock(parseResult, "CLOSING SONG");
  const firstReading = findBlock(parseResult, "FIRST READING");
  const secondReading = findBlock(parseResult, "SECOND READING");
  const gospel = findBlock(parseResult, "GOSPEL");
  const psalm = findBlock(parseResult, "RESPONSORIAL PSALM");
  const intercessions = findBlock(parseResult, "PRAYER OF THE FAITHFUL");
  const kyrie = findBlock(parseResult, "LORD HAVE MERCY");
  const kyrieLines = kyrie?.dialogues.filter((line) => line.speaker === "MC").map((line) => line.text) ?? [];
  const intercessionLines = intercessions?.paragraphs.slice(1).map((paragraph) => paragraph.rawText.trim()).filter(Boolean) ?? [];

  let pendingSongTitle = "";
  let pendingReadingBody: EnglishMassBlock | undefined;
  let pendingPsalmVerses: string[] = [];
  let hasReachedIntercessions = false;
  let kyrieIndex = 0;

  for (const fileName of getSlidesInPresentationOrder(zip, presentationXml, presentationRelsXml)) {
    const slideFile = zip.file(fileName);
    if (!slideFile) continue;
    let slideXml = await slideFile.async("text");
    const slideText = getSlideText(slideXml);
    const textPlaceholderCount = countTokenPlaceholders(slideXml, "teks");

    if (slideText.includes("ENTRANCE SONG")) pendingSongTitle = getSongTitle(entranceSong);
    if (slideText.includes("OFFERTORY SONG")) pendingSongTitle = getSongTitle(offertorySong);
    if (slideText.includes("COMMUNION SONG")) pendingSongTitle = getSongTitle(communionSong);
    if (slideText.includes("CLOSING SONG")) pendingSongTitle = getSongTitle(closingSong);

    // Song lyrics remain as the template's (teks) placeholder by design.
    if (pendingSongTitle && countTokenPlaceholders(slideXml, "judul") > 0) {
      slideXml = replaceTokenPlaceholders(slideXml, "judul", [pendingSongTitle]);
      pendingSongTitle = "";
    } else if (slideText.includes("OPENING PRAYER") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [getBlockBody(findBlock(parseResult, "OPENING PRAYER"))]);
    } else if (slideText.includes("FIRST READING") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [firstReading?.epigraph ?? "", cleanReadingSource(firstReading?.citation)]);
      pendingReadingBody = firstReading;
    } else if (slideText.includes("SECOND READING") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [secondReading?.epigraph ?? "", cleanReadingSource(secondReading?.citation)]);
      pendingReadingBody = secondReading;
    } else if (slideText.includes("RESPONSORIAL PSALM") && textPlaceholderCount) {
      const verses = psalm?.body ?? [];
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [psalm?.refrain ?? "", verses[0] ?? ""]);
      pendingPsalmVerses = verses.slice(1);
    } else if (slideText.includes("GOSPEL ACCLAMATION") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [getBlockBody(findBlock(parseResult, "GOSPEL ACCLAMATION"))]);
    } else if (slideText.includes("GOSPEL") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [cleanReadingSource(gospel?.citation), gospel?.epigraph ?? ""]);
      pendingReadingBody = gospel;
    } else if (slideText.includes("PRAYER OVER THE OFFERINGS") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [getBlockBody(findBlock(parseResult, "PRAYER OVER THE OFFERINGS"))]);
    } else if (slideText.includes("PRAYER AFTER COMMUNION") && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [getBlockBody(findBlock(parseResult, "PRAYER AFTER COMMUNION"))]);
    } else if (hasReachedIntercessions && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", Array.from({ length: textPlaceholderCount }, () => intercessionLines.shift() ?? ""));
    } else if (pendingReadingBody && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [getBlockBody(pendingReadingBody)]);
      pendingReadingBody = undefined;
    } else if (pendingPsalmVerses.length && textPlaceholderCount) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", Array.from({ length: textPlaceholderCount }, () => pendingPsalmVerses.shift() ?? ""));
    } else if (textPlaceholderCount && /(?:LORD HAVE MERCY|CHRIST HAVE MERCY)/.test(slideText)) {
      slideXml = replaceTokenPlaceholders(slideXml, "teks", [kyrieLines[kyrieIndex++] ?? ""]);
    }

    if (slideText.includes("PRAYER OF THE FAITHFUL")) hasReachedIntercessions = true;
    if (slideText.includes("LITURGY OF EUCHARIST")) hasReachedIntercessions = false;
    zip.file(fileName, slideXml);
  }

  return zip.generateAsync({ type: "nodebuffer" });
}
