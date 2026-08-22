import {
  LiturgyParagraph,
  parseDocFile,
  parseDocxFile,
} from "./docxParser";

export type EnglishMassRiteType =
  | "introductory_rites"
  | "liturgy_of_the_word"
  | "liturgy_of_the_eucharist"
  | "communion_rite"
  | "concluding_rite"
  | "unclassified";

export type EnglishMassBlockType =
  | "major_rite"
  | "section_heading"
  | "song"
  | "rubric"
  | "dialogue"
  | "prayer"
  | "reading"
  | "responsorial_psalm"
  | "gospel_acclamation"
  | "creed"
  | "intercessions"
  | "content";

export type EnglishMassSpeaker = "MC" | "Cs" | "R" | "Choir";

export interface EnglishMassDialogueLine {
  speaker: EnglishMassSpeaker;
  text: string;
  paragraphIndex: number;
}

export interface EnglishMassBlock {
  id: string;
  type: EnglishMassBlockType;
  title: string;
  posture?: "Stand" | "Sit" | "Kneel";
  paragraphs: LiturgyParagraph[];
  dialogues: EnglishMassDialogueLine[];
  epigraph?: string;
  citation?: string;
  refrain?: string;
  body?: string[];
}

export interface EnglishMassRite {
  id: EnglishMassRiteType;
  title: string;
  paragraphs: LiturgyParagraph[];
  blocks: EnglishMassBlock[];
}

export interface EnglishMassParseResult {
  massType: "english";
  fileName: string;
  totalParagraphs: number;
  rawParagraphs: LiturgyParagraph[];
  rites: EnglishMassRite[];
  unclassifiedParagraphs: LiturgyParagraph[];
}

const majorRites: Array<{ type: EnglishMassRiteType; title: string; pattern: RegExp }> = [
  { type: "introductory_rites", title: "Introductory Rites", pattern: /^INTRODUCTORY RITES?$/ },
  { type: "liturgy_of_the_word", title: "Liturgy of the Word", pattern: /^LITURGY OF THE WORD$/ },
  { type: "liturgy_of_the_eucharist", title: "Liturgy of the Eucharist", pattern: /^LITURGY OF THE EUCHARIST$/ },
  { type: "communion_rite", title: "Communion Rite", pattern: /^COMMUNION RITE$/ },
  { type: "concluding_rite", title: "Concluding Rite", pattern: /^CONCLUDING RITE$/ },
];

const blockPatterns: Array<{ type: EnglishMassBlockType; pattern: RegExp }> = [
  { type: "song", pattern: /^(ENTRANCE|OFFERTORY|COMMUNION|CLOSING) SONG\b/ },
  { type: "reading", pattern: /^(FIRST|SECOND) READING\b/ },
  { type: "responsorial_psalm", pattern: /^RESPONSORIAL PSALM\b/ },
  { type: "gospel_acclamation", pattern: /^GOSPEL ACCLAMATION\b/ },
  { type: "reading", pattern: /^GOSPEL$/ },
  { type: "creed", pattern: /^(PROFESSION|PROFESSSION) OF FAITH\b|^NICENE CREED$/ },
  { type: "intercessions", pattern: /^PRAYER OF THE FAITHFUL\b/ },
  { type: "rubric", pattern: /^(THE PRIEST ADDRESSED|CHURCH NOTICES)\b/ },
  {
    type: "prayer",
    pattern: /^(?:OPENING PRAYER|PRAYER OVER THE OFFERINGS|PRAYER AFTER COMMUNION)$/,
  },
  {
    type: "section_heading",
    pattern: /^(?:GREETING|LORD HAVE MERCY|GLORIA|HOMILY|INVITATION TO PRAYER|THE EUCHARISTIC PRAYER|PREFACE|ACCLAMATION|EUCHARISTIC PRAYER|THE LORD'S PRAYER|THE LORD’S PRAYER|OUR FATHER|SIGN OF PEACE|LAMB OF GOD|BREAKING OF THE BREAD|COMMUNION|BLESSING AND DISMISSAL)$|^PRAYER FOR ARDAS\b/,
  },
];

function normalize(text: string): string {
  return text.replace(/\s+/g, " ").trim().toUpperCase();
}

function splitPosture(text: string) {
  const match = text.match(/\s+(Stand|Sit|Kneel)\s*$/i);
  if (!match) return { title: text.trim() };
  return {
    title: text.slice(0, match.index).trim(),
    posture: `${match[1][0].toUpperCase()}${match[1].slice(1).toLowerCase()}` as EnglishMassBlock["posture"],
  };
}

function getSpeakerLine(paragraph: LiturgyParagraph, paragraphIndex: number): EnglishMassDialogueLine | null {
  const match = paragraph.rawText.match(/^\s*(Main Celebrant \(MC\)|Celebrants \(Cs\)|MC|Cs|C\s+s|R|Choir)\s*:\s*(.*)$/i);
  if (!match) return null;

  const speakerText = match[1].replace(/\s+/g, "").toLowerCase();
  const speaker: EnglishMassSpeaker =
    speakerText.includes("mc") ? "MC" : speakerText.includes("cs") ? "Cs" : speakerText === "r" ? "R" : "Choir";
  return { speaker, text: match[2].trim(), paragraphIndex };
}

function isHeadingCandidate(paragraph: LiturgyParagraph): boolean {
  const text = paragraph.rawText.trim();
  const letters = text.match(/[A-Za-zÀ-ÿ]/g)?.join("") ?? "";
  return (
    text.length > 0 &&
    text.length <= 100 &&
    (paragraph.isBold || (letters.length >= 2 && letters === letters.toUpperCase()))
  );
}

function findBlockType(title: string): EnglishMassBlockType | null {
  const normalizedTitle = normalize(title);
  return blockPatterns.find(({ pattern }) => pattern.test(normalizedTitle))?.type ?? null;
}

function enrichBlock(block: EnglishMassBlock): EnglishMassBlock {
  if (block.type === "prayer") {
    return {
      ...block,
      body: block.paragraphs
        .slice(1)
        .filter((paragraph) => !getSpeakerLine(paragraph, 0))
        .map((paragraph) => paragraph.rawText.trim())
        .filter(Boolean),
    };
  }

  if (block.type === "reading") {
    const entries = block.paragraphs.map((paragraph, index) => ({
      paragraph,
      index,
      dialogue: getSpeakerLine(paragraph, index),
    }));
    const nonDialogue = entries
      .filter(({ dialogue }) => !dialogue);
    const citationEntry = entries.find(({ paragraph, dialogue }) =>
      /^(A reading from|A reading from the|A reading from the Holy Gospel)/i.test(
        (dialogue?.text ?? paragraph.rawText).trim(),
      ),
    );
    const citation = citationEntry
      ? (citationEntry.dialogue?.text ?? citationEntry.paragraph.rawText).trim()
      : undefined;
    const citationIndex = citationEntry?.index ?? -1;
    const epigraph = nonDialogue.find(({ paragraph, index }) =>
      index > 0 && index < citationIndex && (paragraph.isItalic || paragraph.rawText.trim().length < 160),
    )?.paragraph.rawText.trim();

    return {
      ...block,
      citation,
      epigraph,
      body: citationIndex >= 0
        ? block.paragraphs
            .slice(citationIndex + 1)
            .filter((paragraph) => !getSpeakerLine(paragraph, 0))
            .map((paragraph) => paragraph.rawText)
        : undefined,
    };
  }

  if (block.type === "responsorial_psalm") {
    const content = block.paragraphs.slice(1).filter((paragraph) => paragraph.rawText.trim());
    return {
      ...block,
      refrain: content[0]?.rawText.trim(),
      body: content.slice(1).map((paragraph) => paragraph.rawText),
    };
  }

  return block;
}

export function parseEnglishMassParagraphs(
  fileName: string,
  rawParagraphs: LiturgyParagraph[],
): EnglishMassParseResult {
  const rites: EnglishMassRite[] = [];
  const unclassifiedParagraphs: LiturgyParagraph[] = [];
  let currentRite: EnglishMassRite | null = null;
  let currentBlock: EnglishMassBlock | null = null;

  const finishBlock = () => {
    if (currentRite && currentBlock) {
      currentRite.blocks.push(enrichBlock(currentBlock));
    }
    currentBlock = null;
  };

  rawParagraphs.forEach((paragraph, index) => {
    const { title, posture } = splitPosture(paragraph.rawText);
    const normalizedTitle = normalize(title);
    const majorRite = majorRites.find(({ pattern }) => pattern.test(normalizedTitle));

    if (majorRite) {
      finishBlock();
      currentRite = {
        id: majorRite.type,
        title: majorRite.title,
        paragraphs: [paragraph],
        blocks: [],
      };
      rites.push(currentRite);
      return;
    }

    if (!currentRite) {
      unclassifiedParagraphs.push(paragraph);
      return;
    }

    currentRite.paragraphs.push(paragraph);
    const blockType = findBlockType(title);
    const dialogue = getSpeakerLine(paragraph, index);
    if (blockType || (isHeadingCandidate(paragraph) && !dialogue)) {
      finishBlock();
      currentBlock = {
        id: `${currentRite.id}-${currentRite.blocks.length + 1}`,
        type: blockType ?? "section_heading",
        title,
        posture,
        paragraphs: [paragraph],
        dialogues: [],
      };
      return;
    }

    if (!currentBlock) {
      currentBlock = {
        id: `${currentRite.id}-${currentRite.blocks.length + 1}`,
        type: "content",
        title: "Content",
        paragraphs: [],
        dialogues: [],
      };
    }
    currentBlock.paragraphs.push(paragraph);
    if (dialogue) currentBlock.dialogues.push(dialogue);
  });

  finishBlock();

  return {
    massType: "english",
    fileName,
    totalParagraphs: rawParagraphs.length,
    rawParagraphs,
    rites,
    unclassifiedParagraphs,
  };
}

export async function parseEnglishMassFile(file: File): Promise<EnglishMassParseResult> {
  const source = file.name.toLowerCase().endsWith(".doc")
    ? await parseDocFile(file)
    : await parseDocxFile(file);
  return parseEnglishMassParagraphs(source.fileName, source.rawParagraphs);
}
