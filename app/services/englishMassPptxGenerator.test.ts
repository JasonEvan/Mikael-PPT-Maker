import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { generateEnglishMassPptxBuffer } from "./englishMassPptxGenerator";
import type {
  EnglishMassBlock,
  EnglishMassParseResult,
} from "../utils/englishMassParser";

const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false });
const normalize = (value: string) => value.replace(/\s+/g, " ").trim();
const longText = (label: string) =>
  `${label}_START ${"Mercy is present in our daily work, in our families, and in our care for one another. ".repeat(12)}${label}_END`;

function block(
  title: string,
  type: EnglishMassBlock["type"],
  body: string[] = [],
): EnglishMassBlock {
  return { id: title, title, type, body, paragraphs: [], dialogues: [] };
}

function input(
  blocks: EnglishMassBlock[],
  homily?: string[],
): EnglishMassParseResult {
  return {
    massType: "english",
    fileName: "layout-regression.docx",
    totalParagraphs: 0,
    rawParagraphs: [],
    unclassifiedParagraphs: [],
    rites: [
      { id: "liturgy_of_the_word", title: "Liturgy", paragraphs: [], blocks },
    ],
    ...(homily
      ? { homily: { fileName: "homily.docx", paragraphs: homily } }
      : {}),
  };
}

async function slides(
  buffer: Buffer,
): Promise<{ xml: string; text: string }[]> {
  const zip = await JSZip.loadAsync(buffer);
  const presentation = parser.parse(
    await zip.file("ppt/presentation.xml")!.async("text"),
  );
  const relationships = parser.parse(
    await zip.file("ppt/_rels/presentation.xml.rels")!.async("text"),
  );
  const targets = new Map<string, string>(
    relationships.Relationships.Relationship.map(
      (rel: Record<string, string>) => [rel["@_Id"], rel["@_Target"]],
    ),
  );
  const ids = presentation["p:presentation"]["p:sldIdLst"]["p:sldId"];
  return Promise.all(
    ids.map(async (id: Record<string, string>) => {
      const target = targets.get(id["@_r:id"]);
      assert.ok(target, "every presentation slide has a relationship");
      const xml = await zip
        .file(path.posix.normalize(`ppt/${target}`))!
        .async("text");
      const text = [...xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
        .map((match) => String(parser.parse(`<t>${match[1]}</t>`).t ?? ""))
        .join(" ");
      return { xml, text: normalize(text) };
    }),
  );
}

test("every English dynamic role preserves long text through measured continuations", async () => {
  const expected: string[] = [];
  const blocks: EnglishMassBlock[] = [];
  for (const title of ["FIRST READING", "SECOND READING", "GOSPEL"]) {
    const body = longText(title.replaceAll(" ", "_"));
    const epigraph = `${title} epigraph about the joy of all the saints and the hope of the world.`;
    const citation =
      title === "GOSPEL"
        ? "A reading from the Holy Gospel according to Matthew 5:1-12"
        : "A reading from the Book of Revelation 7:2-4, 9-14";
    blocks.push({ ...block(title, "reading", [body]), epigraph, citation });
    expected.push(body, epigraph);
  }
  for (const title of [
    "OPENING PRAYER",
    "PRAYER OVER THE OFFERINGS",
    "PRAYER AFTER COMMUNION",
    "GOSPEL ACCLAMATION",
  ]) {
    const text = longText(title.replaceAll(" ", "_"));
    blocks.push(
      block(
        title,
        title === "GOSPEL ACCLAMATION" ? "gospel_acclamation" : "prayer",
        [text],
      ),
    );
    expected.push(text);
  }
  for (const song of ["ENTRANCE", "OFFERTORY", "COMMUNION", "CLOSING"]) {
    const title = longText(`${song}_SONG`);
    blocks.push(block(`${song} SONG: ${title}`, "song"));
    expected.push(title);
  }
  const refrain = "Lord, this is the people that longs to see your face.";
  const verses = [longText("PSALM_ONE"), longText("PSALM_TWO")];
  blocks.push({
    ...block("RESPONSORIAL PSALM", "responsorial_psalm", verses),
    refrain,
  });
  expected.push(...verses);
  const kyrieText = longText("KYRIE");
  blocks.push({
    ...block("LORD HAVE MERCY", "dialogue"),
    dialogues: [{ speaker: "MC", text: kyrieText, paragraphIndex: 0 }],
  });
  expected.push(kyrieText);
  const intentions = [
    longText("FAITHFUL_INTRO"),
    longText("FAITHFUL_ONE"),
    longText("FAITHFUL_TWO"),
    "We pray for those intentions offered for the Mass",
    longText("FAITHFUL_OUTRO"),
  ];
  blocks.push({
    ...block("PRAYER OF THE FAITHFUL", "intercessions"),
    paragraphs: ["PRAYER OF THE FAITHFUL", ...intentions].map((rawText) => ({
      rawText,
      runs: [],
      isBold: false,
      isItalic: false,
      isUnderline: false,
    })),
  });
  expected.push(...intentions.filter((text) => text.includes("_START")));
  const homily = [
    longText("HOMILY_ONE"),
    "Holiness is real in patience.\nPeace & joy begin in small things, even <ordinary> things.",
    longText("HOMILY_TWO"),
  ];
  expected.push(...homily);
  const result = await slides(
    await generateEnglishMassPptxBuffer(input(blocks, homily)),
  );
  // Soft wraps, including an oversized word split across slides, must not invent
  // source characters. Paragraph and hard-break semantics are covered separately.
  const text = result
    .map((slide) => slide.text)
    .join("")
    .replace(/\s+/g, "");
  for (const value of expected) {
    assert.ok(
      text.includes(value.replace(/\s+/g, "")),
      `all source characters remain in order for ${value.slice(0, 60)}`,
    );
    const marker = value.match(/^\w+_START/)?.[0];
    if (marker)
      assert.equal(
        text.split(marker).length - 1,
        1,
        `${marker} is not duplicated on continuation slides`,
      );
  }
  const firstReading = result.findIndex((slide) =>
    slide.text.includes("FIRST_READING_START"),
  );
  const secondReading = result.findIndex((slide) =>
    slide.text.includes("SECOND_READING_START"),
  );
  const gospel = result.findIndex((slide) =>
    slide.text.includes("GOSPEL_START"),
  );
  const homilyStart = result.findIndex((slide) =>
    slide.text.includes("HOMILY_ONE_START"),
  );
  const creed = result.findIndex((slide) =>
    slide.text.includes("NICENE CREED"),
  );
  assert.ok(
    firstReading < secondReading &&
      secondReading < gospel &&
      gospel < homilyStart &&
      homilyStart < creed,
    "continuations retain liturgical order",
  );
});

test("omitting homily clears its text slot without moving the creed", async () => {
  const result = await slides(await generateEnglishMassPptxBuffer(input([])));
  const homilyIndex = result.findIndex((slide) =>
    /\bHOMILY\b/.test(slide.text),
  );
  const creedIndex = result.findIndex((slide) =>
    slide.text.includes("NICENE CREED"),
  );
  assert.ok(homilyIndex >= 0 && creedIndex > homilyIndex);
  assert.ok(
    result
      .slice(homilyIndex + 1, creedIndex)
      .every((slide) => !/\(\s*teks\s*\)/i.test(slide.text)),
    "no homily placeholder leaks into the download",
  );
});

test("a long Gospel citation continues before its response and scripture body", async () => {
  const citation = `A reading from the Holy Gospel ${"according to Matthew and the witness of the saints ".repeat(16).trim()}`;
  const body = "The Gospel body follows its introduction.";
  const gospel = {
    ...block("GOSPEL", "reading", [body]),
    citation,
    epigraph: "Blessed are the peacemakers.",
  };
  const result = await slides(
    await generateEnglishMassPptxBuffer(input([gospel])),
  );
  const text = normalize(result.map((slide) => slide.text).join(" "));
  const source = normalize(
    citation.replace(/^A reading from the Holy Gospel\s*/i, ""),
  );
  assert.ok(
    text.includes(source),
    "the entire citation remains before its response",
  );
  const gospelIndex = text.indexOf("GOSPEL MC:");
  const responseIndex = text.indexOf("Glory to you, Lord.", gospelIndex);
  const citationEnd = text.indexOf(source, gospelIndex) + source.length;
  assert.ok(responseIndex >= citationEnd && text.indexOf(body) > responseIndex);
});

test("short reading introductions, Alleluia, and psalm keep the template gaps", async () => {
  const blocks: EnglishMassBlock[] = [
    {
      ...block("FIRST READING", "reading", ["Reading body."]),
      epigraph: "God is love.",
      citation: "A reading from the Book of Wisdom",
    },
    {
      ...block("SECOND READING", "reading", ["Reading body."]),
      epigraph: "We shall see God.",
      citation: "A reading from the Letter of John",
    },
    {
      ...block("RESPONSORIAL PSALM", "responsorial_psalm", [
        "Give thanks.",
        "Sing praise.",
      ]),
      refrain: "Praise the Lord.",
    },
    block("GOSPEL ACCLAMATION", "gospel_acclamation", ["Come to me."]),
  ];
  const generated = await slides(
    await generateEnglishMassPptxBuffer(input(blocks)),
  );
  const template = await slides(
    await fs.readFile(
      path.join(process.cwd(), "app/template/EM_Template.pptx"),
    ),
  );
  const emptyParagraphs = (xml: string) =>
    [...xml.matchAll(/<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g)].filter(
      (match) =>
        ![...match[0].matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
          .map((text) => text[1])
          .join("")
          .trim(),
    ).length;
  for (const heading of [
    "FIRST READING",
    "SECOND READING",
    "GOSPEL ACCLAMATION",
    "RESPONSORIAL PSALM",
  ]) {
    const original = template.find((slide) => slide.text.includes(heading))!;
    const output = generated.find((slide) => slide.text.includes(heading))!;
    assert.equal(
      emptyParagraphs(output.xml),
      emptyParagraphs(original.xml),
      `${heading} retains its original gaps when the content fits`,
    );
  }
  const psalmContinuation = generated.find((slide) =>
    slide.text.includes("Sing praise."),
  )!;
  const originalContinuation = template.find(
    (slide, index) =>
      index >
        template.findIndex((entry) =>
          entry.text.includes("RESPONSORIAL PSALM"),
        ) && /\(\s*teks\s*\)/i.test(slide.text),
  )!;
  assert.equal(
    emptyParagraphs(psalmContinuation.xml),
    emptyParagraphs(originalContinuation.xml),
  );
});
