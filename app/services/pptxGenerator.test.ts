import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { generatePPTXBuffer } from "./pptxGenerator";
import {
  generatePPTXOBSBuffer,
  generatePPTXOBSKuningBuffer,
} from "./pptxGeneratorOBS";
import type { MappedLiturgyData } from "../utils/docxParser";

type Node = Record<string, unknown>;
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  trimValues: false,
  parseTagValue: false,
});
const compact = (value: string) => value.replace(/\s+/g, "");
function find(nodes: Node[], name: string): Node[] {
  return nodes.flatMap((node) =>
    Object.entries(node).flatMap(([key, value]) =>
      key === ":@"
        ? []
        : [
            ...(key === name ? [node] : []),
            ...(Array.isArray(value) ? find(value as Node[], name) : []),
          ],
    ),
  );
}
function attributes(node: Node | undefined): Record<string, string> {
  return (node?.[":@"] ?? {}) as Record<string, string>;
}
function text(nodes: Node[]): string {
  return find(nodes, "a:t")
    .map((node) =>
      (node["a:t"] as Node[])
        .map((entry) => String(entry["#text"] ?? ""))
        .join(""),
    )
    .join("");
}
function mapped(expand = false): MappedLiturgyData {
  const body = (label: string) =>
    `${label}_AWAL ${expand ? "Semoga kami setia dalam kasih, menolong sesama, dan membawa damai bagi keluarga. ".repeat(13) : "Kami mohon rahmat-Mu."} ${label}_AKHIR`;
  return {
    unmatchedSections: [],
    seruanTobat: [body("SERUAN_SATU"), body("SERUAN_DUA"), body("SERUAN_TIGA")],
    doaKolekta: body("KOLEKTA"),
    bacaanPertama: {
      epigraph: expand ? body("EPIGRAF_PERTAMA") : "Kasih itu sabar.",
      citation: "Bacaan dari Kitab Wahyu 7:2-4",
      notesBody:
        "NOTES_PERTAMA\n\nTubuh bacaan pertama tetap berada dalam catatan pembaca.",
    },
    bacaanKedua: {
      epigraph: "Kita akan melihat Allah.",
      citation: "Bacaan dari Surat Yohanes 3:1-3",
      notesBody:
        "NOTES_KEDUA\n\nTubuh bacaan kedua tetap berada dalam catatan pembaca.",
    },
    bacaanInjil: {
      gospelName: "Matius",
      epigraph: "Berbahagialah kamu.",
      citation: "Injil Suci menurut Matius 5:1-12",
      notesBody:
        "NOTES_INJIL\n\nTubuh bacaan Injil tetap berada dalam catatan pembaca.",
    },
    mazmurTanggapan: { verses: [body("MAZMUR_SATU"), body("MAZMUR_DUA")] },
    doaUmat: {
      intro: body("PENGANTAR_UMAT"),
      intentions: [
        {
          title: "Bagi Gereja",
          body: body("INTENSI_SATU"),
          response: "Kabulkanlah permohonan pertama kami.",
        },
        {
          title: "Bagi keluarga",
          body: body("INTENSI_DUA"),
          response: expand
            ? body("RESPONS_PANJANG")
            : "Kabulkanlah permohonan kedua kami.",
        },
      ],
      closing: body("PENUTUP_UMAT"),
    },
    doaAtasPersembahan: body("PERSEMBAHAN"),
    doaSesudahKomuni: body("SESUDAH_KOMUNI"),
  };
}
async function presentation(buffer: Buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const relNodes = parser.parse(
    await zip.file("ppt/_rels/presentation.xml.rels")!.async("text"),
  ) as Node[];
  const targets = new Map(
    find(relNodes, "Relationship").map((node) => [
      attributes(node)["@_Id"],
      attributes(node)["@_Target"],
    ]),
  );
  const presentationNodes = parser.parse(
    await zip.file("ppt/presentation.xml")!.async("text"),
  ) as Node[];
  const slides = await Promise.all(
    find(presentationNodes, "p:sldId").map(async (node) => {
      const target = targets.get(attributes(node)["@_r:id"]);
      assert.ok(target, "presentation slide relationship exists");
      const file = path.posix.normalize(`ppt/${target}`);
      const xml = await zip.file(file)!.async("text");
      assert.equal(XMLValidator.validate(xml), true);
      const nodes = parser.parse(xml) as Node[];
      return { file, xml, nodes, text: text(nodes) };
    }),
  );
  return { zip, slides };
}

test("weekly long text remains complete and ordered while scripture stays in original notes", async () => {
  const data = mapped(true);
  const { zip, slides } = await presentation(await generatePPTXBuffer(data));
  const visible = compact(slides.map((slide) => slide.text).join(""));
  const expected = [
    ...data.seruanTobat!,
    data.doaKolekta!,
    data.bacaanPertama!.epigraph!,
    ...data.mazmurTanggapan!.verses,
    data.doaUmat!.intro!,
    ...data.doaUmat!.intentions.flatMap((item) => [
      item.title,
      item.body,
      item.response!,
    ]),
    data.doaUmat!.closing!,
    data.doaAtasPersembahan!,
    data.doaSesudahKomuni!,
  ];
  for (const value of expected) {
    const valueText = compact(value);
    assert.ok(
      visible.includes(valueText),
      `all source characters remain in order for ${value.slice(0, 60)}`,
    );
    const marker = value.match(/^\w+_AWAL/)?.[0];
    if (marker)
      assert.equal(
        visible.split(marker).length - 1,
        1,
        `${marker} is not duplicated`,
      );
  }
  assert.ok(
    visible.indexOf("MAZMUR_SATU_AWAL") < visible.indexOf("MAZMUR_DUA_AWAL"),
  );
  assert.ok(
    visible.indexOf("INTENSI_SATU_AKHIR") <
      visible.indexOf("Kabulkanlahpermohonanpertamakami."),
  );
  assert.ok(
    visible.indexOf("Kabulkanlahpermohonanpertamakami.") <
      visible.indexOf("INTENSI_DUA_AWAL"),
  );
  assert.ok(
    visible.indexOf("INTENSI_DUA_AKHIR") <
      visible.indexOf("RESPONS_PANJANG_AWAL"),
  );
  const notes = await Promise.all(
    Object.values(zip.files)
      .filter((file) =>
        /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(file.name),
      )
      .map(async (file) =>
        text(parser.parse(await file.async("text")) as Node[]),
      ),
  );
  for (const [heading, reading] of [
    ["BACAAN PERTAMA", data.bacaanPertama!],
    ["BACAAN KEDUA", data.bacaanKedua!],
    ["BACAAN INJIL", data.bacaanInjil!],
  ] as const) {
    const noteText = compact(reading.notesBody!);
    assert.ok(
      !visible.includes(noteText),
      `${heading} scripture is not exposed on slides`,
    );
    assert.equal(
      notes.filter((note) => compact(note).includes(noteText)).length,
      1,
      `${heading} has one complete notes body`,
    );
    const first = slides.find(
      (slide) =>
        slide.text.includes(heading) &&
        (heading !== "BACAAN INJIL" || slide.text.includes("Berbahagialah")),
    )!;
    const part = first.file.match(/slide\d+\.xml$/)![0];
    const rel = await zip.file(`ppt/slides/_rels/${part}.rels`)!.async("text");
    const noteRel = find(parser.parse(rel) as Node[], "Relationship").find(
      (node) => attributes(node)["@_Type"].endsWith("/notesSlide"),
    );
    assert.ok(
      noteRel,
      `${heading} original metadata slide owns the notes relationship`,
    );
    const noteTarget = path.posix.normalize(
      path.posix.join("ppt/slides", attributes(noteRel)["@_Target"]),
    );
    assert.ok(
      compact(
        text(parser.parse(await zip.file(noteTarget)!.async("text")) as Node[]),
      ).includes(noteText),
    );
  }
  const slideIds = find(
    parser.parse(
      await zip.file("ppt/presentation.xml")!.async("text"),
    ) as Node[],
    "p:sldId",
  );
  assert.equal(
    new Set(slideIds.map((node) => attributes(node)["@_id"])).size,
    slideIds.length,
  );
  assert.equal(
    new Set(slideIds.map((node) => attributes(node)["@_r:id"])).size,
    slideIds.length,
  );
});

test("weekly metadata and intentions retain template faces, weights, response style and spaces", async () => {
  const { slides } = await presentation(await generatePPTXBuffer(mapped()));
  const template = await presentation(
    await fs.readFile(
      path.join(process.cwd(), "app/template/Template_Mingguan.pptx"),
    ),
  );
  const first = slides.find((slide) => slide.text.includes("BACAAN PERTAMA"))!;
  const icon = find(first.nodes, "a:r").find((run) =>
    text([run]).includes("\uf026"),
  )!;
  assert.equal(attributes(find([icon], "a:sym")[0])["@_typeface"], "Wingdings");
  assert.equal(attributes(find([icon], "a:rPr")[0])["@_sz"], "6000");
  const epigraph = find(first.nodes, "a:r").find((run) =>
    text([run]).includes("Kasih itu sabar"),
  )!;
  assert.equal(
    attributes(find([epigraph], "a:latin")[0])["@_typeface"],
    "Lucida Sans",
  );
  assert.equal(attributes(find([epigraph], "a:rPr")[0])["@_b"], "1");
  assert.equal(attributes(find([epigraph], "a:rPr")[0])["@_i"], "1");
  const intention = slides.find((slide) =>
    slide.text.includes("INTENSI_SATU_AWAL"),
  )!;
  const title = find(intention.nodes, "a:r").find((run) =>
    text([run]).includes("Bagi Gereja"),
  )!;
  const body = find(intention.nodes, "a:r").find((run) =>
    text([run]).includes("INTENSI_SATU_AWAL"),
  )!;
  const response = find(intention.nodes, "a:r").find((run) =>
    text([run]).includes("Kabulkanlah"),
  )!;
  assert.equal(attributes(find([title], "a:rPr")[0])["@_b"], "1");
  assert.equal(attributes(find([body], "a:rPr")[0])["@_b"], "0");
  assert.equal(attributes(find([response], "a:rPr")[0])["@_i"], "1");
  assert.ok(
    text([response]).startsWith("U\t"),
    "congregational label retains its tab",
  );
  const sourceIntention = template.slides.find(
    (slide) =>
      find(slide.nodes, "a:p").filter((node) =>
        /\(\s*teks\s*\)/i.test(text([node])),
      ).length === 2 && !/BACAAN|INJIL/.test(slide.text),
  )!;
  const sourceResponse = find(sourceIntention.nodes, "a:p").filter((node) =>
    /\(\s*teks\s*\)/i.test(text([node])),
  )[1];
  const sourceFill = find([sourceResponse], "a:solidFill")[0];
  assert.deepEqual(
    find([response], "a:solidFill")[0],
    sourceFill,
    "response keeps the template fill",
  );
  for (const heading of ["BACAAN PERTAMA", "BACAAN KEDUA"]) {
    const output = slides.find((slide) => slide.text.includes(heading))!;
    const original = template.slides.find((slide) =>
      slide.text.includes(heading),
    )!;
    assert.equal(
      find(output.nodes, "a:p").filter((p) => !text([p]).trim()).length,
      find(original.nodes, "a:p").filter((p) => !text([p]).trim()).length,
      `${heading} retains fitting template gaps`,
    );
  }
});

test("both OBS exports remain independent of weekly and English font provisioning", async () => {
  const beforeIndonesian = process.env.INDONESIAN_MASS_FONT_DIR;
  const beforeEnglish = process.env.ENGLISH_MASS_FONT_DIR;
  process.env.INDONESIAN_MASS_FONT_DIR = "/nonexistent-licensed-weekly-fonts";
  process.env.ENGLISH_MASS_FONT_DIR = "/nonexistent-licensed-english-fonts";
  try {
    for (const generate of [
      generatePPTXOBSBuffer,
      generatePPTXOBSKuningBuffer,
    ]) {
      const { zip, slides } = await presentation(await generate(mapped()));
      const visible = compact(slides.map((slide) => slide.text).join(""));
      assert.ok(visible.includes("KOLEKTA_AWAL"));
      assert.ok(visible.includes("Kabulkanlahpermohonanpertamakami."));
      assert.ok(!visible.includes("NOTES_PERTAMA"));
      const noteBodies = await Promise.all(
        Object.values(zip.files)
          .filter((file) =>
            /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(file.name),
          )
          .map(async (file) => file.async("text")),
      );
      assert.ok(noteBodies.some((note) => note.includes("NOTES_PERTAMA")));
    }
  } finally {
    if (beforeIndonesian === undefined)
      delete process.env.INDONESIAN_MASS_FONT_DIR;
    else process.env.INDONESIAN_MASS_FONT_DIR = beforeIndonesian;
    if (beforeEnglish === undefined) delete process.env.ENGLISH_MASS_FONT_DIR;
    else process.env.ENGLISH_MASS_FONT_DIR = beforeEnglish;
  }
});
