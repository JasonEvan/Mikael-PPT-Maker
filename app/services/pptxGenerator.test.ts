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
  generatePPTXOBSUnguBuffer,
} from "./pptxGeneratorOBS";
import type { MappedLiturgyData } from "../utils/docxParser";

type Node = Record<string, unknown>;
type Slide = { file: string; xml: string; nodes: Node[]; text: string };
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

const obsGenerators = [
  ["Hijau", generatePPTXOBSBuffer],
  ["Kuning", generatePPTXOBSKuningBuffer],
  ["Ungu", generatePPTXOBSUnguBuffer],
] as const;

async function obsTemplate(color: string) {
  const source = await presentation(
    await fs.readFile(
      path.join(process.cwd(), `app/template/Template_PPT_OBS_${color}.pptx`),
    ),
  );
  const placeholders = (slide: (typeof source.slides)[number]) =>
    find(slide.nodes, "a:p").filter((p) => /\(\s*teks\s*\)/i.test(text([p])))
      .length;
  const seruanTitle = source.slides.findIndex((slide) =>
    slide.text.includes("SERUAN TOBAT"),
  );
  const seruan = source.slides
    .slice(seruanTitle + 1)
    .filter((slide) => placeholders(slide) === 1)
    .slice(0, 3);
  const umatTitle = source.slides.findIndex((slide) =>
    slide.text.includes("DOA UMAT"),
  );
  const umat = source.slides
    .slice(umatTitle + 1)
    .filter((slide) => placeholders(slide) === 1)
    .slice(0, 3);
  const gospelIntro = source.slides.findIndex((slide) =>
    /Inilah Injil suci menurut/i.test(slide.text),
  );
  return {
    ...source,
    seruan,
    umat,
    metadata: [
      source.slides.find(
        (slide) => slide.text.includes("BACAAN PERTAMA") && placeholders(slide),
      )!,
      source.slides.find(
        (slide) => slide.text.includes("BACAAN KEDUA") && placeholders(slide),
      )!,
      source.slides
        .slice(gospelIntro + 1)
        .find((slide) => placeholders(slide) === 2)!,
    ],
  };
}

async function slideNotes(zip: JSZip, file: string) {
  const relFile = zip.file(
    path.posix.join(
      path.posix.dirname(file),
      "_rels",
      `${path.posix.basename(file)}.rels`,
    ),
  );
  if (!relFile) return undefined;
  const rel = find(
    parser.parse(await relFile.async("text")) as Node[],
    "Relationship",
  ).find((node) => attributes(node)["@_Type"].endsWith("/notesSlide"));
  if (!rel) return undefined;
  const target = path.posix.normalize(
    path.posix.join(path.posix.dirname(file), attributes(rel)["@_Target"]),
  );
  return text(parser.parse(await zip.file(target)!.async("text")) as Node[]);
}

function obsDynamicText(
  slides: readonly Slide[],
  source: { slides: readonly Slide[] },
): string {
  const editedShapes = new Map(
    source.slides.map((slide) => [
      attributes(find(slide.nodes, "p14:creationId")[0])["@_val"],
      find(slide.nodes, "p:sp").flatMap((shape, index) =>
        /\(\s*teks\s*\)/i.test(text([shape])) ? [index] : [],
      ),
    ]),
  );
  return slides
    .flatMap((slide) => {
      const indices =
        editedShapes.get(
          attributes(find(slide.nodes, "p14:creationId")[0])["@_val"],
        ) ?? [];
      const shapes = find(slide.nodes, "p:sp");
      return indices.map((index) => text([shapes[index]]));
    })
    .join("");
}

test("all OBS exports require provisioned Indonesian fonts rather than silently substituting", async () => {
  const previous = process.env.INDONESIAN_MASS_FONT_DIR;
  process.env.INDONESIAN_MASS_FONT_DIR =
    "/nonexistent-licensed-indonesian-fonts";
  try {
    for (const [, generate] of obsGenerators) {
      await assert.rejects(
        generate(mapped()),
        /Indonesian Mass layout requires the exact licensed .*INDONESIAN_MASS_FONT_DIR.*Font substitution is disabled/,
      );
    }
  } finally {
    if (previous === undefined) delete process.env.INDONESIAN_MASS_FONT_DIR;
    else process.env.INDONESIAN_MASS_FONT_DIR = previous;
  }
});

for (const [color, generate] of obsGenerators) {
  test(`OBS ${color} measures every role, keeps complete source/order and original scripture notes ownership`, async () => {
    const data = mapped(true);
    data.bacaanKedua!.epigraph = `EPIGRAF_KEDUA_AWAL ${"Kita akan melihat Allah dalam kasih. ".repeat(18)} EPIGRAF_KEDUA_AKHIR`;
    data.bacaanInjil!.epigraph = `EPIGRAF_INJIL_AWAL ${"Berbahagialah kamu yang membawa damai. ".repeat(18)} EPIGRAF_INJIL_AKHIR`;
    const source = await obsTemplate(color);
    const { zip, slides } = await presentation(await generate(data));
    const visible = compact(obsDynamicText(slides, source));
    const values = [
      ...data.seruanTobat!,
      data.doaKolekta!,
      data.bacaanPertama!.epigraph!,
      data.bacaanPertama!.citation!,
      ...data.mazmurTanggapan!.verses,
      data.bacaanKedua!.epigraph!,
      data.bacaanKedua!.citation!,
      data.bacaanInjil!.epigraph!,
      "Matius 5:1-12",
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
    let previous = -1;
    for (const value of values) {
      const normalized = compact(value);
      const start = visible.indexOf(normalized);
      assert.ok(
        start > previous,
        `complete source in liturgical order: ${value.slice(0, 45)}`,
      );
      previous = start;
      const marker = value.match(/^\w+_AWAL/)?.[0];
      if (marker)
        assert.equal(
          visible.split(marker).length - 1,
          1,
          `${marker} appears once`,
        );
    }
    assert.ok(
      !visible.includes("InjilSucimenurutMatius5:1-12"),
      "OBS citation omits its already spoken prefix",
    );
    const intro = slides.find((slide) =>
      /Inilah Injil suci menurut/i.test(slide.text),
    )!;
    assert.ok(
      intro.text.includes("Matius"),
      "Gospel name fills the intro rather than metadata",
    );
    for (let index = 0; index < source.seruan.length; index++) {
      const original = source.seruan[index];
      const response =
        source.slides[
          source.slides.findIndex((slide) => slide.file === original.file) + 1
        ];
      const responseIndex = slides.findIndex(
        (slide) => slide.file === response.file,
      );
      const endIndex = slides.findIndex((slide) =>
        slide.text.includes(`SERUAN_${["SATU", "DUA", "TIGA"][index]}_AKHIR`),
      );
      assert.ok(
        endIndex < responseIndex,
        "separate fixed Seruan response follows every invocation continuation",
      );
      assert.equal(
        slides[responseIndex].xml,
        response.xml,
        "fixed response slide stays untouched",
      );
      if (index < 2) {
        assert.ok(
          responseIndex <
            slides.findIndex(
              (slide) => slide.file === source.seruan[index + 1].file,
            ),
        );
      }
    }
    for (const [index, reading] of [
      data.bacaanPertama!,
      data.bacaanKedua!,
      data.bacaanInjil!,
    ].entries()) {
      assert.ok(
        !visible.includes(compact(reading.notesBody!)),
        "scripture is notes-only",
      );
      assert.ok(
        compact((await slideNotes(zip, source.metadata[index].file))!).includes(
          compact(reading.notesBody!),
        ),
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
      assert.equal(
        notes.filter((note) =>
          compact(note).includes(compact(reading.notesBody!)),
        ).length,
        1,
      );
    }
    const originalFiles = new Set(source.slides.map((slide) => slide.file));
    for (const clone of slides.filter(
      (slide) => !originalFiles.has(slide.file),
    )) {
      assert.equal(
        await slideNotes(zip, clone.file),
        undefined,
        "continuations never share original notes",
      );
    }
    const ids = find(
      parser.parse(
        await zip.file("ppt/presentation.xml")!.async("text"),
      ) as Node[],
      "p:sldId",
    );
    for (const key of ["@_id", "@_r:id"]) {
      assert.equal(
        new Set(ids.map((node) => attributes(node)[key])).size,
        ids.length,
      );
    }
    const firstIntentionEnd = slides.findIndex((slide) =>
      slide.text.includes("INTENSI_SATU_AKHIR"),
    );
    const firstResponse = slides.findIndex((slide) =>
      slide.text.includes("Kabulkanlah permohonan pertama"),
    );
    const nextIntention = slides.findIndex((slide) =>
      slide.text.includes("Bagi keluarga"),
    );
    assert.ok(
      firstIntentionEnd <= firstResponse && firstResponse < nextIntention,
      "each intention response follows its whole body before the next intention",
    );
    assert.ok(
      !slides[firstResponse].text.includes("Bagi keluarga"),
      "next intention never packs into response space",
    );
  });

  test(`OBS ${color} preserves inline emphasis seams, title breaks, token styling and regular U-tab responses`, async () => {
    const data = mapped();
    data.doaUmat!.intentions = [
      {
        title: "",
        body: "Bagi Gereja. Semoga damai.",
        response: "Dengarkanlah kami.",
      },
      {
        title: "Bagi keluarga",
        body: "Semoga rukun.",
        response: "Bimbinglah kami.",
      },
    ];
    const source = await obsTemplate(color);
    const { slides } = await presentation(await generate(data));
    const token = find(source.umat[1].nodes, "a:p").find((p) =>
      /\(\s*teks\s*\)/i.test(text([p])),
    )!;
    const base = find([token], "a:rPr")[0];
    for (const [label, remainder, responseText] of [
      ["Bagi Gereja.", " Semoga damai.", "U\tDengarkanlah kami."],
      ["Bagi keluarga", "Semoga rukun.", "U\tBimbinglah kami."],
    ]) {
      const page = slides.find((slide) => slide.text.includes(label))!;
      const runs = find(page.nodes, "a:r");
      const combined = runs.map((run) => text([run])).join("");
      const assertStyle = (value: string, bold: string) => {
        const start = combined.indexOf(value);
        assert.ok(start >= 0, `missing intention text: ${value}`);
        let offset = 0;
        for (const run of runs) {
          const end = offset + text([run]).length;
          if (end > start && offset < start + value.length) {
            const properties = find([run], "a:rPr")[0];
            assert.equal(attributes(properties)["@_b"], bold);
            assert.equal(
              attributes(properties)["@_sz"],
              attributes(base)["@_sz"],
            );
            assert.equal(
              attributes(properties)["@_i"] ?? "0",
              attributes(base)["@_i"] ?? "0",
            );
            assert.deepEqual(
              find([properties], "a:solidFill"),
              find([base], "a:solidFill"),
            );
            assert.deepEqual(
              find([properties], "a:latin"),
              find([base], "a:latin"),
            );
          }
          offset = end;
        }
      };
      assertStyle(label, "1");
      assertStyle(remainder.trim(), "0");
      assertStyle(responseText, "0");
      const bodyParagraph = find(page.nodes, "a:p").find((p) =>
        text([p]).includes(label),
      )!;
      if (!data.doaUmat!.intentions.find((item) => item.title === label)) {
        assert.ok(
          text([bodyParagraph]).includes(`${label}${remainder}`),
          "inline sentence emphasis retains its separating whitespace",
        );
      }
    }
  });

  test(`OBS ${color} paginates an oversized single-token response only after the complete body`, async () => {
    const data = mapped();
    const response = `RESPONS_RAKSASA_AWAL ${"Dengarkanlah dan bimbinglah kami dalam kasih. ".repeat(90)} RESPONS_RAKSASA_AKHIR`;
    const source = await obsTemplate(color);
    data.doaUmat!.intentions = [
      {
        title: "Untuk sesama",
        body: "Doa ini berakhir pada BATAS_TUBUH.",
        response,
      },
      { title: "INTENSI_BERIKUTNYA", body: "Semoga damai." },
    ];
    const { slides } = await presentation(await generate(data));
    const visible = compact(obsDynamicText(slides, source));
    assert.ok(
      visible.includes(compact(response)),
      "oversized response loses no source characters",
    );
    assert.ok(
      visible.indexOf("BATAS_TUBUH.") < visible.indexOf("RESPONS_RAKSASA_AWAL"),
    );
    assert.ok(
      visible.indexOf("RESPONS_RAKSASA_AKHIR") <
        visible.indexOf("INTENSI_BERIKUTNYA"),
    );
    const responseStart = slides.findIndex((slide) =>
      slide.text.includes("RESPONS_RAKSASA_AWAL"),
    );
    const responseEnd = slides.findIndex((slide) =>
      slide.text.includes("RESPONS_RAKSASA_AKHIR"),
    );
    assert.ok(
      responseEnd > responseStart,
      "response itself spans measured slides",
    );
    assert.equal(
      slides
        .slice(responseStart, responseEnd + 1)
        .map((slide) => slide.text)
        .join("")
        .split("U\t").length - 1,
      1,
    );
    for (const slide of slides.slice(responseStart, responseEnd + 1)) {
      for (const run of find(slide.nodes, "a:r").filter((run) =>
        /Dengarkanlah|RESPONS_RAKSASA/.test(text([run])),
      )) {
        assert.equal(attributes(find([run], "a:rPr")[0])["@_b"], "0");
      }
    }
  });

  test(`OBS ${color} consumes empty and missing anonymous roles without repeating Seruan or stealing later prayers`, async () => {
    const source = await obsTemplate(color);
    for (const umat of [
      undefined,
      { intro: "", intentions: [], closing: "" },
      {
        intro: "",
        intentions: [{ title: "", body: "INTENSI_TANPA_PENGANTAR" }],
        closing: "PENUTUP_SETELAH_KOSONG",
      },
    ]) {
      const data: MappedLiturgyData = {
        unmatchedSections: [],
        seruanTobat: ["SERUAN_TUNGGAL"],
        mazmurTanggapan: { verses: [] },
        doaUmat: umat,
        doaAtasPersembahan: "PERSEMBAHAN_TETAP",
        doaSesudahKomuni: "KOMUNI_TETAP",
      };
      const { slides } = await presentation(await generate(data));
      const visible = slides.map((slide) => slide.text).join("");
      assert.equal(visible.split("SERUAN_TUNGGAL").length - 1, 1);
      for (const original of source.seruan.slice(1)) {
        assert.ok(
          !/\(\s*teks\s*\)/i.test(
            slides.find((slide) => slide.file === original.file)!.text,
          ),
          "missing Seruan entry is cleared, not repeated",
        );
      }
      for (const original of [...source.umat, ...source.metadata]) {
        assert.ok(
          !/\(\s*teks\s*\)/i.test(
            slides.find((slide) => slide.file === original.file)!.text,
          ),
          "empty anonymous role is consumed",
        );
      }
      const offering = slides.find((slide) =>
        slide.text.includes("PERSEMBAHAN_TETAP"),
      )!;
      const communion = slides.find((slide) =>
        slide.text.includes("KOMUNI_TETAP"),
      )!;
      assert.ok(offering.text.includes("ATAS PERSEMBAHAN"));
      assert.ok(communion.text.includes("SESUDAH KOMUNI"));
      if (umat?.intentions.length) {
        assert.ok(visible.includes("INTENSI_TANPA_PENGANTAR"));
        assert.ok(visible.includes("PENUTUP_SETELAH_KOSONG"));
      }
      const firstReading = source.slides.findIndex(
        (slide) => slide.file === source.metadata[0].file,
      );
      const secondReading = source.slides.findIndex(
        (slide) => slide.file === source.metadata[1].file,
      );
      const psalm = source.slides
        .slice(firstReading + 1, secondReading)
        .find((slide) => /\(\s*teks\s*\)/i.test(slide.text))!;
      assert.equal(
        slides.find((slide) => slide.file === psalm.file)!.xml,
        psalm.xml,
        "zero verses preserve the original psalm placeholder",
      );
      const untouched = source.slides.filter(
        (slide) =>
          /\(\s*teks\s*\)/i.test(slide.text) &&
          ![...source.seruan, ...source.umat, ...source.metadata].some(
            (role) => role.file === slide.file,
          ) &&
          !/KOLEKTA|PERSEMBAHAN|KOMUNI|Inilah Injil/i.test(slide.text) &&
          slide.file !== psalm.file,
      );
      for (const original of untouched) {
        assert.equal(
          slides.find((slide) => slide.file === original.file)!.xml,
          original.xml,
          "unrelated placeholders remain untouched",
        );
      }
    }
  });
}

test("OBS Gospel lower-thirds retain three fitting lines without splitting metadata or its response", async () => {
  for (const [color, generate] of obsGenerators) {
    const data = mapped();
    data.bacaanInjil = {
      gospelName: "Matius 5:1-12a",
      epigraph:
        "Bersukacita dan bergembiralah, karena besarlah ganjaranmu di surga",
      citation: "Injil Suci menurut Matius 5:1-12a",
      notesBody: "The full Gospel stays in its original notes.",
    };
    const source = await obsTemplate(color);
    const intro = source.slides.find((slide) =>
      slide.text.includes("Inilah Injil suci menurut"),
    )!;
    const { slides } = await presentation(await generate(data));
    for (const [original, expected] of [
      [
        intro,
        ["Inilah Injil suci menurut Matius 5:1-12a", "Dimuliakanlah † Tuhan"],
      ],
      [source.metadata[2], [data.bacaanInjil.epigraph!, "Matius 5:1-12a"]],
    ] as const) {
      const origin = attributes(find(original.nodes, "p14:creationId")[0])[
        "@_val"
      ];
      const pages = slides.filter(
        (slide) =>
          attributes(find(slide.nodes, "p14:creationId")[0])["@_val"] ===
          origin,
      );
      assert.equal(
        pages.length,
        1,
        `${color}: the native three-line lower-third fits`,
      );
      for (const value of expected)
        assert.ok(compact(pages[0].text).includes(compact(value)));
    }
  }
});

test("OBS wrapped intention titles end their justified paragraph before the regular body", async () => {
  const title = "Bagi kita yang berkumpul di sekitar altar.";
  const body =
    "Ya Bapa, kuduskanlah kami yang berkumpul mengelilingi meja perjamuan-Mu ini.";
  for (const [, generate] of obsGenerators) {
    const data = mapped();
    data.doaUmat!.intentions = [
      { title, body, response: "Kuduskanlah kami, ya Tuhan." },
    ];
    const { slides } = await presentation(await generate(data));
    const page = slides.find((slide) => slide.text.includes("Bagi kita"))!;
    const paragraphs = find(page.nodes, "a:p");
    const heading = paragraphs.find((p) => text([p]) === title)!;
    const content = paragraphs.find((p) => text([p]).includes(body))!;
    assert.ok(
      heading,
      "the title is a complete paragraph, so its last line is not force-justified",
    );
    assert.ok(
      content && content !== heading,
      "the regular body starts its own paragraph",
    );
    assert.equal(attributes(find([heading], "a:pPr")[0])["@_algn"], "just");
    assert.equal(attributes(find([content], "a:pPr")[0])["@_algn"], "just");
  }
});
