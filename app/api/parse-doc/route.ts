import WordExtractor from "word-extractor";
import {
  createParseResult,
  LiturgyParagraph,
} from "@/app/utils/docxParser";

function isLikelyHeading(text: string): boolean {
  const letters = text.match(/[A-Za-zÀ-ÿ]/g)?.join("") ?? "";
  return (
    letters.length >= 2 &&
    letters === letters.toUpperCase() &&
    text.length <= 100
  );
}

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || typeof file === "string" || !file.name.toLowerCase().endsWith(".doc")) {
      return Response.json({ error: "Unggah berkas Microsoft Word bertipe .doc." }, { status: 400 });
    }

    const extractor = new WordExtractor();
    const document = await extractor.extract(Buffer.from(await file.arrayBuffer()));
    const rawParagraphs: LiturgyParagraph[] = document
      .getBody()
      .split(/\r\n|\r|\n/)
      .map((line) => line.replace(/\u0000/g, "").trim())
      .filter(Boolean)
      .map((rawText) => {
        const isBold = isLikelyHeading(rawText);
        return {
          rawText,
          runs: [{ text: rawText, bold: isBold, underline: false, italic: false }],
          isBold,
          isUnderline: false,
          isItalic: false,
        };
      });

    if (rawParagraphs.length === 0) {
      return Response.json({ error: "Tidak ada teks yang dapat dibaca dari berkas .doc." }, { status: 422 });
    }

    return Response.json(createParseResult(file.name, rawParagraphs));
  } catch (error) {
    console.error("Legacy .doc parse error:", error);
    return Response.json(
      { error: "Gagal membaca berkas .doc. Pastikan berkas tidak rusak atau dilindungi." },
      { status: 422 },
    );
  }
}
