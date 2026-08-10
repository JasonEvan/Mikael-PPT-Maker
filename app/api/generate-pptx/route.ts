import { NextRequest, NextResponse } from "next/server";
import { MappedLiturgyData, selectRelevantSections, DocxParseResult } from "@/app/utils/docxParser";
import { generatePPTXBuffer } from "@/app/services/pptxGenerator";

// ponytail: API route to generate and download filled PPTX
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();

    let mappedData: MappedLiturgyData;
    if (body.mappedData) {
      mappedData = body.mappedData;
    } else if (body.sections || body.rawParagraphs) {
      mappedData = selectRelevantSections(body as DocxParseResult);
    } else {
      return NextResponse.json(
        { error: "Data liturgi tidak ditemukan." },
        { status: 400 }
      );
    }

    const pptxBuffer = await generatePPTXBuffer(mappedData);

    return new Response(new Uint8Array(pptxBuffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": 'attachment; filename="Slide_Liturgi_Misa.pptx"',
        "Content-Length": pptxBuffer.length.toString(),
      },
    });
  } catch (err: any) {
    console.error("PPTX generation error:", err);
    return NextResponse.json(
      { error: err?.message || "Gagal membuat berkas PowerPoint." },
      { status: 500 }
    );
  }
}
