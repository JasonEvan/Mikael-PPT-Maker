import { NextRequest, NextResponse } from "next/server";
import { MappedLiturgyData, selectRelevantSections, DocxParseResult } from "@/app/utils/docxParser";
import { generatePPTXBuffer } from "@/app/services/pptxGenerator";
import {
  generatePPTXOBSBuffer,
  generatePPTXOBSKuningBuffer,
} from "@/app/services/pptxGeneratorOBS";

type TemplateId = "mingguan" | "obs" | "obs-kuning";

// ponytail: API route to generate and download filled PPTX
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const template: TemplateId =
      body.template === "mingguan" || body.template === "obs-kuning"
        ? body.template
        : "obs";

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

    const pptxBuffer =
      template === "mingguan"
        ? await generatePPTXBuffer(mappedData)
        : template === "obs-kuning"
          ? await generatePPTXOBSKuningBuffer(mappedData)
          : await generatePPTXOBSBuffer(mappedData);

    const templateDownloadName =
      template === "mingguan"
        ? "Mingguan"
        : template === "obs-kuning"
          ? "OBS_Kuning"
          : "OBS_Hijau";

    return new Response(new Uint8Array(pptxBuffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="Slide_Liturgi_Misa_${templateDownloadName}.pptx"`,
        "Content-Length": pptxBuffer.length.toString(),
      },
    });
  } catch (err: unknown) {
    console.error("PPTX generation error:", err);
    return NextResponse.json(
      {
        error:
          err instanceof Error
            ? err.message
            : "Gagal membuat berkas PowerPoint.",
      },
      { status: 500 }
    );
  }
}
