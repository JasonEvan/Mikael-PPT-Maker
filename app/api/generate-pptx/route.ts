import { NextRequest, NextResponse } from "next/server";
import { MappedLiturgyData, selectRelevantSections, DocxParseResult } from "@/app/utils/docxParser";
import { generatePPTXBuffer } from "@/app/services/pptxGenerator";
import {
  generatePPTXOBSBuffer,
  generatePPTXOBSKuningBuffer,
  generatePPTXOBSUnguBuffer,
} from "@/app/services/pptxGeneratorOBS";
import { generateEnglishMassPptxBuffer } from "@/app/services/englishMassPptxGenerator";
import { EnglishMassParseResult } from "@/app/utils/englishMassParser";

export const runtime = "nodejs";

function pptxDownload(buffer: Buffer, filename: string): Response {
  let offset = 0;
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      const end = Math.min(offset + 64 * 1024, buffer.length);
      if (offset < end) controller.enqueue(buffer.subarray(offset, end));
      offset = end;
      if (offset === buffer.length) controller.close();
    },
  });
  return new Response(stream, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": `attachment; filename="${filename}"`,
    },
  });
}

type TemplateId = "mingguan" | "obs" | "obs-kuning" | "obs-ungu";

// ponytail: API route to generate and download filled PPTX
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    if (body.englishMass?.massType === "english") {
      const pptxBuffer = await generateEnglishMassPptxBuffer(body.englishMass as EnglishMassParseResult);
      return pptxDownload(pptxBuffer, "Slide_Liturgi_English_Mass.pptx");
    }

    const template: TemplateId =
      body.template === "mingguan" ||
      body.template === "obs-kuning" ||
      body.template === "obs-ungu"
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
          : template === "obs-ungu"
            ? await generatePPTXOBSUnguBuffer(mappedData)
            : await generatePPTXOBSBuffer(mappedData);

    const templateDownloadName =
      template === "mingguan"
        ? "Mingguan"
        : template === "obs-kuning"
          ? "OBS_Kuning"
          : template === "obs-ungu"
            ? "OBS_Ungu"
            : "OBS_Hijau";

    return pptxDownload(
      pptxBuffer,
      `Slide_Liturgi_Misa_${templateDownloadName}.pptx`,
    );
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
