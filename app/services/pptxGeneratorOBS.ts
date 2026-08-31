import { MappedLiturgyData } from "../utils/docxParser";
import { generatePPTXBufferForTemplate } from "./pptxGenerator";

// OBS Hijau-specific generator. It deliberately shares the established replacement,
// slide-overflow, and notes logic with the weekly generator.
export function generatePPTXOBSBuffer(
  mappedData: MappedLiturgyData,
): Promise<Buffer> {
  return generatePPTXOBSBufferForTemplate(
    mappedData,
    "Template_PPT_OBS_Hijau.pptx",
  );
}

// OBS Kuning uses the same slide structure and replacement rules as OBS Hijau.
export function generatePPTXOBSKuningBuffer(
  mappedData: MappedLiturgyData,
): Promise<Buffer> {
  return generatePPTXOBSBufferForTemplate(
    mappedData,
    "Template_PPT_OBS_Kuning.pptx",
  );
}

function generatePPTXOBSBufferForTemplate(
  mappedData: MappedLiturgyData,
  templateFileName: string,
): Promise<Buffer> {
  return generatePPTXBufferForTemplate(
    mappedData,
    templateFileName,
    {
      doaKolektaMaxChars: 200,
      doaUmatIntentionMaxChars: 210,
      doaSesudahKomuniMaxChars: 210,
      usesPlaceholderOnlyInjilContentSlide: true,
      omitsGospelCitationPrefix: true,
      usesSinglePlaceholderDoaUmatIntentions: true,
      usesPlaceholderOnlySeruanTobatSlides: true,
    },
  );
}
