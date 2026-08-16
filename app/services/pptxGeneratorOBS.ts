import { MappedLiturgyData } from "../utils/docxParser";
import { generatePPTXBufferForTemplate } from "./pptxGenerator";

// OBS Hijau-specific generator. It deliberately shares the established replacement,
// slide-overflow, and notes logic with the weekly generator.
export function generatePPTXOBSBuffer(
  mappedData: MappedLiturgyData,
): Promise<Buffer> {
  return generatePPTXBufferForTemplate(
    mappedData,
    "Template_PPT_OBS_Hijau.pptx",
    {
      doaKolektaMaxChars: 200,
      doaUmatIntentionMaxChars: 210,
      doaSesudahKomuniMaxChars: 210,
      usesPlaceholderOnlyInjilContentSlide: true,
      usesSinglePlaceholderDoaUmatIntentions: true,
    },
  );
}
