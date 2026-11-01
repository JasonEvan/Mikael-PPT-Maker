import { MappedLiturgyData } from "../utils/docxParser";
import { generatePPTXBufferForTemplate } from "./pptxGenerator";

// OBS Hijau uses the shared font-measured Indonesian generator.
export function generatePPTXOBSBuffer(
  mappedData: MappedLiturgyData,
): Promise<Buffer> {
  return generatePPTXOBSBufferForTemplate(
    mappedData,
    "Template_PPT_OBS_Hijau.pptx",
  );
}

// OBS Kuning has its own layout, with the same anonymous content-role profile.
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
  return generatePPTXBufferForTemplate(mappedData, templateFileName, {
    usesPlaceholderOnlyInjilContentSlide: true,
    omitsGospelCitationPrefix: true,
    usesSinglePlaceholderDoaUmatIntentions: true,
    usesPlaceholderOnlySeruanTobatSlides: true,
  });
}
