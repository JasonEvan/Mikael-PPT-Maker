"use client";

import { useState } from "react";
import { EnglishMassParseResult } from "../utils/englishMassParser";

interface EnglishMassViewerProps {
  parseResult: EnglishMassParseResult;
  onReset?: () => void;
}

const hiddenBlockTitles = new Set([
  "THE PRIEST ADDRESSED THE CONGREGATION",
  "INVITATION TO PRAYER",
  "THE EUCHARISTIC PRAYER",
  "PREFACE",
  "ACCLAMATION",
  "EUCHARISTIC PRAYER",
]);

export default function EnglishMassViewer({ parseResult, onReset }: EnglishMassViewerProps) {
  const [showRaw, setShowRaw] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generationError, setGenerationError] = useState<string | null>(null);

  const handleGenerate = async () => {
    setIsGenerating(true);
    setGenerationError(null);
    try {
      const response = await fetch("/api/generate-pptx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ englishMass: parseResult }),
      });
      if (!response.ok) throw new Error((await response.json()).error || "Unable to generate the presentation.");

      const url = URL.createObjectURL(await response.blob());
      const link = document.createElement("a");
      link.href = url;
      link.download = `Slide_Liturgi_${parseResult.fileName.replace(/\.(?:doc|docx)$/i, "")}_English_Mass.pptx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (error) {
      setGenerationError(error instanceof Error ? error.message : "Unable to generate the presentation.");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="document-viewer w-100 h-100 d-flex flex-column rounded-4 p-4">
      <div className="d-flex flex-column pb-3 mb-3 border-bottom border-secondary-subtle gap-3">
        <div className="docx-viewer-file-info">
          <span className="badge rounded-pill px-3 py-1 fw-medium mb-1">English Mass parsing</span>
          <h2 className="docx-viewer-file-name h5 text-white mb-0 font-serif" title={parseResult.fileName}>
            {parseResult.fileName}
          </h2>
          <p className="small mb-0 mt-1" style={{ color: "#94a3b8" }}>
            Detected <strong className="text-white">{parseResult.rites.length}</strong> rites and {parseResult.totalParagraphs} paragraphs.
          </p>
        </div>
        <div className="docx-viewer-actions d-flex flex-wrap align-items-center gap-2">
          {onReset && <button type="button" onClick={onReset} className="btn btn-sm btn-outline-light rounded-pill px-3 py-2">Replace file</button>}
          <button type="button" onClick={handleGenerate} disabled={isGenerating} className="btn btn-sm btn-light rounded-pill px-3 py-2">
            {isGenerating ? "Generating…" : "Download PowerPoint"}
          </button>
          <button type="button" onClick={() => setShowRaw((current) => !current)} className="btn btn-sm btn-outline-light rounded-pill px-3 py-2">
            {showRaw ? "Show parsed structure" : "Show raw paragraphs"}
          </button>
        </div>
        {generationError && <div className="alert alert-danger mb-0" role="alert">{generationError}</div>}
      </div>

      {showRaw ? (
        <ol className="mb-0 ps-4 overflow-auto">
          {parseResult.rawParagraphs.map((paragraph, index) => <li key={`${index}-${paragraph.rawText}`} className="mb-2">{paragraph.rawText}</li>)}
        </ol>
      ) : (
        <div className="d-flex flex-column gap-3 overflow-auto">
          {parseResult.rites.map((rite) => (
            <section key={rite.id} className="border border-secondary-subtle rounded-3 p-3">
              <h3 className="h6 mb-3">{rite.title}</h3>
              {rite.blocks
                .filter((block) => {
                  const normalizedTitle = block.title.trim().toUpperCase();
                  if (rite.id === "communion_rite") {
                    return normalizedTitle.startsWith("COMMUNION SONG") || normalizedTitle === "PRAYER AFTER COMMUNION";
                  }
                  if (rite.id === "concluding_rite") {
                    return normalizedTitle.startsWith("CLOSING SONG");
                  }
                  return !(
                    hiddenBlockTitles.has(normalizedTitle) ||
                    normalizedTitle.startsWith("THE PRIEST ADDRESSED THE CONGREGATION")
                  );
                })
                .filter((block) => block.type !== "creed")
                .map((block) => (
                <div key={block.id} className="border-top border-secondary-subtle pt-2 mt-2">
                  <div className="d-flex justify-content-between gap-3"><strong>{block.title}</strong><span className="text-muted small">{block.type}{block.posture ? ` · ${block.posture}` : ""}</span></div>
                  {block.title.trim().toUpperCase() !== "GOSPEL" && block.epigraph && <p className="fst-italic mb-1">{block.epigraph}</p>}
                  {block.title.trim().toUpperCase() !== "GOSPEL" && block.citation && <p className="fw-semibold mb-1">{block.citation}</p>}
                  {block.refrain && <p className="fw-semibold mb-1">Refrain: {block.refrain}</p>}
              {block.type === "prayer" ? block.paragraphs.slice(1).map((paragraph, index) => {
                    const dialogue = block.dialogues.find((line) => paragraph.rawText.trim().endsWith(line.text));
                    return dialogue ? (
                      <p key={`${block.id}-prayer-${index}`} className="mb-1"><strong>{dialogue.speaker}:</strong> {dialogue.text}</p>
                    ) : (
                      <p key={`${block.id}-prayer-${index}`} className="mb-1">{paragraph.rawText}</p>
                    );
                  }) : block.type === "reading" && block.title.trim().toUpperCase() === "GOSPEL" ? <>
                    {block.citation && <p className="fw-semibold mb-1">{block.citation}</p>}
                    {block.epigraph && <p className="fst-italic mb-1">{block.epigraph}</p>}
                    {block.paragraphs.slice(1).map((paragraph, index) => {
                      const paragraphIndex = index + 1;
                      const dialogue = block.dialogues.find((line) => line.paragraphIndex === paragraphIndex);
                      const isSource = dialogue?.text === block.citation;
                      const isEpigraph = !dialogue && paragraph.rawText.trim() === block.epigraph;
                      if (isSource || isEpigraph) return null;
                      return dialogue ? (
                        <p key={`${block.id}-gospel-${paragraphIndex}`} className="mb-1"><strong>{dialogue.speaker}:</strong> {dialogue.text}</p>
                      ) : (
                        <p key={`${block.id}-gospel-${paragraphIndex}`} className="mb-1">{paragraph.rawText}</p>
                      );
                    })}
                  </> : block.type === "reading" ? <>
                    {block.body?.map((paragraph, index) => <p key={`${block.id}-body-${index}`} className="mb-1">{paragraph}</p>)}
                    {block.dialogues.map((line) => <p key={`${line.paragraphIndex}-${line.speaker}`} className="mb-1"><strong>{line.speaker}:</strong> {line.text}</p>)}
                  </> : <>
                    {block.dialogues.map((line) => <p key={`${line.paragraphIndex}-${line.speaker}`} className="mb-1"><strong>{line.speaker}:</strong> {line.text}</p>)}
                    {block.body?.map((paragraph, index) => <p key={`${block.id}-body-${index}`} className="mb-1">{paragraph}</p>)}
                  </>}
                </div>
              ))}
            </section>
          ))}
          {parseResult.unclassifiedParagraphs.length > 0 && <p className="text-muted small mb-0">{parseResult.unclassifiedParagraphs.length} leading paragraphs were retained as unclassified.</p>}
        </div>
      )}
    </div>
  );
}
