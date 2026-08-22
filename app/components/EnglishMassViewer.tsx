"use client";

import { useState } from "react";
import { EnglishMassParseResult } from "../utils/englishMassParser";

interface EnglishMassViewerProps {
  parseResult: EnglishMassParseResult;
  onReset?: () => void;
}

export default function EnglishMassViewer({ parseResult, onReset }: EnglishMassViewerProps) {
  const [showRaw, setShowRaw] = useState(false);

  return (
    <div className="document-viewer w-100 h-100 d-flex flex-column rounded-4 p-4">
      <div className="d-flex flex-column pb-3 mb-3 border-bottom border-secondary-subtle gap-3">
        <div className="docx-viewer-file-info">
          <span className="badge rounded-pill px-3 py-1 fw-medium mb-1">English Mass parsing</span>
          <h2 className="docx-viewer-file-name h5 text-white mb-0 font-serif" title={parseResult.fileName}>
            {parseResult.fileName}
          </h2>
          <p className="small mb-0 mt-1" style={{ color: "#94a3b8" }}>
            Detected <strong className="text-white">{parseResult.rites.length}</strong> rites and {parseResult.totalParagraphs} paragraphs. Template mapping is not enabled yet.
          </p>
        </div>
        <div className="docx-viewer-actions d-flex flex-wrap align-items-center gap-2">
          {onReset && <button type="button" onClick={onReset} className="btn btn-sm btn-outline-light rounded-pill px-3 py-2">Replace file</button>}
          <button type="button" onClick={() => setShowRaw((current) => !current)} className="btn btn-sm btn-outline-light rounded-pill px-3 py-2">
            {showRaw ? "Show parsed structure" : "Show raw paragraphs"}
          </button>
        </div>
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
              {rite.blocks.map((block) => (
                <div key={block.id} className="border-top border-secondary-subtle pt-2 mt-2">
                  <div className="d-flex justify-content-between gap-3"><strong>{block.title}</strong><span className="text-muted small">{block.type}{block.posture ? ` · ${block.posture}` : ""}</span></div>
                  {block.epigraph && <p className="fst-italic mb-1">{block.epigraph}</p>}
                  {block.citation && <p className="fw-semibold mb-1">{block.citation}</p>}
                  {block.refrain && <p className="fw-semibold mb-1">Refrain: {block.refrain}</p>}
                  {block.type === "prayer" ? block.paragraphs.slice(1).map((paragraph, index) => {
                    const dialogue = block.dialogues.find((line) => paragraph.rawText.trim().endsWith(line.text));
                    return dialogue ? (
                      <p key={`${block.id}-prayer-${index}`} className="mb-1"><strong>{dialogue.speaker}:</strong> {dialogue.text}</p>
                    ) : (
                      <p key={`${block.id}-prayer-${index}`} className="mb-1">{paragraph.rawText}</p>
                    );
                  }) : block.type === "reading" ? <>
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
