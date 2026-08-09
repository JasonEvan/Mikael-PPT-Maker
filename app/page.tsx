"use client";

import { useState } from "react";
import WordUpload from "./components/WordUpload";
import DocxViewer from "./components/DocxViewer";
import { parseDocxFile, DocxParseResult } from "./utils/docxParser";

// ponytail: split 2-column UI when docx parsing succeeds
export default function Home() {
  const [parseResult, setParseResult] = useState<DocxParseResult | null>(null);
  const [isParsing, setIsParsing] = useState<boolean>(false);
  const [parseError, setParseError] = useState<string | null>(null);

  const handleFileConfirmed = async (file: File) => {
    setIsParsing(true);
    setParseError(null);
    try {
      const result = await parseDocxFile(file);
      setParseResult(result);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } catch (err: any) {
      console.error("Docx parse error:", err);
      setParseError(
        err?.message ||
          "Gagal membaca berkas .docx. Harap pastikan berkas tidak rusak.",
      );
    } finally {
      setIsParsing(false);
    }
  };

  const handleReset = () => {
    setParseResult(null);
    setParseError(null);
  };

  return (
    <main className="min-vh-100 py-4 px-3 d-flex flex-column align-items-center justify-content-center">
      {parseResult ? (
        /* 2-Column Split View when Parsing Succeeds */
        <div className="container-fluid max-w-7xl">
          <div className="row g-4 align-items-stretch">
            {/* Left Column: Upload Form UI */}
            <div className="col-12 col-lg-5 d-flex flex-column justify-content-center">
              <div className="bg-panel rounded-4 p-4">
                <WordUpload onFileConfirmed={handleFileConfirmed} />
              </div>
            </div>

            {/* Right Column: Scanned Docx Result Viewer */}
            <div className="col-12 col-lg-7 d-flex flex-column">
              <DocxViewer parseResult={parseResult} onReset={handleReset} />
            </div>
          </div>
        </div>
      ) : (
        /* Single Column View (Upload Only) */
        <div className="w-100 max-w-lg mx-auto">
          {isParsing ? (
            <div className="bg-panel rounded-4 p-5 text-center text-white">
              <div className="spinner-border text-danger mb-3" role="status">
                <span className="visually-hidden">Memproses...</span>
              </div>
              <h2 className="h5 font-serif mb-2">
                Membaca & Memproses Berkas Word
              </h2>
              <p className="text-sub small mb-0">
                Mengekstrak paragraf, format cetak tebal, dan bagian liturgi...
              </p>
            </div>
          ) : (
            <>
              <WordUpload onFileConfirmed={handleFileConfirmed} />
              {parseError && (
                <div className="alert alert-danger mt-3 py-2 px-3 small border-0 bg-danger bg-opacity-10 text-danger text-center rounded-3">
                  {parseError}
                </div>
              )}
            </>
          )}
        </div>
      )}
    </main>
  );
}
