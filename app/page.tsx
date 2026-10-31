"use client";

import { useState } from "react";
import WordUpload from "./components/WordUpload";
import DocxViewer from "./components/DocxViewer";
import { parseDocFile, parseDocxFile, DocxParseResult } from "./utils/docxParser";
import EnglishMassViewer from "./components/EnglishMassViewer";
import { EnglishMassParseResult, parseEnglishMassFile } from "./utils/englishMassParser";
import HomilyUpload from "./components/HomilyUpload";

type MassType = "indonesian" | "english";
type ParseResult = DocxParseResult | EnglishMassParseResult;

export default function Home() {
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [isParsing, setIsParsing] = useState<boolean>(false);
  const [parseError, setParseError] = useState<string | null>(null);
  const [massType, setMassType] = useState<MassType>("indonesian");
  const [homily, setHomily] = useState<EnglishMassParseResult["homily"]>();
  const [isParsingHomily, setIsParsingHomily] = useState(false);

  const handleFileConfirmed = async (file: File) => {
    setIsParsing(true);
    setParseError(null);
    try {
      const result = massType === "english"
        ? await parseEnglishMassFile(file)
        : file.name.toLowerCase().endsWith(".doc")
          ? await parseDocFile(file)
          : await parseDocxFile(file);
      setParseResult(result);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } catch (err: any) {
      console.error("Docx parse error:", err);
      setParseError(
        err?.message ||
          "Gagal membaca berkas Word. Harap pastikan berkas tidak rusak.",
      );
    } finally {
      setIsParsing(false);
    }
  };

  const handleReset = () => {
    setParseResult(null);
    setParseError(null);
    setHomily(undefined);
  };

  return (
    <main className="app-shell">
      <div className="page-frame">
        <header className="site-header">
          <a className="brand" href="#" aria-label="Liturgi Slides">
            <span className="brand-mark" aria-hidden="true">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none">
                <path d="M7 4.5h8.5A2.5 2.5 0 0 1 18 7v12.5H9.5A2.5 2.5 0 0 1 7 17V4.5Z" />
                <path d="M7 7H5.5A1.5 1.5 0 0 0 4 8.5v9A2.5 2.5 0 0 0 6.5 20H18" />
                <path d="M12.5 8v6M9.5 11h6" />
              </svg>
            </span>
            <span>
              <strong>Liturgi</strong> Slides
            </span>
          </a>
          <span className="header-note">Alat internal paroki</span>
        </header>

        {parseResult ? (
          <section className="workspace-layout" aria-label="Hasil dokumen">
            <aside className="compact-panel">
              <div className="compact-panel-heading">
                <span className="eyebrow">Dokumen sumber</span>
                <h1>Ganti naskah</h1>
                <p>Unggah dokumen lain untuk memulai pemetaan baru.</p>
              </div>
              <fieldset disabled={isParsingHomily || isParsing} className="border-0 p-0 m-0">
                <WordUpload compact onFileConfirmed={handleFileConfirmed} />
              </fieldset>
              {"massType" in parseResult && (
                <HomilyUpload homily={homily} onChange={setHomily} onBusyChange={setIsParsingHomily} />
              )}
              {parseError && <div className="form-alert" role="alert">{parseError}</div>}
            </aside>
            {"massType" in parseResult ? (
              <EnglishMassViewer
                parseResult={{ ...parseResult, homily }}
                onReset={handleReset}
                isUploadingHomily={isParsingHomily}
              />
            ) : (
              <DocxViewer parseResult={parseResult} onReset={handleReset} />
            )}
          </section>
        ) : (
          <section className="landing-layout">
            <div className="hero-copy">
              <span className="eyebrow">Penyusun slide liturgi</span>
              <h1>
                Dari naskah liturgi ke slide yang <em>siap tayang.</em>
              </h1>
              <p className="hero-lead">
                Ubah dokumen Word mingguan menjadi PowerPoint lengkap dalam
                hitungan detik—tanpa menyalin teks satu per satu.
              </p>

              <ol className="workflow-list" aria-label="Alur kerja">
                <li>
                  <span>01</span>
                  <div>
                    <strong>Unggah naskah</strong>
                    <p>Pilih dokumen liturgi berformat .doc atau .docx.</p>
                  </div>
                </li>
                <li>
                  <span>02</span>
                  <div>
                    <strong>Periksa pemetaan</strong>
                    <p>Pastikan setiap bagian terbaca dengan tepat.</p>
                  </div>
                </li>
                <li>
                  <span>03</span>
                  <div>
                    <strong>Unduh presentasi</strong>
                    <p>Pilih template dan dapatkan berkas siap pakai.</p>
                  </div>
                </li>
              </ol>
            </div>

            <div className="upload-panel">
            <div className="upload-panel-heading">
              <span className="step-pill">Langkah 1 dari 3</span>
              <h2>Unggah naskah minggu ini</h2>
              <p>Kami akan membaca struktur dan memetakan isinya otomatis.</p>
            </div>

            <div className="mass-type-selector" role="group" aria-label="Jenis misa">
              <button
                type="button"
                className={massType === "indonesian" ? "is-selected" : ""}
                aria-pressed={massType === "indonesian"}
                disabled={isParsingHomily}
                onClick={() => { setMassType("indonesian"); setParseError(null); }}
              >
                Misa Bahasa Indonesia
              </button>
              <button
                type="button"
                className={massType === "english" ? "is-selected" : ""}
                aria-pressed={massType === "english"}
                disabled={isParsingHomily}
                onClick={() => { setMassType("english"); setParseError(null); }}
              >
                English Mass
              </button>
            </div>

            {isParsing ? (
                <div className="processing-state" aria-live="polite">
                  <div className="processing-spinner" role="status">
                    <span className="visually-hidden">Memproses...</span>
                  </div>
                  <h2>Membaca naskah liturgi</h2>
                  <p>
                    Mengekstrak paragraf, format cetak tebal, dan bagian
                    liturgi…
                  </p>
                </div>
              ) : (
                <>
                  <fieldset disabled={isParsingHomily} className="border-0 p-0 m-0">
                    <WordUpload onFileConfirmed={handleFileConfirmed} />
                  </fieldset>
                  {parseError && (
                    <div className="form-alert" role="alert">
                      {parseError}
                    </div>
                  )}
                </>
              )}
            </div>
          </section>
        )}

        <footer className="site-footer">
          <span>Liturgi Slides</span>
          <span>Dokumen diproses langsung untuk kebutuhan presentasi.</span>
        </footer>
      </div>
    </main>
  );
}
