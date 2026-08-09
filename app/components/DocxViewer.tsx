"use client";

import { useState } from "react";
import {
  DocxParseResult,
  selectRelevantSections,
  MappedLiturgyData,
} from "../utils/docxParser";

interface DocxViewerProps {
  parseResult: DocxParseResult;
  onReset?: () => void;
}

// ponytail: modern segmented tab bar with search support across all views
export default function DocxViewer({ parseResult, onReset }: DocxViewerProps) {
  const [activeTab, setActiveTab] = useState<"mapped" | "sections" | "raw">("mapped");
  const [searchTerm, setSearchTerm] = useState<string>("");

  const mappedData: MappedLiturgyData = selectRelevantSections(parseResult);

  // Filter sections for "Semua Bagian"
  const filteredSections = parseResult.sections.filter(
    (sec) =>
      sec.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
      sec.body?.some((b) => b.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  // Filter raw paragraphs for "Teks Mentah"
  const filteredRawParagraphs = parseResult.rawParagraphs.filter((p) =>
    p.rawText.toLowerCase().includes(searchTerm.toLowerCase())
  );

  // Helper to check if mapped item matches search term
  const matchesSearch = (title: string, content?: string | string[]) => {
    if (!searchTerm) return true;
    const term = searchTerm.toLowerCase();
    if (title.toLowerCase().includes(term)) return true;
    if (typeof content === "string") return content.toLowerCase().includes(term);
    if (Array.isArray(content)) return content.some((c) => c.toLowerCase().includes(term));
    return false;
  };

  return (
    <div className="w-100 h-100 d-flex flex-column bg-panel rounded-4 p-4 text-white shadow-lg border border-secondary-subtle">
      {/* Header Info */}
      <div className="d-flex align-items-center justify-content-between pb-3 mb-3 border-bottom border-secondary-subtle">
        <div>
          <span className="badge rounded-pill px-3 py-1 fw-medium mb-1" style={{ backgroundColor: "rgba(74, 222, 128, 0.15)", color: "#4ade80", border: "1px solid rgba(74, 222, 128, 0.3)" }}>
            Parsing & Pemetaan Berhasil
          </span>
          <h2 className="h5 text-white mb-0 font-serif text-truncate max-w-sm" title={parseResult.fileName}>
            {parseResult.fileName}
          </h2>
          <p className="small mb-0 mt-1" style={{ color: "#94a3b8" }}>
            Terdeteksi <strong className="text-white">{parseResult.sections.length}</strong> Bagian ({parseResult.totalParagraphs} paragraf)
          </p>
        </div>
        {onReset && (
          <button
            type="button"
            onClick={onReset}
            className="btn btn-sm btn-outline-light rounded-pill px-3 py-1.5"
          >
            Tutup
          </button>
        )}
      </div>

      {/* Modern Segmented Tab Bar & Integrated Search */}
      <div className="d-flex flex-column flex-md-row align-items-md-center justify-content-between gap-3 mb-4">
        {/* Segmented Control */}
        <div
          className="d-flex p-1 rounded-3 gap-1 border border-secondary-subtle"
          style={{ backgroundColor: "rgba(0, 0, 0, 0.35)" }}
        >
          <button
            type="button"
            className={`btn btn-sm rounded-2 px-3 py-1.5 fw-medium transition-all ${
              activeTab === "mapped"
                ? "btn-wine text-white shadow-sm"
                : "text-sub bg-transparent border-0"
            }`}
            style={{
              color: activeTab === "mapped" ? "#ffffff" : "#94a3b8",
            }}
            onClick={() => setActiveTab("mapped")}
          >
            Hasil Pemetaan (MVP)
          </button>
          <button
            type="button"
            className={`btn btn-sm rounded-2 px-3 py-1.5 fw-medium transition-all ${
              activeTab === "sections"
                ? "btn-wine text-white shadow-sm"
                : "text-sub bg-transparent border-0"
            }`}
            style={{
              color: activeTab === "sections" ? "#ffffff" : "#94a3b8",
            }}
            onClick={() => setActiveTab("sections")}
          >
            Semua Bagian ({parseResult.sections.length})
          </button>
          <button
            type="button"
            className={`btn btn-sm rounded-2 px-3 py-1.5 fw-medium transition-all ${
              activeTab === "raw"
                ? "btn-wine text-white shadow-sm"
                : "text-sub bg-transparent border-0"
            }`}
            style={{
              color: activeTab === "raw" ? "#ffffff" : "#94a3b8",
            }}
            onClick={() => setActiveTab("raw")}
          >
            Teks Mentah ({parseResult.totalParagraphs})
          </button>
        </div>

        {/* Universal Search Input */}
        <div className="position-relative min-w-xs">
          <svg
            className="position-absolute top-50 start-0 translate-middle-y ms-3"
            width="16"
            height="16"
            fill="none"
            viewBox="0 0 24 24"
            stroke="#94a3b8"
            strokeWidth="2"
          >
            <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
          </svg>
          <input
            type="text"
            placeholder="Cari dalam dokumen..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="form-control form-control-sm text-white border-secondary-subtle rounded-3 ps-5 pe-3 py-1.5"
            style={{
              color: "#f8fafc",
              backgroundColor: "rgba(0, 0, 0, 0.35)",
              fontSize: "0.875rem",
            }}
          />
        </div>
      </div>

      {/* Content Scroll View */}
      <div className="flex-grow-1 overflow-auto pe-2" style={{ maxHeight: "68vh" }}>
        {activeTab === "mapped" ? (
          /* Mapped Liturgy Data View (MVP Target Output) */
          <div className="d-flex flex-column gap-3">
            {/* Unmatched Sections Alert Box */}
            {mappedData.unmatchedSections.length > 0 && (
              <div
                className="alert border-0 p-3 rounded-3 mb-1"
                style={{ backgroundColor: "rgba(234, 179, 8, 0.15)", borderLeft: "4px solid #facc15" }}
              >
                <strong className="d-block mb-1" style={{ color: "#fef08a" }}>
                  ⚠️ Bagian Tidak Terpetakan ({mappedData.unmatchedSections.length}):
                </strong>
                <p className="small mb-0" style={{ color: "#fef9c3" }}>
                  {mappedData.unmatchedSections.join(" • ")}
                </p>
                <span className="small mt-1 d-block" style={{ fontSize: "0.78rem", color: "#fde047" }}>
                  *Bagian di atas tidak cocok dengan 9 target slide MVP. Periksa jika ada kesalahan penamaan header pada berkas Word.
                </span>
              </div>
            )}

            {/* Mapped Item Cards */}
            <div className="d-flex flex-column gap-3">
              {/* 1. Seruan Tobat */}
              {matchesSearch("Seruan Tobat", mappedData.seruanTobat) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">1. Seruan Tobat</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.seruanTobat ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.seruanTobat ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.seruanTobat ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.seruanTobat ? "Terisi" : "Kosong"}
                    </span>
                  </div>
                  {mappedData.seruanTobat ? (
                    <ul className="small mb-0 ps-3" style={{ color: "#e2e8f0" }}>
                      {mappedData.seruanTobat.map((line, i) => (
                        <li key={i}>{line}</li>
                      ))}
                    </ul>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan di berkas (Placeholder akan dikosongkan)
                    </p>
                  )}
                </div>
              )}

              {/* 2. Doa Kolekta */}
              {matchesSearch("Doa Kolekta", mappedData.doaKolekta) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">2. Doa Kolekta</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.doaKolekta ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.doaKolekta ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.doaKolekta ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.doaKolekta ? "Terisi" : "Kosong"}
                    </span>
                  </div>
                  {mappedData.doaKolekta ? (
                    <p className="small mb-0 whitespace-pre-wrap" style={{ color: "#e2e8f0" }}>
                      {mappedData.doaKolekta}
                    </p>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}

              {/* 3. Bacaan Pertama */}
              {matchesSearch("Bacaan Pertama", [
                mappedData.bacaanPertama?.epigraph || "",
                mappedData.bacaanPertama?.citation || "",
                mappedData.bacaanPertama?.notesBody || "",
              ]) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">3. Bacaan Pertama</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.bacaanPertama ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.bacaanPertama ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.bacaanPertama ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.bacaanPertama ? "Terisi" : "Kosong"}
                    </span>
                  </div>
                  {mappedData.bacaanPertama ? (
                    <div className="small d-flex flex-column gap-2">
                      {mappedData.bacaanPertama.epigraph && (
                        <div className="p-2.5 rounded-end fst-italic" style={{ backgroundColor: "rgba(139, 38, 62, 0.25)", borderLeft: "3px solid #f43f5e", color: "#fecdd3" }}>
                          "{mappedData.bacaanPertama.epigraph}"
                        </div>
                      )}
                      {mappedData.bacaanPertama.citation && (
                        <div className="fw-semibold" style={{ color: "#fbbf24" }}>
                          📍 {mappedData.bacaanPertama.citation}
                        </div>
                      )}
                      {mappedData.bacaanPertama.notesBody && (
                        <div style={{ color: "#cbd5e1" }}>
                          <strong className="text-white">Notes Pane Text:</strong> {mappedData.bacaanPertama.notesBody.slice(0, 180)}...
                        </div>
                      )}
                    </div>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}

              {/* 4. Mazmur Tanggapan */}
              {matchesSearch("Mazmur Tanggapan", mappedData.mazmurTanggapan?.verses) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">4. Mazmur Tanggapan</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.mazmurTanggapan ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.mazmurTanggapan ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.mazmurTanggapan ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.mazmurTanggapan ? `${mappedData.mazmurTanggapan.verses.length} Bait` : "Kosong"}
                    </span>
                  </div>
                  {mappedData.mazmurTanggapan ? (
                    <div className="d-flex flex-column gap-2">
                      {mappedData.mazmurTanggapan.verses.map((v, i) => (
                        <div key={i} className="bg-black bg-opacity-20 p-2.5 rounded small border border-secondary-subtle" style={{ color: "#e2e8f0" }}>
                          <strong className="text-white d-block mb-1">Bait {i + 1}:</strong>
                          <span className="whitespace-pre-wrap">{v}</span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}

              {/* 5. Bacaan Kedua */}
              {matchesSearch("Bacaan Kedua", [
                mappedData.bacaanKedua?.epigraph || "",
                mappedData.bacaanKedua?.citation || "",
              ]) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">5. Bacaan Kedua</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.bacaanKedua ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.bacaanKedua ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.bacaanKedua ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.bacaanKedua ? "Terisi" : "Kosong"}
                    </span>
                  </div>
                  {mappedData.bacaanKedua ? (
                    <div className="small d-flex flex-column gap-2">
                      {mappedData.bacaanKedua.epigraph && (
                        <div className="p-2.5 rounded-end fst-italic" style={{ backgroundColor: "rgba(139, 38, 62, 0.25)", borderLeft: "3px solid #f43f5e", color: "#fecdd3" }}>
                          "{mappedData.bacaanKedua.epigraph}"
                        </div>
                      )}
                      {mappedData.bacaanKedua.citation && (
                        <div className="fw-semibold" style={{ color: "#fbbf24" }}>
                          📍 {mappedData.bacaanKedua.citation}
                        </div>
                      )}
                    </div>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}

              {/* 6. Bacaan Injil */}
              {matchesSearch("Bacaan Injil", [
                mappedData.bacaanInjil?.gospelName || "",
                mappedData.bacaanInjil?.citation || "",
              ]) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">6. Bacaan Injil</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.bacaanInjil ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.bacaanInjil ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.bacaanInjil ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.bacaanInjil ? "Terisi" : "Kosong"}
                    </span>
                  </div>
                  {mappedData.bacaanInjil ? (
                    <div className="small d-flex flex-column gap-2">
                      {mappedData.bacaanInjil.gospelName && (
                        <div className="badge rounded-pill px-2.5 py-1 w-fit" style={{ backgroundColor: "rgba(239, 68, 68, 0.2)", color: "#fca5a5", border: "1px solid rgba(239, 68, 68, 0.3)" }}>
                          📖 {mappedData.bacaanInjil.gospelName}
                        </div>
                      )}
                      {mappedData.bacaanInjil.citation && (
                        <div className="fw-semibold" style={{ color: "#fbbf24" }}>
                          📍 {mappedData.bacaanInjil.citation}
                        </div>
                      )}
                    </div>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}

              {/* 7. Doa Umat */}
              {matchesSearch("Doa Umat", [
                mappedData.doaUmat?.intro || "",
                ...(mappedData.doaUmat?.intentions.map((i) => i.title + " " + i.body) || []),
                mappedData.doaUmat?.closing || "",
              ]) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">7. Doa Umat</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.doaUmat ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.doaUmat ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.doaUmat ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.doaUmat ? `${mappedData.doaUmat.intentions.length} Niatan` : "Kosong"}
                    </span>
                  </div>
                  {mappedData.doaUmat ? (
                    <div className="small d-flex flex-column gap-2">
                      {mappedData.doaUmat.intro && (
                        <p className="mb-1" style={{ color: "#cbd5e1" }}>
                          <strong className="text-white">Intro:</strong> {mappedData.doaUmat.intro}
                        </p>
                      )}
                      {mappedData.doaUmat.intentions.map((item, i) => (
                        <div key={i} className="bg-black bg-opacity-20 p-2.5 rounded border border-secondary-subtle" style={{ color: "#e2e8f0" }}>
                          <strong className="text-white d-block mb-1">{item.title}</strong>
                          <span>{item.body}</span>
                        </div>
                      ))}
                      {mappedData.doaUmat.closing && (
                        <p className="mb-0 mt-1" style={{ color: "#cbd5e1" }}>
                          <strong className="text-white">Penutup:</strong> {mappedData.doaUmat.closing}
                        </p>
                      )}
                    </div>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}

              {/* 8. Doa Atas Persembahan */}
              {matchesSearch("Doa Atas Persembahan", mappedData.doaAtasPersembahan) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">8. Doa Atas Persembahan</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.doaAtasPersembahan ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.doaAtasPersembahan ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.doaAtasPersembahan ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.doaAtasPersembahan ? "Terisi" : "Kosong"}
                    </span>
                  </div>
                  {mappedData.doaAtasPersembahan ? (
                    <p className="small mb-0" style={{ color: "#e2e8f0" }}>{mappedData.doaAtasPersembahan}</p>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}

              {/* 9. Doa Sesudah Komuni */}
              {matchesSearch("Doa Sesudah Komuni", mappedData.doaSesudahKomuni) && (
                <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-center justify-content-between mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">9. Doa Sesudah Komuni</h3>
                    <span
                      className="badge rounded-pill px-2.5 py-1"
                      style={{
                        backgroundColor: mappedData.doaSesudahKomuni ? "rgba(74, 222, 128, 0.2)" : "rgba(148, 163, 184, 0.2)",
                        color: mappedData.doaSesudahKomuni ? "#4ade80" : "#94a3b8",
                        border: `1px solid ${mappedData.doaSesudahKomuni ? "rgba(74, 222, 128, 0.3)" : "rgba(148, 163, 184, 0.3)"}`,
                      }}
                    >
                      {mappedData.doaSesudahKomuni ? "Terisi" : "Kosong"}
                    </span>
                  </div>
                  {mappedData.doaSesudahKomuni ? (
                    <p className="small mb-0" style={{ color: "#e2e8f0" }}>{mappedData.doaSesudahKomuni}</p>
                  ) : (
                    <p className="small mb-0 fst-italic" style={{ color: "#94a3b8" }}>
                      Tidak ditemukan
                    </p>
                  )}
                </div>
              )}
            </div>
          </div>
        ) : activeTab === "sections" ? (
          /* All Parsed Sections View */
          <div className="d-flex flex-column gap-3">
            {filteredSections.length === 0 ? (
              <p className="small text-center py-4" style={{ color: "#94a3b8" }}>
                Tidak ada bagian yang cocok dengan pencarian.
              </p>
            ) : (
              filteredSections.map((section, idx) => (
                <div key={section.id} className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                  <div className="d-flex align-items-start justify-content-between gap-2 mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">
                      <span className="me-2" style={{ color: "#94a3b8" }}>{idx + 1}.</span>
                      {section.title}
                    </h3>
                    <span
                      className="badge rounded-pill text-uppercase px-2.5 py-1"
                      style={{
                        backgroundColor:
                          section.type === "major_header"
                            ? "rgba(239, 68, 68, 0.2)"
                            : section.type === "reading"
                            ? "rgba(245, 158, 11, 0.2)"
                            : section.type === "doa_umat"
                            ? "rgba(14, 165, 233, 0.2)"
                            : "rgba(148, 163, 184, 0.2)",
                        color:
                          section.type === "major_header"
                            ? "#fca5a5"
                            : section.type === "reading"
                            ? "#fde047"
                            : section.type === "doa_umat"
                            ? "#38bdf8"
                            : "#cbd5e1",
                        border: `1px solid ${
                          section.type === "major_header"
                            ? "rgba(239, 68, 68, 0.3)"
                            : section.type === "reading"
                            ? "rgba(245, 158, 11, 0.3)"
                            : section.type === "doa_umat"
                            ? "rgba(14, 165, 233, 0.3)"
                            : "rgba(148, 163, 184, 0.3)"
                        }`,
                      }}
                    >
                      {section.type.replace("_", " ")}
                    </span>
                  </div>

                  {section.epigraph && (
                    <div className="p-2.5 mb-2 rounded-end small fst-italic" style={{ backgroundColor: "rgba(139, 38, 62, 0.25)", borderLeft: "3px solid #f43f5e", color: "#fecdd3" }}>
                      "{section.epigraph}"
                    </div>
                  )}

                  {section.citation && (
                    <div className="small fw-semibold mb-2" style={{ color: "#fbbf24" }}>
                      📍 {section.citation}
                    </div>
                  )}

                  {section.intentions && section.intentions.length > 0 && (
                    <div className="d-flex flex-column gap-2 my-2">
                      <p className="small mb-1 font-monospace" style={{ color: "#cbd5e1" }}>
                        Niatan Doa Umat ({section.intentions.length} niatan):
                      </p>
                      {section.intentions.map((item, iIndex) => (
                        <div key={iIndex} className="bg-black bg-opacity-20 p-2.5 rounded small border border-secondary-subtle" style={{ color: "#e2e8f0" }}>
                          <strong className="text-white d-block mb-1">{item.title}</strong>
                          <span>{item.body || "(Tanpa badan doa)"}</span>
                        </div>
                      ))}
                    </div>
                  )}

                  {section.body && section.body.length > 0 && !section.intentions && (
                    <div className="small mt-2" style={{ color: "#cbd5e1" }}>
                      {section.body.map((line, lIdx) => (
                        <p key={lIdx} className="mb-1">{line}</p>
                      ))}
                    </div>
                  )}
                </div>
              ))
            )}
          </div>
        ) : (
          /* Raw Paragraphs View */
          <div className="d-flex flex-column gap-2">
            {filteredRawParagraphs.length === 0 ? (
              <p className="small text-center py-4" style={{ color: "#94a3b8" }}>
                Tidak ada paragraf yang cocok dengan pencarian.
              </p>
            ) : (
              filteredRawParagraphs.map((p, idx) => (
                <div key={idx} className="bg-black bg-opacity-20 p-2.5 rounded border border-secondary-subtle d-flex flex-column gap-1">
                  <div className="d-flex align-items-center justify-content-between small">
                    <span className="font-monospace" style={{ color: "#94a3b8" }}>#{idx + 1}</span>
                    <div className="d-flex gap-1">
                      {p.isBold && <span className="badge px-2 py-0.5" style={{ backgroundColor: "rgba(239, 68, 68, 0.25)", color: "#fca5a5" }}>Bold</span>}
                      {p.isUnderline && <span className="badge px-2 py-0.5" style={{ backgroundColor: "rgba(59, 130, 246, 0.25)", color: "#93c5fd" }}>Underline</span>}
                      {p.isItalic && <span className="badge px-2 py-0.5" style={{ backgroundColor: "rgba(245, 158, 11, 0.25)", color: "#fde047" }}>Italic</span>}
                    </div>
                  </div>
                  <p className="mb-0 small text-white">{p.rawText}</p>
                </div>
              ))
            )}
          </div>
        )}
      </div>
    </div>
  );
}
