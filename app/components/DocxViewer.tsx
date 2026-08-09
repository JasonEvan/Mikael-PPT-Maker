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

// ponytail: UI viewer with 3 tabs: Mapped MVP Data, Parsed Sections, Raw Paragraphs
export default function DocxViewer({ parseResult, onReset }: DocxViewerProps) {
  const [activeTab, setActiveTab] = useState<"mapped" | "sections" | "raw">(
    "mapped",
  );
  const [searchTerm, setSearchTerm] = useState<string>("");

  const mappedData: MappedLiturgyData = selectRelevantSections(parseResult);

  const filteredSections = parseResult.sections.filter(
    (sec) =>
      sec.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
      sec.body?.some((b) => b.toLowerCase().includes(searchTerm.toLowerCase())),
  );

  return (
    <div className="w-100 h-100 d-flex flex-column bg-panel rounded-4 p-4 text-white">
      {/* Header Info */}
      <div className="d-flex align-items-center justify-content-between pb-3 mb-3 border-bottom border-secondary-subtle">
        <div>
          <span className="badge bg-success bg-opacity-20 text-success border border-success-subtle mb-1">
            Parsing & Pemetaan Berhasil
          </span>
          <h2
            className="h5 text-white mb-0 font-serif text-truncate max-w-sm"
            title={parseResult.fileName}
          >
            {parseResult.fileName}
          </h2>
          <p className="text-sub small mb-0 mt-1">
            Terdeteksi{" "}
            <strong className="text-white">
              {parseResult.sections.length}
            </strong>{" "}
            Bagian ({parseResult.totalParagraphs} paragraf)
          </p>
        </div>
        {onReset && (
          <button
            type="button"
            onClick={onReset}
            className="btn btn-sm btn-outline-light rounded-pill px-3"
          >
            Tutup
          </button>
        )}
      </div>

      {/* Navigation Tabs & Search */}
      <div className="d-flex flex-column flex-sm-row align-items-sm-center justify-content-between gap-2 mb-3">
        <div
          className="btn-group btn-group-sm bg-black bg-opacity-30 p-1 rounded-pill"
          role="group"
        >
          <button
            type="button"
            className={`btn rounded-pill px-3 ${activeTab === "mapped" ? "btn-wine text-white" : "text-sub"}`}
            onClick={() => setActiveTab("mapped")}
          >
            Hasil Pemetaan (MVP)
          </button>
          <button
            type="button"
            className={`btn rounded-pill px-3 ${activeTab === "sections" ? "btn-wine text-white" : "text-sub"}`}
            onClick={() => setActiveTab("sections")}
          >
            Semua Bagian ({parseResult.sections.length})
          </button>
          <button
            type="button"
            className={`btn rounded-pill px-3 ${activeTab === "raw" ? "btn-wine text-white" : "text-sub"}`}
            onClick={() => setActiveTab("raw")}
          >
            Teks Mentah ({parseResult.totalParagraphs})
          </button>
        </div>

        {activeTab !== "mapped" && (
          <input
            type="text"
            placeholder="Cari teks..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="form-control form-control-sm bg-black bg-opacity-30 text-white border-secondary-subtle rounded-pill px-3 max-w-xs"
          />
        )}
      </div>

      {/* Content Scroll View */}
      <div className="grow overflow-auto pe-2" style={{ maxHeight: "70vh" }}>
        {activeTab === "mapped" ? (
          /* Mapped Liturgy Data View (MVP Target Output) */
          <div className="d-flex flex-column gap-3">
            {/* Unmatched Sections Alert Box */}
            {mappedData.unmatchedSections.length > 0 && (
              <div className="alert alert-warning border-0 bg-warning bg-opacity-10 text-warning p-3 rounded-3 mb-1">
                <strong className="d-block mb-1">
                  ⚠️ Bagian Tidak Terpetakan (
                  {mappedData.unmatchedSections.length}):
                </strong>
                <p className="small mb-0">
                  {mappedData.unmatchedSections.join(" • ")}
                </p>
                <span
                  className="small text-muted-custom mt-1 d-block"
                  style={{ fontSize: "0.78rem" }}
                >
                  *Bagian di atas tidak cocok dengan 9 target slide MVP. Periksa
                  jika ada kesalahan penamaan header pada berkas Word.
                </span>
              </div>
            )}

            {/* Mapped Item Cards */}
            <div className="d-flex flex-column gap-3">
              {/* 1. Seruan Tobat */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    1. Seruan Tobat
                  </h3>
                  <span
                    className={`badge ${mappedData.seruanTobat ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.seruanTobat ? "Terisi" : "Kosong"}
                  </span>
                </div>
                {mappedData.seruanTobat ? (
                  <ul className="small text-sub mb-0 ps-3">
                    {mappedData.seruanTobat.map((line, i) => (
                      <li key={i}>{line}</li>
                    ))}
                  </ul>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan di berkas (Placeholder akan dikosongkan)
                  </p>
                )}
              </div>

              {/* 2. Doa Kolekta */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    2. Doa Kolekta
                  </h3>
                  <span
                    className={`badge ${mappedData.doaKolekta ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.doaKolekta ? "Terisi" : "Kosong"}
                  </span>
                </div>
                {mappedData.doaKolekta ? (
                  <p className="small text-sub mb-0 whitespace-pre-wrap">
                    {mappedData.doaKolekta}
                  </p>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>

              {/* 3. Bacaan Pertama */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    3. Bacaan Pertama
                  </h3>
                  <span
                    className={`badge ${mappedData.bacaanPertama ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.bacaanPertama ? "Terisi" : "Kosong"}
                  </span>
                </div>
                {mappedData.bacaanPertama ? (
                  <div className="small d-flex flex-column gap-2">
                    {mappedData.bacaanPertama.epigraph && (
                      <div className="bg-wine bg-opacity-10 border-start border-3 border-danger p-2 rounded-end text-sub fst-italic">
                        &quot;{mappedData.bacaanPertama.epigraph}&quot;
                      </div>
                    )}
                    {mappedData.bacaanPertama.citation && (
                      <div className="fw-semibold text-warning">
                        📍 {mappedData.bacaanPertama.citation}
                      </div>
                    )}
                    {mappedData.bacaanPertama.notesBody && (
                      <div className="text-muted-custom">
                        <strong>Notes Pane Text:</strong>{" "}
                        {mappedData.bacaanPertama.notesBody.slice(0, 150)}...
                      </div>
                    )}
                  </div>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>

              {/* 4. Mazmur Tanggapan */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    4. Mazmur Tanggapan
                  </h3>
                  <span
                    className={`badge ${mappedData.mazmurTanggapan ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.mazmurTanggapan
                      ? `${mappedData.mazmurTanggapan.verses.length} Bait`
                      : "Kosong"}
                  </span>
                </div>
                {mappedData.mazmurTanggapan ? (
                  <div className="d-flex flex-column gap-2">
                    {mappedData.mazmurTanggapan.verses.map((v, i) => (
                      <div
                        key={i}
                        className="bg-black bg-opacity-20 p-2 rounded small text-sub"
                      >
                        <strong className="text-white d-block mb-1">
                          Bait {i + 1}:
                        </strong>
                        <span className="whitespace-pre-wrap">{v}</span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>

              {/* 5. Bacaan Kedua */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    5. Bacaan Kedua
                  </h3>
                  <span
                    className={`badge ${mappedData.bacaanKedua ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.bacaanKedua ? "Terisi" : "Kosong"}
                  </span>
                </div>
                {mappedData.bacaanKedua ? (
                  <div className="small d-flex flex-column gap-2">
                    {mappedData.bacaanKedua.epigraph && (
                      <div className="bg-wine bg-opacity-10 border-start border-3 border-danger p-2 rounded-end text-sub fst-italic">
                        &quot;{mappedData.bacaanKedua.epigraph}&quot;
                      </div>
                    )}
                    {mappedData.bacaanKedua.citation && (
                      <div className="fw-semibold text-warning">
                        📍 {mappedData.bacaanKedua.citation}
                      </div>
                    )}
                  </div>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>

              {/* 6. Bacaan Injil */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    6. Bacaan Injil
                  </h3>
                  <span
                    className={`badge ${mappedData.bacaanInjil ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.bacaanInjil ? "Terisi" : "Kosong"}
                  </span>
                </div>
                {mappedData.bacaanInjil ? (
                  <div className="small d-flex flex-column gap-2">
                    {mappedData.bacaanInjil.gospelName && (
                      <div className="badge bg-danger bg-opacity-20 text-danger border border-danger-subtle w-fit">
                        📖 {mappedData.bacaanInjil.gospelName}
                      </div>
                    )}
                    {mappedData.bacaanInjil.citation && (
                      <div className="fw-semibold text-warning">
                        📍 {mappedData.bacaanInjil.citation}
                      </div>
                    )}
                  </div>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>

              {/* 7. Doa Umat */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    7. Doa Umat
                  </h3>
                  <span
                    className={`badge ${mappedData.doaUmat ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.doaUmat
                      ? `${mappedData.doaUmat.intentions.length} Niatan`
                      : "Kosong"}
                  </span>
                </div>
                {mappedData.doaUmat ? (
                  <div className="small d-flex flex-column gap-2">
                    {mappedData.doaUmat.intro && (
                      <p className="text-sub mb-1">
                        <strong>Intro:</strong> {mappedData.doaUmat.intro}
                      </p>
                    )}
                    {mappedData.doaUmat.intentions.map((item, i) => (
                      <div
                        key={i}
                        className="bg-black bg-opacity-20 p-2 rounded border border-secondary-subtle"
                      >
                        <strong className="text-white d-block mb-1">
                          {item.title}
                        </strong>
                        <span className="text-sub">{item.body}</span>
                      </div>
                    ))}
                    {mappedData.doaUmat.closing && (
                      <p className="text-sub mb-0 mt-1">
                        <strong>Penutup:</strong> {mappedData.doaUmat.closing}
                      </p>
                    )}
                  </div>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>

              {/* 8. Doa Atas Persembahan */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    8. Doa Atas Persembahan
                  </h3>
                  <span
                    className={`badge ${mappedData.doaAtasPersembahan ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.doaAtasPersembahan ? "Terisi" : "Kosong"}
                  </span>
                </div>
                {mappedData.doaAtasPersembahan ? (
                  <p className="small text-sub mb-0">
                    {mappedData.doaAtasPersembahan}
                  </p>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>

              {/* 9. Doa Sesudah Komuni */}
              <div className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle">
                <div className="d-flex align-items-center justify-content-between mb-2">
                  <h3 className="h6 text-white mb-0 fw-semibold">
                    9. Doa Sesudah Komuni
                  </h3>
                  <span
                    className={`badge ${mappedData.doaSesudahKomuni ? "bg-success" : "bg-secondary"} bg-opacity-30`}
                  >
                    {mappedData.doaSesudahKomuni ? "Terisi" : "Kosong"}
                  </span>
                </div>
                {mappedData.doaSesudahKomuni ? (
                  <p className="small text-sub mb-0">
                    {mappedData.doaSesudahKomuni}
                  </p>
                ) : (
                  <p className="small text-muted-custom mb-0 fst-italic">
                    Tidak ditemukan
                  </p>
                )}
              </div>
            </div>
          </div>
        ) : activeTab === "sections" ? (
          /* All Parsed Sections View */
          <div className="d-flex flex-column gap-3">
            {filteredSections.length === 0 ? (
              <p className="text-sub small text-center py-4">
                Tidak ada bagian yang cocok dengan pencarian.
              </p>
            ) : (
              filteredSections.map((section, idx) => (
                <div
                  key={section.id}
                  className="bg-black bg-opacity-30 rounded-3 p-3 border border-secondary-subtle"
                >
                  <div className="d-flex align-items-start justify-content-between gap-2 mb-2">
                    <h3 className="h6 text-white mb-0 fw-semibold">
                      <span className="text-muted-custom me-2">{idx + 1}.</span>
                      {section.title}
                    </h3>
                    <span
                      className={`badge rounded-pill text-uppercase ${
                        section.type === "major_header"
                          ? "bg-danger bg-opacity-20 text-danger border border-danger-subtle"
                          : section.type === "reading"
                            ? "bg-warning bg-opacity-20 text-warning border border-warning-subtle"
                            : section.type === "doa_umat"
                              ? "bg-info bg-opacity-20 text-info border border-info-subtle"
                              : "bg-secondary bg-opacity-20 text-sub"
                      }`}
                    >
                      {section.type.replace("_", " ")}
                    </span>
                  </div>

                  {section.epigraph && (
                    <div className="bg-wine bg-opacity-10 border-start border-3 border-danger p-2 mb-2 rounded-end small text-sub fst-italic">
                      &quot;{section.epigraph}&quot;
                    </div>
                  )}

                  {section.citation && (
                    <div className="small fw-semibold text-warning mb-2">
                      📍 {section.citation}
                    </div>
                  )}

                  {section.intentions && section.intentions.length > 0 && (
                    <div className="d-flex flex-column gap-2 my-2">
                      <p className="small text-sub mb-1 font-monospace">
                        Niatan Doa Umat ({section.intentions.length} niatan):
                      </p>
                      {section.intentions.map((item, iIndex) => (
                        <div
                          key={iIndex}
                          className="bg-black bg-opacity-20 p-2 rounded small border border-secondary-subtle"
                        >
                          <strong className="text-white d-block mb-1">
                            {item.title}
                          </strong>
                          <span className="text-sub">
                            {item.body || "(Tanpa badan doa)"}
                          </span>
                        </div>
                      ))}
                    </div>
                  )}

                  {section.body &&
                    section.body.length > 0 &&
                    !section.intentions && (
                      <div className="small text-sub mt-2">
                        {section.body.map((line, lIdx) => (
                          <p key={lIdx} className="mb-1">
                            {line}
                          </p>
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
            {parseResult.rawParagraphs.map((p, idx) => (
              <div
                key={idx}
                className="bg-black bg-opacity-20 p-2 rounded border border-secondary-subtle d-flex flex-column gap-1"
              >
                <div className="d-flex align-items-center justify-content-between small">
                  <span className="text-muted-custom font-monospace">
                    #{idx + 1}
                  </span>
                  <div className="d-flex gap-1">
                    {p.isBold && (
                      <span className="badge bg-danger bg-opacity-30 text-white">
                        Bold
                      </span>
                    )}
                    {p.isUnderline && (
                      <span className="badge bg-primary bg-opacity-30 text-white">
                        Underline
                      </span>
                    )}
                    {p.isItalic && (
                      <span className="badge bg-warning bg-opacity-30 text-white">
                        Italic
                      </span>
                    )}
                  </div>
                </div>
                <p className="mb-0 small text-white">{p.rawText}</p>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
