"use client";

import { useState } from "react";
import WordUpload from "./components/WordUpload";

// ponytail: step 1 upload with visual confirmation feedback
export default function Home() {
  const [confirmedFile, setConfirmedFile] = useState<File | null>(null);

  const handleFileConfirmed = (file: File) => {
    setConfirmedFile(file);
  };

  const handleReset = () => {
    setConfirmedFile(null);
  };

  return (
    <main className="min-vh-100 d-flex flex-column align-items-center justify-content-center py-4">
      {confirmedFile ? (
        <div className="w-100 max-w-lg mx-auto px-3 py-4 text-center">
          <div className="bg-panel p-4 rounded-4 text-center">
            <div className="d-inline-flex align-items-center justify-content-center bg-success bg-opacity-10 text-success rounded-circle p-3 mb-3">
              <svg width="32" height="32" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
                <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
              </svg>
            </div>
            <h2 className="h5 text-white mb-2 font-serif">Berkas Berhasil Diterima</h2>
            <p className="text-sub small mb-3">
              <strong className="text-white">{confirmedFile.name}</strong> siap untuk tahap berikutnya (Tinjau & Ekspor).
            </p>
            <button
              onClick={handleReset}
              className="btn btn-outline-light rounded-pill px-4 py-2 small"
            >
              Pilih Berkas Lain
            </button>
          </div>
        </div>
      ) : (
        <WordUpload onFileConfirmed={handleFileConfirmed} />
      )}
    </main>
  );
}


