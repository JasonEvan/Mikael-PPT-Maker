"use client";

import { useState } from "react";
import WordUpload from "./WordUpload";
import { parseDocFile, parseDocxFile } from "../utils/docxParser";
import { EnglishMassParseResult } from "../utils/englishMassParser";

type Homily = NonNullable<EnglishMassParseResult["homily"]>;

interface HomilyUploadProps {
  homily?: Homily;
  onChange: (homily: Homily | undefined) => void;
  onBusyChange: (busy: boolean) => void;
}

export default function HomilyUpload({
  homily,
  onChange,
  onBusyChange,
}: HomilyUploadProps) {
  const [isParsing, setIsParsing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleUpload = async (file: File) => {
    setIsParsing(true);
    onBusyChange(true);
    setError(null);
    try {
      const result = file.name.toLowerCase().endsWith(".doc")
        ? await parseDocFile(file)
        : await parseDocxFile(file);
      const paragraphs = result.rawParagraphs
        .map((paragraph) => paragraph.rawText.trim())
        .filter(Boolean);
      if (!paragraphs.length)
        throw new Error(
          "Tidak ada teks homili yang dapat dibaca dari dokumen.",
        );
      onChange({ fileName: file.name, paragraphs });
    } catch (error) {
      setError(
        error instanceof Error
          ? error.message
          : "Gagal membaca dokumen homili.",
      );
    } finally {
      setIsParsing(false);
      onBusyChange(false);
    }
  };

  return (
    <section className="mt-4" aria-label="Dokumen homili">
      <h2 className="h5">Teks homili (opsional)</h2>
      <p className="small">
        Unggah dokumen terpisah berisi teks homili saja. Seluruh teks akan
        ditampilkan setelah HOMILY.
      </p>
      {isParsing ? (
        <p role="status">Membaca teks homili…</p>
      ) : (
        <WordUpload compact onFileConfirmed={handleUpload} />
      )}
      {error && (
        <div className="form-alert" role="alert">
          {error}
        </div>
      )}
      {homily && (
        <div className="mt-2">
          <p className="small">
            Homili aktif: <strong>{homily.fileName}</strong>
          </p>
          <button
            type="button"
            className="quiet-button"
            disabled={isParsing}
            onClick={() => onChange(undefined)}
          >
            Hapus homili
          </button>
        </div>
      )}
    </section>
  );
}
