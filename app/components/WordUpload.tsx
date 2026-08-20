"use client";

import { useState, useRef, ChangeEvent, DragEvent, KeyboardEvent } from "react";

// ponytail: native helper for human-readable file sizes without extra dependencies
function formatFileSize(bytes: number): string {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
}

interface WordUploadProps {
  onFileConfirmed?: (file: File) => void;
}

export default function WordUpload({ onFileConfirmed }: WordUploadProps) {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isDragActive, setIsDragActive] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const validateAndSelectFile = (file: File) => {
    // Validate file extension (.docx) and MIME type
    const isDocxExtension = file.name.toLowerCase().endsWith(".docx");
    const validMimeType =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const isDocxMime = !file.type || file.type === validMimeType || file.type === "application/zip";

    if (!isDocxExtension || (!isDocxMime && file.type)) {
      setErrorMessage("Berkas tidak valid. Harap unggah berkas Microsoft Word bertipe .docx.");
      setSelectedFile(null);
      return;
    }

    setErrorMessage(null);
    setSelectedFile(file);
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      validateAndSelectFile(files[0]);
    }
  };

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(true);
  };

  const handleDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      validateAndSelectFile(files[0]);
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fileInputRef.current?.click();
    }
  };

  const handleRemoveFile = () => {
    setSelectedFile(null);
    setErrorMessage(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const handleContinue = () => {
    if (selectedFile && onFileConfirmed) {
      onFileConfirmed(selectedFile);
    }
  };

  return (
    <div className="w-100 max-w-lg mx-auto px-3 py-4">
      {/* Main Title & Context */}
      <div className="text-center mb-4">
        <h1 className="font-serif display-6 text-white mb-2 fw-normal">
          Otomatisasi Slide Liturgi
        </h1>
        <p className="text-sub small mb-0 px-2">
          Unggah berkas Word (.docx) liturgi Misa minggu ini untuk diproses dan dicocokkan secara otomatis dengan template PowerPoint.
        </p>
      </div>

      {/* Hidden Native File Input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        onChange={handleFileChange}
        className="d-none"
        id="docx-file-input"
        aria-label="Pilih berkas Word .docx"
      />

      {/* Dropzone Container */}
      <div
        role="button"
        tabIndex={0}
        onClick={() => fileInputRef.current?.click()}
        onKeyDown={handleKeyDown}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={`dropzone p-4 text-center ${isDragActive ? "drag-active" : ""}`}
        aria-describedby="dropzone-instructions"
      >
        {selectedFile ? (
          /* Confirmation State */
          <div className="py-3 px-2">
            <div className="d-inline-flex align-items-center justify-content-center bg-white bg-opacity-10 text-white rounded-circle p-3 mb-3">
              <svg
                width="32"
                height="32"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                strokeWidth="1.5"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                />
              </svg>
            </div>
            <h2 className="h6 text-white text-truncate mb-1 px-3" title={selectedFile.name}>
              {selectedFile.name}
            </h2>
            <p className="text-sub small mb-3">{formatFileSize(selectedFile.size)}</p>
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                handleRemoveFile();
              }}
              className="btn btn-sm btn-outline-light rounded-pill px-3 py-1 text-sub border-secondary-subtle"
            >
              Ganti Berkas
            </button>
          </div>
        ) : (
          /* Resting / Dragging State */
          <div className="py-4 px-2" id="dropzone-instructions">
            <div className="d-inline-flex align-items-center justify-content-center text-sub rounded-circle p-3 mb-3">
              <svg
                width="40"
                height="40"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                strokeWidth="1.5"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                />
              </svg>
            </div>
            <p className="text-white fw-medium mb-1">
              Tarik & lepas berkas <span className="text-white fw-bold">.docx</span> di sini
            </p>
            <p className="text-sub small mb-0">atau klik untuk memilih dari perangkat Anda</p>
          </div>
        )}
      </div>

      {/* Inline Validation Error Message */}
      {errorMessage && (
        <div
          className="alert alert-danger mt-3 py-2 px-3 small border-0 bg-danger bg-opacity-10 text-danger text-center rounded-3"
          role="alert"
        >
          {errorMessage}
        </div>
      )}

      {/* Scope Disclaimer */}
      <p className="text-center text-muted-custom small mt-4 mb-4" style={{ fontSize: "0.82rem" }}>
        *Catatan: Teks lagu Misa dan pengumuman gereja diproses secara terpisah dan tidak termasuk dalam otomatisasi ini.
      </p>

      {/* Primary Action Button */}
      <button
        type="button"
        disabled={!selectedFile}
        onClick={handleContinue}
        className="btn btn-wine btn-lg w-100 py-3 rounded-3 fw-semibold"
      >
        Proses Berkas
      </button>
    </div>
  );
}
