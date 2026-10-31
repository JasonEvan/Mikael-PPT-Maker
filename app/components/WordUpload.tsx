"use client";

import {
  useState,
  useRef,
  useId,
  ChangeEvent,
  DragEvent,
  KeyboardEvent,
} from "react";

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
  compact?: boolean;
}

export default function WordUpload({
  onFileConfirmed,
  compact = false,
}: WordUploadProps) {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isDragActive, setIsDragActive] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const instructionsId = useId();
  const inputId = useId();

  const validateAndSelectFile = (file: File) => {
    // Validate Word extensions and MIME types. Legacy .doc files use the
    // compound binary format, while .docx files are ZIP-based.
    const isDocx = file.name.toLowerCase().endsWith(".docx");
    const isDoc = file.name.toLowerCase().endsWith(".doc");
    const validMimeTypes = [
      "application/msword",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/zip",
    ];
    const isWordMime = !file.type || validMimeTypes.includes(file.type);

    if ((!isDoc && !isDocx) || !isWordMime) {
      setErrorMessage(
        "Berkas tidak valid. Harap unggah berkas Microsoft Word bertipe .doc atau .docx.",
      );
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
    <div className={`word-upload ${compact ? "word-upload-compact" : ""}`}>
      {/* Hidden Native File Input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        onChange={handleFileChange}
        className="d-none"
        id={inputId}
        aria-label="Pilih berkas Word .doc atau .docx"
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
        className={`dropzone ${isDragActive ? "drag-active" : ""}`}
        aria-describedby={selectedFile ? undefined : instructionsId}
      >
        {selectedFile ? (
          /* Confirmation State */
          <div className="selected-file">
            <div className="file-icon">
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
            <h3 title={selectedFile.name}>{selectedFile.name}</h3>
            <p>{formatFileSize(selectedFile.size)} · Dokumen Word</p>
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                handleRemoveFile();
              }}
              className="quiet-button"
            >
              Ganti Berkas
            </button>
          </div>
        ) : (
          /* Resting / Dragging State */
          <div className="dropzone-content" id={instructionsId}>
            <div className="upload-icon">
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
            <p className="dropzone-title">
              Tarik dan lepas berkas <strong>.doc atau .docx</strong>
            </p>
            <p className="dropzone-hint">
              atau klik untuk memilih dari perangkat
            </p>
          </div>
        )}
      </div>

      {/* Inline Validation Error Message */}
      {errorMessage && (
        <div className="form-alert" role="alert">
          {errorMessage}
        </div>
      )}

      <div className="upload-meta">
        <span>.doc atau .docx</span>
        <span>Teks lagu & pengumuman tidak disertakan</span>
      </div>

      <button
        type="button"
        disabled={!selectedFile}
        onClick={handleContinue}
        className="primary-button"
      >
        <span>Proses berkas</span>
        <svg
          width="18"
          height="18"
          viewBox="0 0 24 24"
          fill="none"
          aria-hidden="true"
        >
          <path d="m9 18 6-6-6-6" />
        </svg>
      </button>
    </div>
  );
}
