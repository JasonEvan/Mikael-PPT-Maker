import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Liturgi Slides — Otomatisasi PowerPoint Misa",
  description:
    "Ubah dokumen liturgi mingguan menjadi presentasi PowerPoint siap tayang.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="id">
      <body>{children}</body>
    </html>
  );
}
