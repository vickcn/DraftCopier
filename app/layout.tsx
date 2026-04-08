import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "DraftCopier",
  description: "DOCX + Excel 批次產生 Gmail 草稿",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="zh-Hant">
      <body>{children}</body>
    </html>
  );
}
