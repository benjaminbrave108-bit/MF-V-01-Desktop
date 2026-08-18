import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Maliye-Finans Önizleme",
  description: "Maliye ve finans yönetimi önizleme uygulaması.",
  other: {
    "codex-preview": "development",
  },
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="tr">
      <body>{children}</body>
    </html>
  );
}
