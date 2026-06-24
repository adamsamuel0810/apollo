import type { Metadata } from "next";
import { Carlito } from "next/font/google";
import "./globals.css";

const carlito = Carlito({
  weight: ["400", "700"],
  style: ["normal", "italic"],
  subsets: ["latin"],
  display: "swap",
  variable: "--slide-font-family",
});

export const metadata: Metadata = {
  title: "ACME Brand Compliance Checker",
  description:
    "Upload a PowerPoint and review brand-guideline violations per slide with accept/reject controls.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className={carlito.variable}>{children}</body>
    </html>
  );
}
