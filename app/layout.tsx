import type { Metadata } from "next";
import "./globals.css";

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
      <body>{children}</body>
    </html>
  );
}
