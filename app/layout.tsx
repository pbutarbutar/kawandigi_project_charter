import type { Metadata } from "next";
import type { ReactNode } from "react";
import { readFileSync } from "node:fs";
import { join } from "node:path";

export const metadata: Metadata = {
  title: "Project Charter - Koperasi Digital KawanDigi",
  description:
    "Project Charter inisiasi Koperasi Digital KawanDigi di Brawijaya Hospital Antasari.",
};

function getPageStyles() {
  const source = readFileSync(join(process.cwd(), "index.html"), "utf8");
  return source.match(/<style[^>]*>([\s\S]*?)<\/style>/i)?.[1] ?? "";
}

export default function RootLayout({ children }: Readonly<{ children: ReactNode }>) {
  return (
    <html lang="id">
      <body>
        <style dangerouslySetInnerHTML={{ __html: getPageStyles() }} />
        {children}
      </body>
    </html>
  );
}
