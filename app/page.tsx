import { readFileSync } from "node:fs";
import { join } from "node:path";

function getPageContent() {
  const source = readFileSync(join(process.cwd(), "index.html"), "utf8");
  const body = source.match(/<body[^>]*>([\s\S]*?)<\/body>/i)?.[1];

  if (!body) {
    throw new Error("Konten <body> tidak ditemukan di index.html");
  }

  return body;
}

export default function HomePage() {
  return <div dangerouslySetInnerHTML={{ __html: getPageContent() }} />;
}
