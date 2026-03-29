import { PDFDocument } from "pdf-lib";

/**
 * @param {File[]} files
 * @returns {Promise<Blob>}
 */
export async function mergePdfs(files) {
  const merged = await PDFDocument.create();
  for (const file of files) {
    const bytes = await file.arrayBuffer();
    const doc = await PDFDocument.load(bytes);
    const indices = doc.getPageIndices();
    const pages = await merged.copyPages(doc, indices);
    pages.forEach((p) => merged.addPage(p));
  }
  const out = await merged.save();
  return new Blob([out], { type: "application/pdf" });
}
