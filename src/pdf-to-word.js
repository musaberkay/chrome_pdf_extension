import * as pdfjsLib from "pdfjs-dist";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
} from "docx";

let workerConfigured = false;

export function configurePdfWorker() {
  if (workerConfigured) return;
  workerConfigured = true;
  if (typeof chrome !== "undefined" && chrome.runtime?.getURL) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL(
      "pdf.worker.min.mjs"
    );
  }
}

function textFromTextContent(textContent) {
  const items = [...textContent.items].sort((a, b) => {
    const ay = a.transform[5];
    const by = b.transform[5];
    if (Math.abs(ay - by) > 5) return by - ay;
    return a.transform[4] - b.transform[4];
  });
  return items.map((i) => i.str).join(" ");
}

function baseName(name) {
  return name.replace(/\.pdf$/i, "");
}

/**
 * @param {File[]} files
 * @returns {Promise<Blob>}
 */
export async function convertPdfsToDocx(files) {
  configurePdfWorker();
  const children = [];

  for (let fi = 0; fi < files.length; fi++) {
    const file = files[fi];
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun(baseName(file.name))],
      })
    );

    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    for (let pi = 1; pi <= pdf.numPages; pi++) {
      const page = await pdf.getPage(pi);
      const textContent = await page.getTextContent();
      const text = textFromTextContent(textContent).trim();
      const fallback =
        pdf.numPages > 1
          ? `[Page ${pi}: no extractable text]`
          : "[No extractable text]";
      children.push(
        new Paragraph({
          children: [new TextRun(text || fallback)],
        })
      );
    }

    if (fi < files.length - 1) {
      children.push(new Paragraph({ children: [new TextRun("")] }));
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children,
      },
    ],
  });

  return Packer.toBlob(doc);
}
