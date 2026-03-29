import * as pdfjsLib from "pdfjs-dist";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";

let workerConfigured = false;

export function configurePdfWorker() {
  if (workerConfigured) return;
  workerConfigured = true;
  if (typeof chrome !== "undefined" && chrome.runtime?.getURL) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL("pdf.worker.min.mjs");
  }
}

function baseName(name) {
  return name.replace(/\.pdf$/i, "");
}

/** Extract font size from a PDF transform matrix. */
function getFontSize(transform) {
  return Math.hypot(transform[0], transform[1]) || Math.abs(transform[3]);
}

function isBold(fontName) {
  return /bold/i.test(fontName ?? "");
}

function isItalic(fontName) {
  return /italic|oblique/i.test(fontName ?? "");
}

/**
 * Find the most common font size across items — this is the body text size.
 * @param {{ fontSize: number }[]} items
 */
function detectBodyFontSize(items) {
  const freq = {};
  for (const { fontSize } of items) {
    const rounded = Math.round(fontSize);
    if (rounded > 0) freq[rounded] = (freq[rounded] || 0) + 1;
  }
  const entries = Object.entries(freq);
  if (!entries.length) return 12;
  return Number(entries.sort((a, b) => b[1] - a[1])[0][0]);
}

/**
 * Group annotated text items into lines (items sharing the same Y coordinate ± 3pt).
 * @param {{ x: number, y: number }[]} items
 * @returns {Array<typeof items>}
 */
function groupIntoLines(items) {
  if (!items.length) return [];

  const sorted = [...items].sort((a, b) => {
    if (Math.abs(a.y - b.y) > 3) return b.y - a.y; // top → bottom
    return a.x - b.x;                               // left → right
  });

  const lines = [];
  let current = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    if (Math.abs(sorted[i].y - current[0].y) <= 3) {
      current.push(sorted[i]);
    } else {
      lines.push(current);
      current = [sorted[i]];
    }
  }
  lines.push(current);
  return lines;
}

/**
 * Choose a heading level based on font size relative to body, or undefined for body text.
 * @param {number} avgSize
 * @param {number} bodySize
 * @returns {string | undefined}
 */
function headingLevel(avgSize, bodySize) {
  if (avgSize >= bodySize * 1.8) return HeadingLevel.HEADING_1;
  if (avgSize >= bodySize * 1.4) return HeadingLevel.HEADING_2;
  if (avgSize >= bodySize * 1.15) return HeadingLevel.HEADING_3;
  return undefined;
}

/**
 * Convert a group of lines (one logical paragraph) into a docx Paragraph.
 * @param {Array<Array<object>>} paraLines
 * @param {number} bodySize
 * @returns {Paragraph | null}
 */
function buildParagraph(paraLines, bodySize) {
  const allItems = paraLines.flat();
  const text = allItems.map((i) => i.str).join("").trim();
  if (!text) return null;

  const avgSize = allItems.reduce((s, i) => s + i.fontSize, 0) / allItems.length;
  const level = headingLevel(avgSize, bodySize);

  const runs = allItems.map(
    (item) =>
      new TextRun({
        text: item.str,
        bold: item.isBold || !!level,
        italics: item.isItalic,
        // docx size is in half-points; clamp to a sane range
        size: Math.max(16, Math.min(96, Math.round(item.fontSize * 2))),
      })
  );

  return new Paragraph({
    heading: level,
    children: runs,
    spacing: { after: level ? 160 : 80 },
  });
}

/**
 * Turn a pdfjs textContent object into an array of docx Paragraphs.
 * @param {import("pdfjs-dist").TextContent} textContent
 * @param {number} pageIndex  1-based, used for fallback message
 * @param {boolean} isOnlyPage
 */
function pageToDocxParagraphs(textContent, pageIndex, isOnlyPage) {
  const annotated = textContent.items
    .filter((item) => item.str)
    .map((item) => ({
      str: item.str,
      x: item.transform[4],
      y: item.transform[5],
      fontSize: getFontSize(item.transform),
      isBold: isBold(item.fontName),
      isItalic: isItalic(item.fontName),
    }));

  if (!annotated.length) {
    const msg = isOnlyPage ? "[No extractable text]" : `[Page ${pageIndex}: no extractable text]`;
    return [new Paragraph({ children: [new TextRun(msg)] })];
  }

  const bodySize = detectBodyFontSize(annotated);
  const lines = groupIntoLines(annotated);

  // Detect paragraph breaks: a gap larger than 1.4× the average line spacing
  let totalGap = 0;
  for (let i = 1; i < lines.length; i++) {
    totalGap += Math.abs(lines[i - 1][0].y - lines[i][0].y);
  }
  const avgLineHeight = lines.length > 1 ? totalGap / (lines.length - 1) : 14;
  const paraBreakThreshold = avgLineHeight * 1.4;

  // Group lines into paragraphs
  const paragraphGroups = [];
  let currentGroup = [lines[0]];

  for (let i = 1; i < lines.length; i++) {
    const gap = Math.abs(lines[i - 1][0].y - lines[i][0].y);
    if (gap > paraBreakThreshold) {
      paragraphGroups.push(currentGroup);
      currentGroup = [];
    }
    currentGroup.push(lines[i]);
  }
  paragraphGroups.push(currentGroup);

  return paragraphGroups
    .map((group) => buildParagraph(group, bodySize))
    .filter(Boolean);
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
        spacing: { after: 200 },
      })
    );

    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

    for (let pi = 1; pi <= pdf.numPages; pi++) {
      const page = await pdf.getPage(pi);
      const textContent = await page.getTextContent();
      const paragraphs = pageToDocxParagraphs(textContent, pi, pdf.numPages === 1);
      children.push(...paragraphs);

      // Blank line between pages
      if (pi < pdf.numPages) {
        children.push(new Paragraph({ children: [] }));
      }
    }

    if (fi < files.length - 1) {
      children.push(new Paragraph({ children: [] }));
    }
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  return Packer.toBlob(doc);
}
