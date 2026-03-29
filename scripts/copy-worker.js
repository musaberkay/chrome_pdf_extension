const fs = require("fs");
const path = require("path");

const src = path.join(
  __dirname,
  "..",
  "node_modules",
  "pdfjs-dist",
  "build",
  "pdf.worker.min.mjs"
);
const dest = path.join(__dirname, "..", "pdf.worker.min.mjs");
fs.copyFileSync(src, dest);
