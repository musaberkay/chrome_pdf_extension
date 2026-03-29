/**
 * Convert PDF files to DOCX via a locally running Stirling PDF instance.
 * Requires Stirling PDF running at the configured URL (default: http://localhost:8080).
 * Each PDF produces one DOCX file — returns an array so the caller can trigger
 * one download per file.
 *
 * Setup: https://github.com/Stirling-Tools/Stirling-PDF
 *   docker run -d -p 8080:8080 stirlingtools/stirling-pdf:latest
 *
 * @param {File[]} files
 * @param {string} baseUrl  e.g. "http://localhost:8080"
 * @returns {Promise<Array<{ blob: Blob, name: string }>>}
 */
export async function convertViaStirlingPdf(files, baseUrl) {
  const endpoint = `${baseUrl.replace(/\/$/, "")}/api/v1/convert/pdf/word`;
  const results = [];

  for (const file of files) {
    const form = new FormData();
    form.append("fileInput", file);

    const res = await fetch(endpoint, { method: "POST", body: form });

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(
        `Stirling PDF returned ${res.status}${body ? `: ${body.slice(0, 120)}` : ""}`
      );
    }

    results.push({
      blob: await res.blob(),
      name: file.name.replace(/\.pdf$/i, ".docx"),
    });
  }

  return results;
}
