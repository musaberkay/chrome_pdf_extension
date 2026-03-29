/**
 * Convert PDF files to DOCX via ConvertAPI (https://www.convertapi.com).
 *
 * Free tier: 250 conversions on signup — no credit card required.
 * Get a free key at: https://www.convertapi.com/a/auth/register
 *
 * @param {File[]} files
 * @param {string} apiKey  Your ConvertAPI secret key
 * @returns {Promise<Array<{ blob: Blob, name: string }>>}
 */
export async function convertViaConvertApi(files, apiKey) {
  const results = [];

  for (const file of files) {
    const form = new FormData();
    form.append("File", file);

    const res = await fetch("https://v2.convertapi.com/convert/pdf/to/docx", {
      method: "POST",
      headers: { Authorization: `Bearer ${apiKey}` },
      body: form,
    });

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      const msg = body.slice(0, 200);
      if (res.status === 401 || res.status === 403) {
        throw new Error("ConvertAPI: invalid or expired API key.");
      }
      if (res.status === 422) {
        throw new Error("ConvertAPI: no conversions remaining on free plan.");
      }
      throw new Error(`ConvertAPI error ${res.status}${msg ? `: ${msg}` : ""}`);
    }

    const data = await res.json();
    const entry = data?.Files?.[0];
    if (!entry?.FileData) {
      throw new Error("ConvertAPI: unexpected response — missing FileData.");
    }

    // Decode base64 → Blob
    const binary = atob(entry.FileData);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

    results.push({
      blob: new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
      name: file.name.replace(/\.pdf$/i, ".docx"),
    });
  }

  return results;
}
