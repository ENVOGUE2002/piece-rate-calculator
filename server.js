const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { spawn } = require("child_process");
const { URL } = require("url");

const HOST = process.env.HOST || "127.0.0.1";
const PORT = Number(process.env.PORT || 3100);
const DEFAULT_TALLY_URL = process.env.TALLY_URL || "http://127.0.0.1:9000";
const TALLY_TIMEOUT_MS = Number(process.env.TALLY_TIMEOUT_MS || 5000);
const ROOT_DIR = __dirname;
const BUNDLED_PYTHON = "C:\\Users\\Lenovo\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".toml": "text/plain; charset=utf-8",
  ".sql": "text/plain; charset=utf-8",
  ".cmd": "text/plain; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};

const server = http.createServer(async (req, res) => {
  try {
    const requestUrl = new URL(req.url, `http://${req.headers.host}`);

    if (requestUrl.pathname === "/health") {
      return sendJson(res, 200, {
        ok: true,
        appUrl: `http://${HOST}:${PORT}`,
        tallyProxy: `${requestUrl.origin}/tally`,
        defaultTallyUrl: DEFAULT_TALLY_URL
      });
    }

    if (requestUrl.pathname === "/tally-test") {
      return handleTallyTest(req, res, requestUrl);
    }

    if (requestUrl.pathname === "/tally") {
      return handleTallyProxy(req, res, requestUrl);
    }

    if (requestUrl.pathname === "/extract-washcare-report") {
      return handleWashcareReportExtract(req, res);
    }

    if (requestUrl.pathname === "/extract-pdf-text") {
      return handleGenericPdfTextExtract(req, res);
    }

    if (requestUrl.pathname === "/launch-garment-erp") {
      return handleLaunchGarmentErp(req, res);
    }

    if (requestUrl.pathname === "/sync-erp") {
      return handleSyncErp(req, res);
    }

    if (requestUrl.pathname === "/erp-pickup") {
      return handleErpPickup(req, res);
    }

    if (requestUrl.pathname === "/proxy-image") {
      return handleProxyImage(req, res, requestUrl);
    }

    return serveStaticFile(req, res, requestUrl);
  } catch (error) {
    console.error(error);
    sendJson(res, 500, { error: "Server error", details: error.message });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Piece Rate app running at http://${HOST}:${PORT}`);
  console.log(`Tally proxy ready at http://${HOST}:${PORT}/tally`);
  console.log(`Default Tally target: ${DEFAULT_TALLY_URL}`);
});

async function handleTallyProxy(req, res, requestUrl) {
  if (req.method !== "POST") {
    res.writeHead(405, { Allow: "POST" });
    res.end("Method Not Allowed");
    return;
  }

  const target = requestUrl.searchParams.get("target") || DEFAULT_TALLY_URL;
  let targetUrl;
  try {
    targetUrl = new URL(target);
  } catch {
    return sendJson(res, 400, { error: "Invalid Tally target URL." });
  }

  const body = await readRequestBody(req);
  try {
    const proxyResponse = await fetchWithTimeout(targetUrl, {
      method: "POST",
      headers: {
        "Content-Type": req.headers["content-type"] || "text/xml; charset=utf-8"
      },
      body
    });

    const responseText = await proxyResponse.text();
    res.writeHead(proxyResponse.status, {
      "Content-Type": proxyResponse.headers.get("content-type") || "text/xml; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*"
    });
    res.end(responseText);
  } catch (error) {
    sendJson(res, 502, {
      error: "Could not connect to Tally through the proxy.",
      details: error.message,
      target: targetUrl.toString()
    });
  }
}

async function handleTallyTest(req, res, requestUrl) {
  if (!["GET", "POST"].includes(req.method)) {
    res.writeHead(405, { Allow: "GET, POST" });
    res.end("Method Not Allowed");
    return;
  }
  const target = requestUrl.searchParams.get("target") || DEFAULT_TALLY_URL;
  let targetUrl;
  try {
    targetUrl = new URL(target);
  } catch {
    return sendJson(res, 400, { ok: false, error: "Invalid Tally target URL." });
  }

  const xml = `<?xml version="1.0" encoding="utf-8"?>
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>List of Accounts</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>`;

  try {
    const response = await fetchWithTimeout(targetUrl, {
      method: "POST",
      headers: { "Content-Type": "text/xml; charset=utf-8" },
      body: xml
    });
    const text = await response.text();
    if (!response.ok) {
      return sendJson(res, response.status, {
        ok: false,
        error: `Tally returned HTTP ${response.status}.`,
        target: targetUrl.toString(),
        snippet: text.slice(0, 300)
      });
    }
    return sendJson(res, 200, {
      ok: true,
      target: targetUrl.toString(),
      snippet: text.slice(0, 300)
    });
  } catch (error) {
    return sendJson(res, 502, {
      ok: false,
      error: "Could not connect to Tally. Make sure Tally is open and listening on the configured port.",
      details: error.message,
      target: targetUrl.toString()
    });
  }
}

async function handleWashcareReportExtract(req, res) {
  if (req.method !== "POST") {
    res.writeHead(405, { Allow: "POST" });
    res.end("Method Not Allowed");
    return;
  }

  const body = await readRequestBody(req);
  if (!body?.length) {
    return sendJson(res, 400, { error: "No PDF file received." });
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "washcare-report-"));
  const pdfPath = path.join(tempDir, "report.pdf");
  fs.writeFileSync(pdfPath, body);

  try {
    const extracted = await extractPdfText(pdfPath);
    return sendJson(res, 200, extracted);
  } catch (error) {
    return sendJson(res, 500, {
      error: "Could not extract text from the PDF.",
      details: error.message
    });
  } finally {
    try { fs.unlinkSync(pdfPath); } catch {}
    try { fs.rmdirSync(tempDir); } catch {}
  }
}

async function handleGenericPdfTextExtract(req, res) {
  if (req.method !== "POST") {
    res.writeHead(405, { Allow: "POST" });
    res.end("Method Not Allowed");
    return;
  }

  const body = await readRequestBody(req);
  if (!body?.length) {
    return sendJson(res, 400, { error: "No PDF file received." });
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "generic-pdf-"));
  const pdfPath = path.join(tempDir, "document.pdf");
  fs.writeFileSync(pdfPath, body);

  try {
    const extracted = await extractPdfText(pdfPath);
    return sendJson(res, 200, extracted);
  } catch (error) {
    return sendJson(res, 500, {
      error: "Could not extract text from the PDF.",
      details: error.message
    });
  } finally {
    try { fs.unlinkSync(pdfPath); } catch {}
    try { fs.rmdirSync(tempDir); } catch {}
  }
}

async function handleLaunchGarmentErp(req, res) {
  if (req.method !== "POST") {
    res.writeHead(405, { Allow: "POST" });
    res.end("Method Not Allowed");
    return;
  }

  const erpDir = path.join(ROOT_DIR, "garment_erp");
  const mainScript = path.join(erpDir, "main.py");
  if (!fs.existsSync(mainScript)) {
    return sendJson(res, 404, {
      ok: false,
      error: "garment_erp/main.py not found.",
      expected: mainScript
    });
  }

  // Pick a Python that actually has PyQt6 + sqlalchemy installed. The
  // bundled runtime Python (3.12) does NOT have these, so the spawned
  // process dies immediately. Probe candidates in priority order.
  let chosen;
  try {
    chosen = await pickPythonForErp();
  } catch (error) {
    return sendJson(res, 500, {
      ok: false,
      error:
        "No Python interpreter with PyQt6 and SQLAlchemy was found. Install the desktop module dependencies: py -3 -m pip install -r garment_erp/requirements.txt",
      details: error.message
    });
  }

  let stderrBuffer = "";
  try {
    const child = spawn(chosen.command, [...chosen.args, mainScript], {
      cwd: erpDir,
      detached: true,
      stdio: ["ignore", "ignore", "pipe"],
      windowsHide: false
    });
    child.stderr.on("data", (chunk) => {
      stderrBuffer += chunk.toString();
      if (stderrBuffer.length > 4000) {
        stderrBuffer = stderrBuffer.slice(-4000);
      }
    });
    child.on("error", (error) => {
      console.error("Failed to launch Garment ERP:", error);
    });

    // Give the process a brief window to crash on imports. If it's still
    // alive after the grace period we treat it as successfully launched.
    const earlyExit = await new Promise((resolve) => {
      let resolved = false;
      const finish = (payload) => {
        if (resolved) return;
        resolved = true;
        resolve(payload);
      };
      child.once("exit", (code, signal) => finish({ code, signal }));
      setTimeout(() => finish(null), 1500);
    });

    if (earlyExit) {
      return sendJson(res, 500, {
        ok: false,
        error:
          "Garment ERP exited immediately after launch. Check the desktop module dependencies.",
        exitCode: earlyExit.code,
        signal: earlyExit.signal,
        python: chosen.command,
        stderr: stderrBuffer.trim().slice(-1200)
      });
    }

    child.stderr.removeAllListeners("data");
    child.unref();
    return sendJson(res, 200, {
      ok: true,
      pid: child.pid,
      python: chosen.command,
      script: mainScript
    });
  } catch (error) {
    return sendJson(res, 500, {
      ok: false,
      error: "Could not spawn Python process.",
      details: error.message,
      python: chosen.command
    });
  }
}

async function handleSyncErp(req, res) {
  if (req.method !== "POST") {
    res.writeHead(405, { Allow: "POST" });
    res.end("Method Not Allowed");
    return;
  }

  // Accept the same JSON shape the existing "Download Style Engineering
  // Bridge" button produces. We write it to a temp file and hand it to
  // the Python importer module.
  const raw = await readRequestBody(req);
  if (!raw?.length) {
    return sendJson(res, 400, { ok: false, error: "Empty payload." });
  }

  let payload;
  try {
    payload = JSON.parse(raw.toString("utf-8"));
  } catch (error) {
    return sendJson(res, 400, {
      ok: false,
      error: "Invalid JSON payload.",
      details: error.message
    });
  }

  const styles = Array.isArray(payload?.styles) ? payload.styles : [];
  if (!styles.length) {
    return sendJson(res, 400, {
      ok: false,
      error: "Payload has no styles to sync."
    });
  }

  const erpDir = path.join(ROOT_DIR, "garment_erp");
  if (!fs.existsSync(erpDir)) {
    return sendJson(res, 404, {
      ok: false,
      error: "garment_erp directory not found.",
      expected: erpDir
    });
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "erp-sync-"));
  const payloadPath = path.join(tempDir, "bridge.json");
  fs.writeFileSync(payloadPath, JSON.stringify(payload), "utf-8");

  let chosen;
  try {
    chosen = await pickPythonForErp();
  } catch (error) {
    cleanupTempDir(tempDir);
    return sendJson(res, 500, {
      ok: false,
      error:
        "No Python interpreter with PyQt6 and SQLAlchemy was found. Install the desktop module dependencies: py -3 -m pip install -r garment_erp/requirements.txt",
      details: error.message
    });
  }

  try {
    const result = await runPythonImporter(chosen, erpDir, payloadPath);
    return sendJson(res, 200, {
      ok: true,
      synced_styles: styles.length,
      python: chosen.command,
      summary: parseBridgeSummary(result.stdout),
      stdout: result.stdout.slice(-800),
      stderr: result.stderr.slice(-800)
    });
  } catch (error) {
    return sendJson(res, 500, {
      ok: false,
      error: "Python importer failed.",
      details: error.message,
      python: chosen.command
    });
  } finally {
    cleanupTempDir(tempDir);
  }
}

function runPythonImporter(chosen, erpDir, payloadPath) {
  return new Promise((resolve, reject) => {
    const child = spawn(
      chosen.command,
      [
        ...chosen.args,
        "-m",
        "style_development.sample_data.import_piece_rate_export",
        payloadPath
      ],
      { cwd: erpDir, windowsHide: true }
    );
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout += chunk.toString(); });
    child.stderr.on("data", (chunk) => { stderr += chunk.toString(); });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) resolve({ stdout, stderr });
      else reject(new Error(stderr.trim() || `Importer exited with code ${code}`));
    });
  });
}

async function handleProxyImage(req, res, requestUrl) {
  // CORS-free image fetcher. The browser can't pull Firebase Storage URLs
  // into a canvas (no Access-Control-Allow-Origin), but Node can — so we
  // proxy the bytes back as an ArrayBuffer the JS app can dataUrl-encode.
  if (!["GET", "HEAD"].includes(req.method)) {
    res.writeHead(405, { Allow: "GET, HEAD" });
    res.end("Method Not Allowed");
    return;
  }
  const target = requestUrl.searchParams.get("url") || "";
  let targetUrl;
  try {
    targetUrl = new URL(target);
  } catch {
    return sendJson(res, 400, { error: "Invalid url parameter." });
  }
  if (!/^https?:$/.test(targetUrl.protocol)) {
    return sendJson(res, 400, { error: "Only http(s) URLs are allowed." });
  }
  // Tight allow-list to keep this from being a generic open proxy. Add more
  // hosts here if you start serving images from other CDNs.
  const allowedHosts = [
    "firebasestorage.googleapis.com",
    "storage.googleapis.com",
    "firebasestorage.app"
  ];
  const okHost = allowedHosts.some(
    (host) => targetUrl.hostname === host || targetUrl.hostname.endsWith(`.${host}`)
  );
  if (!okHost) {
    return sendJson(res, 403, { error: `Host not allowed: ${targetUrl.hostname}` });
  }
  try {
    const response = await fetchWithTimeout(targetUrl, { method: "GET" });
    if (!response.ok) {
      return sendJson(res, response.status, {
        error: `Upstream returned HTTP ${response.status}`
      });
    }
    const buffer = Buffer.from(await response.arrayBuffer());
    res.writeHead(200, {
      "Content-Type": response.headers.get("content-type") || "application/octet-stream",
      "Cache-Control": "private, max-age=300",
      "Access-Control-Allow-Origin": "*"
    });
    res.end(req.method === "HEAD" ? undefined : buffer);
  } catch (error) {
    return sendJson(res, 502, {
      error: "Proxy fetch failed.",
      details: error.message
    });
  }
}

function parseBridgeSummary(stdout) {
  // The Python importer prints `BRIDGE_SUMMARY {json}` on its last line so
  // we can show image-save stats in the UI instead of guessing.
  const match = String(stdout || "").match(/BRIDGE_SUMMARY\s+(\{.*\})/);
  if (!match) return null;
  try {
    return JSON.parse(match[1]);
  } catch {
    return null;
  }
}

function cleanupTempDir(dir) {
  try {
    for (const file of fs.readdirSync(dir)) {
      try { fs.unlinkSync(path.join(dir, file)); } catch {}
    }
    fs.rmdirSync(dir);
  } catch {}
}

async function handleErpPickup(req, res) {
  if (!["GET", "HEAD"].includes(req.method)) {
    res.writeHead(405, { Allow: "GET, HEAD" });
    res.end("Method Not Allowed");
    return;
  }

  const csvPath = path.join(ROOT_DIR, "erp_styles_export.csv");
  if (!fs.existsSync(csvPath)) {
    return sendJson(res, 200, { ok: true, fresh: false, rows: [] });
  }

  if (req.method === "HEAD") {
    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
    res.end();
    return;
  }

  let content;
  try {
    content = fs.readFileSync(csvPath, "utf-8");
  } catch (error) {
    return sendJson(res, 500, {
      ok: false,
      error: "Could not read erp_styles_export.csv.",
      details: error.message
    });
  }

  const rows = parseErpCsv(content);

  // Rename so the same file is never returned twice. The JS app keeps
  // the imported rows in its own state; this dance is just to avoid
  // re-importing on every page refresh.
  const timestamp = new Date()
    .toISOString()
    .replace(/[:.]/g, "-")
    .slice(0, 19);
  const consumedPath = path.join(ROOT_DIR, `erp_styles_export.consumed-${timestamp}.csv`);
  try {
    fs.renameSync(csvPath, consumedPath);
  } catch (error) {
    console.error("Could not archive consumed ERP CSV:", error.message);
  }

  return sendJson(res, 200, {
    ok: true,
    fresh: true,
    rows,
    consumed_as: path.basename(consumedPath)
  });
}

function parseErpCsv(text) {
  const lines = text.replace(/^﻿/, "").split(/\r?\n/).filter((line) => line.trim());
  if (!lines.length) return [];
  const headers = splitCsvLine(lines[0]);
  return lines.slice(1).map((line) => {
    const values = splitCsvLine(line);
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

function splitCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result.map((value) => value.trim());
}

async function pickPythonForErp() {
  // py.exe -3 first (Windows launcher — typically points at the user's
  // installed Python that actually has PyQt6). Bundled runtime second
  // (only useful if someone installed the deps into it).
  const candidates = [
    { command: "py", args: ["-3"] },
    { command: BUNDLED_PYTHON, args: [] },
    { command: "python", args: [] }
  ];

  let lastError = null;
  for (const candidate of candidates) {
    if (candidate.command === BUNDLED_PYTHON && !fs.existsSync(BUNDLED_PYTHON)) {
      continue;
    }
    try {
      await runProbe(candidate);
      return candidate;
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("No suitable Python interpreter found.");
}

function runProbe(candidate) {
  return new Promise((resolve, reject) => {
    const probe = spawn(
      candidate.command,
      [...candidate.args, "-c", "import PyQt6, sqlalchemy"],
      { windowsHide: true }
    );
    let stderr = "";
    probe.stderr.on("data", (chunk) => { stderr += chunk.toString(); });
    probe.on("error", reject);
    probe.on("close", (code) => {
      if (code === 0) resolve();
      else reject(new Error(stderr.trim() || `probe exited with code ${code}`));
    });
  });
}

async function serveStaticFile(req, res, requestUrl) {
  if (!["GET", "HEAD"].includes(req.method)) {
    res.writeHead(405, { Allow: "GET, HEAD" });
    res.end("Method Not Allowed");
    return;
  }

  const requestPath = requestUrl.pathname === "/" ? "/index.html" : requestUrl.pathname;
  const safePath = path.normalize(requestPath).replace(/^(\.\.[\\/])+/, "");
  const absolutePath = path.join(ROOT_DIR, safePath);

  if (!absolutePath.startsWith(ROOT_DIR)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  let finalPath = absolutePath;
  if (!fs.existsSync(finalPath) || fs.statSync(finalPath).isDirectory()) {
    finalPath = path.join(ROOT_DIR, "index.html");
  }

  const ext = path.extname(finalPath).toLowerCase();
  const contentType = MIME_TYPES[ext] || "application/octet-stream";
  res.writeHead(200, {
    "Content-Type": contentType,
    "Cache-Control": ext === ".html" ? "no-store" : "no-cache"
  });

  if (req.method === "HEAD") {
    res.end();
    return;
  }

  fs.createReadStream(finalPath).pipe(res);
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": "*"
  });
  res.end(JSON.stringify(data));
}

async function fetchWithTimeout(url, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), TALLY_TIMEOUT_MS);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
}

function getPythonExecutable() {
  if (fs.existsSync(BUNDLED_PYTHON)) return BUNDLED_PYTHON;
  return "python";
}

async function extractPdfText(pdfPath) {
  try {
    return await extractPdfTextWithPdfJs(pdfPath);
  } catch (error) {
    console.error("PDF.js extraction failed, trying Python extractor:", error.message);
    return extractPdfTextWithPython(pdfPath);
  }
}

async function extractPdfTextWithPdfJs(pdfPath) {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const data = new Uint8Array(await fs.promises.readFile(pdfPath));
  const loadingTask = pdfjs.getDocument({
    data,
    disableWorker: true,
    isEvalSupported: false,
    useWorkerFetch: false
  });
  const pdf = await loadingTask.promise;
  const rawPages = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    rawPages.push(buildPdfPageText(content.items || []));
  }
  const text = rawPages.join("\n");
  const normalized = text.replace(/\s+/g, " ").trim();
  return {
    text: normalized,
    rawText: text,
    normalizedText: normalized,
    rawPages,
    normalizedPages: rawPages.map((page) => page.replace(/\s+/g, " ").trim())
  };
}

function buildPdfPageText(items = []) {
  const lines = [];
  let currentLine = [];
  let currentY = null;
  sortedPdfTextItems(items).forEach((item) => {
    const text = cleanText(item?.str);
    const y = Array.isArray(item?.transform) ? Number(item.transform[5]) : null;
    const shouldBreakLine = currentLine.length
      && currentY !== null
      && y !== null
      && Math.abs(y - currentY) > 2;
    if (shouldBreakLine) {
      lines.push(currentLine.join(" ").replace(/\s+/g, " ").trim());
      currentLine = [];
    }
    if (text) currentLine.push(text);
    if (y !== null) currentY = y;
    if (item?.hasEOL && currentLine.length) {
      lines.push(currentLine.join(" ").replace(/\s+/g, " ").trim());
      currentLine = [];
      currentY = null;
    }
  });
  if (currentLine.length) {
    lines.push(currentLine.join(" ").replace(/\s+/g, " ").trim());
  }
  return lines.join("\n");
}

function sortedPdfTextItems(items = []) {
  return [...items].sort((a, b) => {
    const ay = Array.isArray(a?.transform) ? Number(a.transform[5]) : 0;
    const by = Array.isArray(b?.transform) ? Number(b.transform[5]) : 0;
    if (Math.abs(by - ay) > 2) return by - ay;
    const ax = Array.isArray(a?.transform) ? Number(a.transform[4]) : 0;
    const bx = Array.isArray(b?.transform) ? Number(b.transform[4]) : 0;
    return ax - bx;
  });
}

function cleanText(value) {
  return String(value || "").trim();
}

function extractPdfTextWithPython(pdfPath) {
  const script = `
import json
import re
import sys
from pathlib import Path

try:
    from pypdf import PdfReader
except Exception as exc:
    raise SystemExit(f"pypdf import failed: {exc}")

path = Path(sys.argv[1])
reader = PdfReader(str(path))
raw_pages = [(page.extract_text() or "") for page in reader.pages]
text = "\\n".join(raw_pages)
normalized = re.sub(r"\\s+", " ", text).strip()
normalized_pages = [re.sub(r"\\s+", " ", page).strip() for page in raw_pages]

def extract(pattern):
    match = re.search(pattern, normalized, re.IGNORECASE | re.DOTALL)
    return match.group(1).strip() if match else ""

payload = {
    "text": normalized,
    "rawText": text,
    "normalizedText": normalized,
    "rawPages": raw_pages,
    "normalizedPages": normalized_pages,
    "reportNumber": extract(r"\\b(RRLD\\d{8,})\\b"),
    "reportDate": extract(r"Received date\\s+(\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4})") or extract(r"Report No\\.\\s*:?\\s*(\\d{1,2}\\s+[A-Za-z]{3}\\s+\\d{4})"),
    "labName": "Reliance Trends Product Testing Laboratory" if "Reliance Trends PRODUCT TESTING LABORATORY".lower() in normalized.lower() else "Testing Laboratory",
    "reportStyleCode": extract(r"Style Code\\s*:?\\s*(.+?)\\s*No\\.?\\s*of Sample"),
    "reportColor": extract(r"Color\\s*:?\\s*(.+?)\\s*Style Code"),
    "reportBrand": extract(r"Brand\\s*:?\\s*(.+?)\\s*Supplier\\s*:"),
    "reportSupplier": extract(r"Supplier\\s*:?\\s*(.+?)\\s*Trf ID") or extract(r"Mill Supplier\\s*:?\\s*(.+?)\\s*Type of Testing"),
    "composition": extract(r"Fiber Content\\s*:?\\s*(.+?)\\s*Fabric Code"),
    "washcareText": extract(r"Submitted Wash Care\\s*:?\\s*(.+?)\\s*Test Name")
}
print(json.dumps(payload))
`.trim();

  return new Promise((resolve, reject) => {
    const child = spawn(getPythonExecutable(), ["-c", script, pdfPath], {
      windowsHide: true
    });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) {
        reject(new Error(cleanShellText(stderr) || `Python exited with code ${code}.`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(new Error(`Invalid extractor response: ${error.message}`));
      }
    });
  });
}

function cleanShellText(value) {
  return String(value || "").trim();
}
