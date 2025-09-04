#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";
import { spawn } from "child_process";
import { createReadStream } from "fs";
import { createServer } from "http";
import { randomUUID as uuidv4 } from "node:crypto";

const MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");

// Best-effort experimental in-process API from @marp-team/marp-cli
async function generateWithApi(markdown) {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "mcp-pptmaker-"));
  const inputName = "deck.md";
  const outputName = "presentation.pptx";
  const inputPath = path.join(tmpDir, inputName);
  const outputPath = path.join(tmpDir, outputName);
  await fs.writeFile(inputPath, markdown, "utf8");

  let mod;
  try {
    mod = await import("@marp-team/marp-cli");
  } catch (e) {
    await safeRm(tmpDir);
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error(`Experimental Marp API not available: ${msg}`);
  }

  const runner =
    typeof mod.default === "function"
      ? mod.default
      : typeof mod.cli === "function"
      ? mod.cli
      : null;

  if (!runner) {
    await safeRm(tmpDir);
    throw new Error("Experimental Marp API not found in @marp-team/marp-cli");
  }

  const cwdBefore = process.cwd();
  try {
    // Work in isolated temp dir to avoid picking up ambient .marprc.*
    process.chdir(tmpDir);
    const args = ["--pptx", inputName, "-o", outputName, "--quiet"];

    // Some versions return exit code (number), others may throw on error.
    await runner(args);

    // Verify output exists
    await fs.access(outputPath);
    const data = await fs.readFile(outputPath);
    const base64 = data.toString("base64");
    return { filename: outputName, base64, outputPath, tmpDir };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    await safeRm(tmpDir);
    throw new Error(`Marp API failed: ${msg}`);
  } finally {
    try {
      process.chdir(cwdBefore);
    } catch {
      // ignore
    }
  }
}

async function resolveMarpBin() {
  const binName = process.platform === "win32" ? "marp.cmd" : "marp";
  const candidate = path.resolve(projectRoot, "node_modules", ".bin", binName);
  try {
    await fs.access(candidate);
    return candidate;
  } catch {
    // Fallback to PATH
    return binName;
  }
}

async function generateWithCli(markdown) {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "mcp-pptmaker-"));
  const inputPath = path.join(tmpDir, "deck.md");
  const outputPath = path.join(tmpDir, "presentation.pptx");
  await fs.writeFile(inputPath, markdown, "utf8");

  const marpBin = await resolveMarpBin();
  const args = [inputPath, "--pptx", "-o", outputPath, "--quiet"];

  const child = spawn(marpBin, args, {
    cwd: tmpDir,
    stdio: ["ignore", "pipe", "pipe"],
    env: {
      ...process.env,
    },
  });

  let stderr = "";
  child.stderr.on("data", (d) => {
    stderr += d.toString();
  });

  const exitCode = await new Promise((resolve) => {
    child.on("close", resolve);
  });

  if (exitCode !== 0) {
    await safeRm(tmpDir);
    throw new Error(`Marp CLI failed (exit ${exitCode}): ${stderr.trim()}`);
  }

  // Ensure output exists
  try {
    await fs.access(outputPath);
  } catch {
    await safeRm(tmpDir);
    throw new Error(`Marp CLI did not produce output at ${outputPath}`);
  }

  const data = await fs.readFile(outputPath);
  const base64 = data.toString("base64");
  return { filename: "presentation.pptx", base64, outputPath, tmpDir };
}

async function safeRm(target) {
  try {
    await fs.rm(target, { recursive: true, force: true });
  } catch {
    // ignore
  }
}

// Lightweight HTTP file server for temporary downloads
const FILE_HOST = process.env.MCP_PPT_HOST || "127.0.0.1";
const FILE_PORT = process.env.MCP_PPT_PORT ? parseInt(process.env.MCP_PPT_PORT, 10) : 0; // 0 = ephemeral port
const FILE_TTL_MS =
  (process.env.MCP_PPT_TTL_SEC ? parseInt(process.env.MCP_PPT_TTL_SEC, 10) : 1800) * 1000; // default 30min
const FILE_BASE_DIR = path.join(os.tmpdir(), "mcp-pptmaker-files");

let fileServerPromise = null;
let fileServerPort = null;
// id -> { path, filename, mimeType, expiresAt }
const filesIndex = new Map();

function getBaseUrl() {
  const host = FILE_HOST.includes(":") ? `[${FILE_HOST}]` : FILE_HOST;
  return `http://${host}:${fileServerPort}`;
}

async function cleanExpiredFiles() {
  const now = Date.now();
  for (const [id, meta] of Array.from(filesIndex.entries())) {
    if (now > meta.expiresAt) {
      filesIndex.delete(id);
      try {
        await fs.rm(path.dirname(meta.path), { recursive: true, force: true });
      } catch {
        // ignore
      }
    }
  }
}

async function ensureFileServer() {
  if (fileServerPromise) return fileServerPromise;
  await fs.mkdir(FILE_BASE_DIR, { recursive: true });

  fileServerPromise = new Promise((resolve, reject) => {
    const srv = createServer(async (req, res) => {
      try {
        const url = new URL(req.url, "http://localhost");
        if (req.method !== "GET") {
          res.statusCode = 405;
          res.end("Method Not Allowed");
          return;
        }

        const parts = url.pathname.split("/").filter(Boolean); // e.g., ["files", "{id}", "{name}"]
        if (parts.length >= 2 && parts[0] === "files") {
          const id = parts[1];
          const meta = filesIndex.get(id);
          if (!meta) {
            res.statusCode = 404;
            res.end("Not found");
            return;
          }
          if (Date.now() > meta.expiresAt) {
            filesIndex.delete(id);
            try {
              await fs.rm(path.dirname(meta.path), { recursive: true, force: true });
            } catch {}
            res.statusCode = 410;
            res.end("Gone");
            return;
          }

          res.setHeader("Content-Type", meta.mimeType);
          res.setHeader(
            "Content-Disposition",
            `attachment; filename="${encodeURIComponent(meta.filename)}"`
          );
          res.setHeader("X-Expires-At", new Date(meta.expiresAt).toISOString());
          res.setHeader("Cache-Control", "public, max-age=600");
          res.setHeader("X-Content-Type-Options", "nosniff");
          res.setHeader("Access-Control-Allow-Origin", "*");

          const stream = createReadStream(meta.path);
          stream.on("error", () => {
            res.statusCode = 404;
            res.end("Not found");
          });
          stream.pipe(res);
          return;
        }

        res.statusCode = 404;
        res.end("Not found");
      } catch {
        res.statusCode = 500;
        res.end("Internal error");
      }
    });

    srv.listen(FILE_PORT, FILE_HOST, () => {
      const addr = srv.address();
      fileServerPort = typeof addr === "object" && addr ? addr.port : FILE_PORT;
      // Periodic TTL cleanup
      setInterval(cleanExpiredFiles, 60 * 1000).unref();
      resolve(srv);
    });

    srv.on("error", reject);
  });

  return fileServerPromise;
}

async function publishAndLinkFile(tempPath, filename, mimeType) {
  await ensureFileServer();
  const id = typeof uuidv4 === "function" ? uuidv4() : Math.random().toString(36).slice(2);
  const dir = path.join(FILE_BASE_DIR, id);
  await fs.mkdir(dir, { recursive: true });
  const dstPath = path.join(dir, filename);
  await fs.copyFile(tempPath, dstPath);

  const expiresAt = Date.now() + FILE_TTL_MS;
  filesIndex.set(id, { path: dstPath, filename, mimeType, expiresAt });

  const link = `${getBaseUrl()}/files/${id}/${encodeURIComponent(filename)}`;
  return { link, expiresAt, id };
}

// Create MCP Server
const server = new McpServer({
  name: "mcp-pptmaker",
  version: "0.1.0",
});

server.tool(
  "generate_pptx",
  {
    markdown: z
      .string()
      .min(1, "markdown cannot be empty")
      .describe("Marp Markdown content with optional front-matter controlling theme"),
  },
  async ({ markdown }) => {
    try {
      // Try experimental in-process API first
      let result;
      try {
        result = await generateWithApi(markdown);
      } catch (apiErr) {
        // Fallback to CLI
        result = await generateWithCli(markdown);
      }

      const { link, expiresAt } = await publishAndLinkFile(
        result.outputPath,
        result.filename,
        MIME_PPTX
      );

      // Clean up marp temp dir
      try { await safeRm(result.tmpDir); } catch {}

      return {
        content: [
          {
            type: "text",
            text:
              `Download PPTX: ${link}\n` +
              `Expires: ${new Date(expiresAt).toISOString()}\n` +
              `Filename: ${result.filename}\n` +
              `MimeType: ${MIME_PPTX}`
          }
        ],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text", text: message }],
        isError: true,
      };
    }
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("mcp-pptmaker server running (stdio) - tool: generate_pptx");