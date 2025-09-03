#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";
import { spawn } from "child_process";

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
    await safeRm(tmpDir);
    return { filename: outputName, base64 };
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

  await safeRm(tmpDir);

  return { filename: "presentation.pptx", base64 };
}

async function safeRm(target) {
  try {
    await fs.rm(target, { recursive: true, force: true });
  } catch {
    // ignore
  }
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

      const payload = {
        filename: result.filename,
        mimeType: MIME_PPTX,
        base64: result.base64,
      };
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(payload, null, 2),
          },
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