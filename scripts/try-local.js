#!/usr/bin/env node
import fs from "fs/promises";
import os from "os";
import path from "path";
import { fileURLToPath } from "url";
import { spawn } from "child_process";

const MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");

async function resolveMarpBin() {
  const binName = process.platform === "win32" ? "marp.cmd" : "marp";
  const candidate = path.resolve(projectRoot, "node_modules", ".bin", binName);
  try {
    await fs.access(candidate);
    return candidate;
  } catch {
    return binName;
  }
}

async function safeRm(target) {
  try {
    await fs.rm(target, { recursive: true, force: true });
  } catch {
    // ignore
  }
}

async function generatePptxFromMarkdown(markdown) {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "mcp-pptmaker-test-"));
  const inputPath = path.join(tmpDir, "deck.md");
  const outputPath = path.join(tmpDir, "presentation.pptx");
  await fs.writeFile(inputPath, markdown, "utf8");

  const marpBin = await resolveMarpBin();
  const args = [inputPath, "--pptx", "-o", outputPath, "--quiet"];

  const child = spawn(marpBin, args, {
    cwd: tmpDir,
    stdio: ["ignore", "pipe", "pipe"],
    env: { ...process.env },
  });

  let stderr = "";
  child.stderr.on("data", (d) => {
    stderr += d.toString();
  });

  const exitCode = await new Promise((resolve) => child.on("close", resolve));

  if (exitCode !== 0) {
    await safeRm(tmpDir);
    throw new Error(`Marp CLI failed (exit ${exitCode}): ${stderr.trim()}`);
  }

  try {
    await fs.access(outputPath);
  } catch {
    await safeRm(tmpDir);
    throw new Error(`Marp CLI did not produce output at ${outputPath}`);
  }

  const data = await fs.readFile(outputPath);
  const base64 = data.toString("base64");

  await safeRm(tmpDir);

  return { filename: "presentation.pptx", mimeType: MIME_PPTX, base64 };
}

async function main() {
  const samplePath = path.resolve(projectRoot, "samples", "example.md");
  const markdown = await fs.readFile(samplePath, "utf8");
  const result = await generatePptxFromMarkdown(markdown);
  process.stdout.write(JSON.stringify(result, null, 2) + "\n");
}

main().catch((err) => {
  const msg = err instanceof Error ? err.message : String(err);
  console.error(msg);
  process.exit(1);
});