import fs from "node:fs";
import path from "node:path";
import { Readable, Transform } from "node:stream";
import { pipeline } from "node:stream/promises";
import type { AppConfig } from "./config.js";
import { clearStoredFile } from "./persistent-storage.js";

export type DownloadResponse = {
  headers: { get(name: string): string | null };
  body: { cancel(): Promise<void> } | null;
};

export async function saveDownloadResponse(
  config: AppConfig,
  response: DownloadResponse,
  filename: string,
  maxBytes = config.policy.maxDownloadBytes
): Promise<string> {
  const contentLength = Number(response.headers.get("content-length"));
  if (Number.isFinite(contentLength) && contentLength > maxBytes) {
    await response.body?.cancel();
    throw new Error(`Download exceeds policy limit of ${maxBytes} bytes.`);
  }
  if (!response.body) throw new Error("Download response contained no file body.");

  fs.mkdirSync(config.downloadDir, { recursive: true });
  const outputPath = uniqueOutputPath(config.downloadDir, sanitizeFilename(filename));
  let receivedBytes = 0;
  const limiter = new Transform({
    transform(chunk: Buffer, _encoding, callback) {
      receivedBytes += chunk.length;
      if (receivedBytes > maxBytes) {
        callback(new Error(`Download exceeds policy limit of ${maxBytes} bytes.`));
        return;
      }
      callback(null, chunk);
    }
  });

  try {
    await pipeline(
      Readable.fromWeb(response.body as never),
      limiter,
      fs.createWriteStream(outputPath, { flags: "wx", mode: 0o600 })
    );
  } catch (error) {
    clearStoredFile(outputPath);
    throw error;
  }
  return outputPath;
}

export function sanitizeFilename(value: string): string {
  const sanitized = value
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/[. ]+$/g, "")
    .trim()
    .slice(0, 180);
  const fallback = sanitized || "download.bin";
  const baseName = path.parse(fallback).name;
  return /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(baseName) ? `_${fallback}` : fallback;
}

function uniqueOutputPath(directory: string, filename: string): string {
  const parsed = path.parse(filename);
  for (let index = 0; index < 10_000; index += 1) {
    const suffix = index === 0 ? "" : ` (${index})`;
    const candidate = path.join(directory, `${parsed.name}${suffix}${parsed.ext}`);
    if (!fs.existsSync(candidate)) return candidate;
  }
  throw new Error("Could not allocate a unique download filename.");
}
