import fs from "node:fs";
import path from "node:path";
import { Readable, Transform } from "node:stream";
import { pipeline } from "node:stream/promises";
import type { AppConfig } from "./config.js";
import { clearStoredFile, usesDeleteRestrictedStorage } from "./persistent-storage.js";

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

export type ExactDownloadOptions = {
  maxBytes: number;
  expectedBytes?: number;
  startOffset?: number;
  temporarySuffix?: string;
  forceCopyFinalization?: boolean;
};

export async function saveDownloadResponseToPath(
  response: DownloadResponse & { status?: number },
  outputPath: string,
  options: ExactDownloadOptions
): Promise<number> {
  const requestedOffset = Math.max(0, Math.floor(options.startOffset ?? 0));
  const contentLength = Number(response.headers.get("content-length"));
  let acceptedOffset = 0;
  if (response.status === 206) {
    const contentRange = response.headers.get("content-range");
    const match = /^bytes (\d+)-(\d+)\/(\d+|\*)$/i.exec(contentRange ?? "");
    const start = Number(match?.[1]);
    const end = Number(match?.[2]);
    const total = match?.[3] === "*" ? Number.NaN : Number(match?.[3]);
    const valid =
      requestedOffset > 0 &&
      Number.isFinite(start) &&
      start === requestedOffset &&
      Number.isFinite(end) &&
      end >= start &&
      (!Number.isFinite(contentLength) || contentLength === end - start + 1) &&
      (!Number.isFinite(options.expectedBytes) || total === options.expectedBytes);
    if (!valid) {
      await response.body?.cancel();
      throw new Error("Partial download returned an invalid Content-Range.");
    }
    acceptedOffset = requestedOffset;
  }
  if (
    Number.isFinite(contentLength) &&
    acceptedOffset + contentLength > options.maxBytes
  ) {
    await response.body?.cancel();
    throw new Error(`Download exceeds policy limit of ${options.maxBytes} bytes.`);
  }
  if (!response.body) throw new Error("Download response contained no file body.");

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  if (fs.existsSync(outputPath)) {
    await response.body.cancel();
    throw new Error("Destination file already exists.");
  }

  const suffix = options.temporarySuffix ?? ".hare-part";
  const temporaryPath = `${outputPath}${suffix}`;
  if (acceptedOffset > 0) {
    const partialSize = fs.existsSync(temporaryPath) ? fs.statSync(temporaryPath).size : 0;
    if (partialSize !== acceptedOffset) {
      await response.body.cancel();
      throw new Error("Partial download size changed before resume.");
    }
  }

  let receivedBytes = acceptedOffset;
  const limiter = new Transform({
    transform(chunk: Buffer, _encoding, callback) {
      receivedBytes += chunk.length;
      if (receivedBytes > options.maxBytes) {
        callback(new Error(`Download exceeds policy limit of ${options.maxBytes} bytes.`));
        return;
      }
      callback(null, chunk);
    }
  });

  await pipeline(
    Readable.fromWeb(response.body as never),
    limiter,
    fs.createWriteStream(temporaryPath, {
      flags: acceptedOffset > 0 ? "a" : "w",
      mode: 0o600
    })
  );
  if (Number.isFinite(options.expectedBytes) && options.expectedBytes !== receivedBytes) {
    throw new Error(
      `Downloaded byte count ${receivedBytes} did not match metadata size ${options.expectedBytes}.`
    );
  }
  finalizeTemporaryDownload(temporaryPath, outputPath, {
    deleteRestricted: options.forceCopyFinalization
  });
  return receivedBytes;
}

export type FinalizeTemporaryDownloadOptions = {
  deleteRestricted?: boolean;
  availableDiskBytes?: () => number;
};

export function finalizeTemporaryDownload(
  temporaryPath: string,
  outputPath: string,
  options: FinalizeTemporaryDownloadOptions = {}
): void {
  const deleteRestricted =
    options.deleteRestricted ?? usesDeleteRestrictedStorage(outputPath);

  if (!deleteRestricted) {
    try {
      fs.linkSync(temporaryPath, outputPath);
      try {
        fs.rmSync(temporaryPath, { force: true });
      } catch {
        // The final hard link is complete; a stale temporary link is harmless.
      }
      return;
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code;
      if (code === "EEXIST") {
        throw new Error("Destination file already exists.");
      }
      if (!["EPERM", "EACCES", "ENOSYS", "ENOTSUP", "EOPNOTSUPP", "EXDEV", "EINVAL"].includes(code ?? "")) {
        throw error;
      }
    }
  }

  const temporarySize = fs.statSync(temporaryPath).size;
  const availableBytes =
    options.availableDiskBytes?.() ?? availableBytesForPath(outputPath);
  if (temporarySize > availableBytes) {
    throw new Error(
      "Insufficient destination disk space to finalize this file without overwriting."
    );
  }

  try {
    fs.copyFileSync(temporaryPath, outputPath, fs.constants.COPYFILE_EXCL);
    clearStoredFile(temporaryPath, { deleteRestricted });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST") {
      throw new Error("Destination file already exists.");
    }
    throw error;
  }
}

function availableBytesForPath(outputPath: string): number {
  const stat = fs.statfsSync(path.dirname(outputPath), { bigint: true });
  const available = stat.bavail * stat.bsize;
  return available > BigInt(Number.MAX_SAFE_INTEGER)
    ? Number.MAX_SAFE_INTEGER
    : Number(available);
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
