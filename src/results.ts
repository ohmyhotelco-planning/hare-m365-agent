import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";

export function resolveResultPath(config: AppConfig, outputPath: string): string {
  if (path.isAbsolute(outputPath)) return path.resolve(outputPath);

  const resultsRoot = path.resolve(config.resultsDir);
  const resolvedPath = path.resolve(resultsRoot, outputPath);
  if (resolvedPath !== resultsRoot && !resolvedPath.startsWith(`${resultsRoot}${path.sep}`)) {
    throw new Error("Relative --out paths must stay inside Hare resultsDir.");
  }

  return resolvedPath;
}

export function cleanupExpiredResults(config: AppConfig, now = Date.now()): number {
  const retentionDays = config.policy.retentionDays;
  if (!Number.isFinite(retentionDays) || retentionDays <= 0 || !fs.existsSync(config.resultsDir)) {
    return 0;
  }

  const cutoff = now - retentionDays * 24 * 60 * 60 * 1000;
  let removed = 0;

  function cleanDirectory(directory: string): void {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const entryPath = path.join(directory, entry.name);
      if (entry.isSymbolicLink()) continue;

      if (entry.isDirectory()) {
        cleanDirectory(entryPath);
        if (fs.readdirSync(entryPath).length === 0) fs.rmdirSync(entryPath);
        continue;
      }

      if (entry.isFile() && fs.statSync(entryPath).mtimeMs < cutoff) {
        fs.rmSync(entryPath, { force: true });
        removed += 1;
      }
    }
  }

  cleanDirectory(config.resultsDir);
  return removed;
}
