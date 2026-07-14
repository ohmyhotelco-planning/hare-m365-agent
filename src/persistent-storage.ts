import fs from "node:fs";
import path from "node:path";
import { isCoworkHostMountPath } from "./config.js";

export function usesDeleteRestrictedStorage(filePath: string): boolean {
  return isCoworkHostMountPath(path.resolve(filePath));
}

type StorageOptions = {
  deleteRestricted?: boolean;
};

export function writeStoredText(
  filePath: string,
  contents: string,
  options: StorageOptions = {}
): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });

  if (options.deleteRestricted ?? usesDeleteRestrictedStorage(filePath)) {
    fs.writeFileSync(filePath, contents, { encoding: "utf8", mode: 0o600 });
    return;
  }

  const temporaryPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  try {
    fs.writeFileSync(temporaryPath, contents, { encoding: "utf8", mode: 0o600 });
    try {
      fs.renameSync(temporaryPath, filePath);
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code;
      if ((code === "EPERM" || code === "EACCES" || code === "EEXIST") && fs.existsSync(filePath)) {
        fs.writeFileSync(filePath, contents, { encoding: "utf8", mode: 0o600 });
      } else {
        throw error;
      }
    }
  } finally {
    fs.rmSync(temporaryPath, { force: true });
  }
}

export function clearStoredFile(
  filePath: string,
  options: StorageOptions = {}
): void {
  if (!fs.existsSync(filePath)) return;

  if (options.deleteRestricted ?? usesDeleteRestrictedStorage(filePath)) {
    fs.writeFileSync(filePath, "", { encoding: "utf8", mode: 0o600 });
    return;
  }

  fs.rmSync(filePath, { force: true });
}

export function storedFileHasContent(filePath: string): boolean {
  return fs.existsSync(filePath) && fs.statSync(filePath).size > 0;
}
