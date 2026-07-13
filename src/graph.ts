import { getAccessToken } from "./auth.js";
import type { AppConfig } from "./config.js";
import { fetchWithProxy } from "./proxy.js";

const graphRoot = "https://graph.microsoft.com/v1.0";
const retryableStatuses = new Set([429, 500, 502, 503, 504]);
const maxAttempts = 4;
const requestTimeoutMs = 30_000;

export type GraphPage<T> = {
  value?: T[];
  "@odata.nextLink"?: string;
};

export async function graphGet<T>(config: AppConfig, pathOrUrl: string): Promise<T> {
  const response = await graphRequest(config, pathOrUrl, "GET");
  return (await response.json()) as T;
}

export async function graphPost<T>(config: AppConfig, pathOrUrl: string, body: unknown): Promise<T> {
  const response = await graphRequest(config, pathOrUrl, "POST", JSON.stringify(body));
  return (await response.json()) as T;
}

export async function graphDownloadResponse(config: AppConfig, pathOrUrl: string) {
  return graphRequest(config, pathOrUrl, "GET", undefined, false);
}

async function graphRequest(
  config: AppConfig,
  pathOrUrl: string,
  method: "GET" | "POST",
  body?: string,
  acceptJson = true
) {
  const token = await getAccessToken(config);
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${graphRoot}${pathOrUrl}`;
  let lastError: unknown;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), requestTimeoutMs);
    try {
      const response = await fetchWithProxy(url, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          ...(acceptJson ? { Accept: "application/json" } : {}),
          ...(body ? { "Content-Type": "application/json" } : {})
        },
        body,
        signal: controller.signal
      });

      if (response.ok) return response;

      if (retryableStatuses.has(response.status) && attempt < maxAttempts) {
        await response.body?.cancel();
        await sleep(retryDelayMs(response.headers.get("retry-after"), attempt));
        continue;
      }

      const responseBody = await response.text();
      throw new Error(
        `Graph ${method} failed (${response.status} ${response.statusText}): ${responseBody}`
      );
    } catch (error) {
      lastError = error;
      if (!isTransientNetworkError(error) || attempt >= maxAttempts) throw error;
      await sleep(retryDelayMs(undefined, attempt));
    } finally {
      clearTimeout(timer);
    }
  }

  throw lastError instanceof Error ? lastError : new Error("Graph request failed.");
}

export function retryDelayMs(retryAfter: string | null | undefined, attempt: number): number {
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds >= 0) return Math.min(seconds * 1000, 60_000);
    const dateDelay = Date.parse(retryAfter) - Date.now();
    if (Number.isFinite(dateDelay) && dateDelay > 0) return Math.min(dateDelay, 60_000);
  }
  return Math.min(500 * 2 ** (attempt - 1), 8_000);
}

function isTransientNetworkError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  if (error.name === "AbortError") return true;
  return /fetch failed|network|socket|ECONNRESET|ETIMEDOUT|EAI_AGAIN/i.test(
    `${error.message} ${(error as Error & { cause?: unknown }).cause ?? ""}`
  );
}

function sleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

export function encodeQuery(value: string): string {
  return encodeURIComponent(value).replace(/'/g, "%27");
}
