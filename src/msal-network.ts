import type {
  INetworkModule,
  NetworkRequestOptions,
  NetworkResponse
} from "@azure/msal-node";
import { fetchWithProxy } from "./proxy.js";

export type AuthNetworkFailure = {
  method: "GET" | "POST";
  hostname: string;
  stage: "request" | "response";
  code: string;
};

const networkCodes = new Set([
  "EACCES", "EPERM", "ENOTFOUND", "EAI_AGAIN", "ETIMEDOUT",
  "ECONNRESET", "ECONNREFUSED", "ENETUNREACH", "EHOSTUNREACH",
  "UND_ERR_CONNECT_TIMEOUT", "UND_ERR_HEADERS_TIMEOUT", "UND_ERR_BODY_TIMEOUT",
  "UND_ERR_SOCKET", "CERT_HAS_EXPIRED", "UNABLE_TO_VERIFY_LEAF_SIGNATURE",
  "DEPTH_ZERO_SELF_SIGNED_CERT", "SELF_SIGNED_CERT_IN_CHAIN"
]);

function networkErrorCode(error: unknown, timedOut: boolean): string {
  if (timedOut) return "ETIMEDOUT";
  const pending: unknown[] = [error];
  const seen = new Set<unknown>();
  while (pending.length && seen.size < 16) {
    const current = pending.shift();
    if (!current || typeof current !== "object" || seen.has(current)) continue;
    seen.add(current);
    const value = current as { code?: unknown; cause?: unknown; errors?: unknown };
    if (typeof value.code === "string" && networkCodes.has(value.code)) return value.code;
    pending.push(value.cause);
    if (Array.isArray(value.errors)) pending.push(...value.errors.slice(0, 16));
  }
  return "NETWORK_ERROR";
}

export class ProxyAwareNetworkClient implements INetworkModule {
  private networkFailure?: AuthNetworkFailure;

  constructor(private readonly fetcher: typeof fetchWithProxy = fetchWithProxy) {}

  // MSAL can replace the thrown error; retain only allowlisted, operation-local diagnostics.
  takeNetworkFailure(): AuthNetworkFailure | undefined {
    const failure = this.networkFailure;
    this.networkFailure = undefined;
    return failure;
  }

  async sendGetRequestAsync<T>(
    url: string,
    options?: NetworkRequestOptions,
    timeout?: number
  ): Promise<NetworkResponse<T>> {
    return this.send<T>(url, "GET", options, timeout);
  }

  async sendGetHeadersAsync(url: string, timeout: number): Promise<NetworkResponse<Record<string, never>>> {
    return this.send(url, "GET", undefined, timeout, true);
  }

  async sendPostRequestAsync<T>(
    url: string,
    options?: NetworkRequestOptions
  ): Promise<NetworkResponse<T>> {
    return this.send<T>(url, "POST", options, 30_000);
  }

  private async send<T>(
    url: string,
    method: "GET" | "POST",
    options?: NetworkRequestOptions,
    timeout?: number,
    headersOnly = false
  ): Promise<NetworkResponse<T>> {
    const effectiveTimeout = timeout && timeout > 0 ? timeout : 30_000;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), effectiveTimeout);
    this.networkFailure = undefined;
    let stage: AuthNetworkFailure["stage"] = "request";

    try {
      const response = await this.fetcher(url, {
        method,
        headers: options?.headers,
        body: method === "POST" ? options?.body : undefined,
        signal: controller.signal
      });
      stage = "response";
      if (headersOnly) {
        // A diagnostic needs headers only; body errors must not hide an allowlist denial.
        void response.body?.cancel().catch(() => undefined);
        return { headers: Object.fromEntries(response.headers.entries()), body: {} as T, status: response.status };
      }
      const text = await response.text();

      return {
        headers: Object.fromEntries(response.headers.entries()),
        body: parseResponseBody<T>(text),
        status: response.status
      };
    } catch (error) {
      this.networkFailure = {
        method,
        hostname: new URL(url).hostname,
        stage,
        code: networkErrorCode(error, controller.signal.aborted)
      };
      throw error;
    } finally {
      clearTimeout(timer);
    }
  }
}

function parseResponseBody<T>(text: string): T {
  if (!text) return {} as T;

  try {
    return JSON.parse(text) as T;
  } catch {
    return text as T;
  }
}
