import type {
  INetworkModule,
  NetworkRequestOptions,
  NetworkResponse
} from "@azure/msal-node";
import { fetchWithProxy } from "./proxy.js";

export class ProxyAwareNetworkClient implements INetworkModule {
  async sendGetRequestAsync<T>(
    url: string,
    options?: NetworkRequestOptions,
    timeout?: number
  ): Promise<NetworkResponse<T>> {
    return this.send<T>(url, "GET", options, timeout);
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
    timeout?: number
  ): Promise<NetworkResponse<T>> {
    const effectiveTimeout = timeout && timeout > 0 ? timeout : 30_000;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), effectiveTimeout);

    try {
      const response = await fetchWithProxy(url, {
        method,
        headers: options?.headers,
        body: method === "POST" ? options?.body : undefined,
        signal: controller.signal
      });
      const text = await response.text();

      return {
        headers: Object.fromEntries(response.headers.entries()),
        body: parseResponseBody<T>(text),
        status: response.status
      };
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
