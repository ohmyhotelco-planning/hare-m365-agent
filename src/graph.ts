import { getAccessToken } from "./auth.js";
import type { AppConfig } from "./config.js";

const graphRoot = "https://graph.microsoft.com/v1.0";

export type GraphPage<T> = {
  value?: T[];
  "@odata.nextLink"?: string;
};

export async function graphGet<T>(config: AppConfig, pathOrUrl: string): Promise<T> {
  const token = await getAccessToken(config);
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${graphRoot}${pathOrUrl}`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph GET failed (${response.status} ${response.statusText}): ${body}`);
  }

  return (await response.json()) as T;
}

export async function graphPost<T>(config: AppConfig, pathOrUrl: string, body: unknown): Promise<T> {
  const token = await getAccessToken(config);
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${graphRoot}${pathOrUrl}`;
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const responseBody = await response.text();
    throw new Error(`Graph POST failed (${response.status} ${response.statusText}): ${responseBody}`);
  }

  return (await response.json()) as T;
}

export async function graphDownload(config: AppConfig, pathOrUrl: string): Promise<Buffer> {
  const token = await getAccessToken(config);
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${graphRoot}${pathOrUrl}`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph download failed (${response.status} ${response.statusText}): ${body}`);
  }

  return Buffer.from(await response.arrayBuffer());
}

export function encodeQuery(value: string): string {
  return encodeURIComponent(value).replace(/'/g, "%27");
}
