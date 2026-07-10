import {
  EnvHttpProxyAgent,
  fetch as undiciFetch,
  type Dispatcher,
  type RequestInit
} from "undici";

let proxyDispatcher: Dispatcher | undefined;

export function getHttpsProxyUrl(): string | undefined {
  const value =
    process.env.HTTPS_PROXY ??
    process.env.https_proxy ??
    process.env.HTTP_PROXY ??
    process.env.http_proxy;

  return value?.trim() || undefined;
}

function getProxyDispatcher(): Dispatcher | undefined {
  if (!getHttpsProxyUrl()) return undefined;
  proxyDispatcher ??= new EnvHttpProxyAgent();
  return proxyDispatcher;
}

export function fetchWithProxy(input: string | URL, init?: RequestInit) {
  const dispatcher = getProxyDispatcher();
  return undiciFetch(input, dispatcher ? { ...init, dispatcher } : init);
}
