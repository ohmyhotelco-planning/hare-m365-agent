import fs from "node:fs";
import path from "node:path";
import type {
  INetworkModule,
  NetworkRequestOptions,
  NetworkResponse
} from "@azure/msal-node";
import type { AppConfig } from "./config.js";
import { prepareAuthProfile } from "./auth-profile.js";
import { ProxyAwareNetworkClient } from "./msal-network.js";
import { clearStoredFile, writeStoredText } from "./persistent-storage.js";

type ServerDeviceCodeResponse = {
  user_code: string;
  device_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
};

export type DeviceLoginState = {
  version: 1;
  clientId: string;
  authority: string;
  scopes: string[];
  createdAt: string;
  expiresAt: string;
  response: ServerDeviceCodeResponse;
};

export type DeviceLoginStartResult = {
  ok: true;
  stage: "WAITING_FOR_USER";
  verificationUri: string;
  userCode: string;
  message: string;
  expiresAt: string;
};

export function deviceLoginStatePath(config: AppConfig): string {
  return path.join(config.cacheDir, "device-login-state.json");
}

export function hasPendingDeviceLoginState(config: AppConfig, scopes?: string[]): boolean {
  // Status inspection must not retire or rewrite a pending login.
  try {
    const state = readValidatedDeviceLoginState(config, scopes);
    return Date.now() < Date.parse(state.expiresAt);
  } catch {
    return false;
  }
}

export async function startDeviceLogin(
  config: AppConfig,
  scopes: string[],
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<DeviceLoginStartResult> {
  requirePersistentDataDir(config);
  const profile = prepareAuthProfile(config, scopes);
  if (profile.reason) throw new Error(profile.reason);
  const endpoint = `${config.authority}/oauth2/v2.0/devicecode`;
  const body = new URLSearchParams({
    client_id: config.clientId,
    scope: scopes.join(" ")
  }).toString();
  const networkResponse = await networkClient.sendPostRequestAsync<ServerDeviceCodeResponse>(endpoint, {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });

  if (networkResponse.status < 200 || networkResponse.status >= 300) {
    throw new Error(`Microsoft device-code request failed with HTTP ${networkResponse.status}.`);
  }
  validateServerResponse(networkResponse.body);

  const createdAt = new Date();
  const state: DeviceLoginState = {
    version: 1,
    clientId: config.clientId,
    authority: config.authority,
    scopes: [...scopes],
    createdAt: createdAt.toISOString(),
    expiresAt: new Date(createdAt.getTime() + networkResponse.body.expires_in * 1000).toISOString(),
    response: networkResponse.body
  };
  writeState(deviceLoginStatePath(config), state);

  return {
    ok: true,
    stage: "WAITING_FOR_USER",
    verificationUri: state.response.verification_uri,
    userCode: state.response.user_code,
    message: state.response.message,
    expiresAt: state.expiresAt
  };
}

export function readDeviceLoginState(config: AppConfig, scopes?: string[]): DeviceLoginState {
  requirePersistentDataDir(config);
  const state = readValidatedDeviceLoginState(config, scopes);
  if (Date.now() >= Date.parse(state.expiresAt)) {
    clearDeviceLoginState(config);
    throw new Error("The Microsoft device code expired. Run auth login-start again.");
  }
  return state;
}

function readValidatedDeviceLoginState(config: AppConfig, scopes?: string[]): DeviceLoginState {
  let text: string;
  try {
    text = fs.readFileSync(deviceLoginStatePath(config), "utf8");
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      throw new Error("LOGIN_START_REQUIRED: No pending Hare login. Run auth login-start first.");
    }
    throw new Error("Pending Hare login state is unreadable.");
  }
  if (!text.trim()) {
    throw new Error("LOGIN_START_REQUIRED: No pending Hare login. Run auth login-start first.");
  }

  let state: DeviceLoginState;
  try {
    state = JSON.parse(text) as DeviceLoginState;
  } catch {
    throw new Error("Pending Hare login state is unreadable. Run auth login-start again.");
  }

  if (
    !state ||
    state.version !== 1 ||
    state.clientId !== config.clientId ||
    state.authority !== config.authority ||
    !Array.isArray(state.scopes) ||
    state.scopes.length === 0 ||
    !state.scopes.every((scope) => typeof scope === "string" && scope.trim().length > 0)
  ) {
    throw new Error("Pending Hare login state does not match the current configuration.");
  }
  validateServerResponse(state.response);
  if (
    typeof state.createdAt !== "string" || !Number.isFinite(Date.parse(state.createdAt)) ||
    typeof state.expiresAt !== "string" || !Number.isFinite(Date.parse(state.expiresAt)) ||
    Date.parse(state.expiresAt) <= Date.parse(state.createdAt)
  ) {
    throw new Error("Pending Hare login state has invalid timestamps. Run auth login-start again.");
  }
  if (scopes) {
    const expected = new Set(scopes);
    const pending = new Set(state.scopes);
    if (expected.size !== pending.size || ![...expected].every((scope) => pending.has(scope))) {
      throw new Error("Pending login scopes changed. Run auth login-start again.");
    }
  }
  return state;
}

export function clearDeviceLoginState(config: AppConfig): void {
  clearStoredFile(deviceLoginStatePath(config));
}

export class ResumeDeviceCodeNetworkClient implements INetworkModule {
  private servedDeviceCode = false;

  constructor(
    private readonly state: DeviceLoginState,
    private readonly delegate: INetworkModule = new ProxyAwareNetworkClient()
  ) {}

  sendGetRequestAsync<T>(
    url: string,
    options?: NetworkRequestOptions,
    timeout?: number
  ): Promise<NetworkResponse<T>> {
    return this.delegate.sendGetRequestAsync<T>(url, options, timeout);
  }

  sendPostRequestAsync<T>(
    url: string,
    options?: NetworkRequestOptions
  ): Promise<NetworkResponse<T>> {
    if (!this.servedDeviceCode && /\/oauth2\/v2\.0\/devicecode(?:\?|$)/i.test(url)) {
      this.servedDeviceCode = true;
      return Promise.resolve({
        status: 200,
        headers: { "content-type": "application/json" },
        body: this.state.response as T
      });
    }
    return this.delegate.sendPostRequestAsync<T>(url, options);
  }
}

export function requirePersistentDataDir(config: AppConfig): void {
  if (config.dataDirPersistent) return;
  throw new Error(
    "FOLDER_REQUIRED: Start Cowork with the user's existing Hare project selected and rerun startup with that selected project root as the persistent store before login."
  );
}

function validateServerResponse(value: ServerDeviceCodeResponse): void {
  if (
    !value ||
    typeof value.user_code !== "string" || !value.user_code.trim() ||
    typeof value.device_code !== "string" || !value.device_code.trim() ||
    typeof value.verification_uri !== "string" || !value.verification_uri.trim() ||
    typeof value.message !== "string" || !value.message.trim() ||
    !Number.isFinite(value.expires_in) || value.expires_in <= 0 ||
    !Number.isFinite(value.interval) || value.interval <= 0
  ) {
    throw new Error("Microsoft returned an invalid device-code response.");
  }
}

function writeState(filePath: string, state: DeviceLoginState): void {
  writeStoredText(filePath, JSON.stringify(state));
}
