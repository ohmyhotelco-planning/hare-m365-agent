import fs from "node:fs";
import path from "node:path";
import type {
  INetworkModule,
  NetworkRequestOptions,
  NetworkResponse
} from "@azure/msal-node";
import type { AppConfig } from "./config.js";
import { ProxyAwareNetworkClient } from "./msal-network.js";

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

export async function startDeviceLogin(
  config: AppConfig,
  scopes: string[],
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<DeviceLoginStartResult> {
  requirePersistentDataDir(config);
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

export function readDeviceLoginState(config: AppConfig): DeviceLoginState {
  requirePersistentDataDir(config);
  const statePath = deviceLoginStatePath(config);
  if (!fs.existsSync(statePath)) {
    throw new Error("No pending Hare login. Run auth login-start first.");
  }

  let state: DeviceLoginState;
  try {
    state = JSON.parse(fs.readFileSync(statePath, "utf8")) as DeviceLoginState;
  } catch (error) {
    throw new Error(`Pending Hare login state is unreadable: ${errorMessage(error)}`);
  }

  if (
    state.version !== 1 ||
    state.clientId !== config.clientId ||
    state.authority !== config.authority ||
    !Array.isArray(state.scopes)
  ) {
    throw new Error("Pending Hare login state does not match the current configuration.");
  }
  validateServerResponse(state.response);
  if (Date.now() >= Date.parse(state.expiresAt)) {
    clearDeviceLoginState(config);
    throw new Error("The Microsoft device code expired. Run auth login-start again.");
  }
  return state;
}

export function clearDeviceLoginState(config: AppConfig): void {
  fs.rmSync(deviceLoginStatePath(config), { force: true });
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
    typeof value.user_code !== "string" ||
    typeof value.device_code !== "string" ||
    typeof value.verification_uri !== "string" ||
    typeof value.message !== "string" ||
    !Number.isFinite(value.expires_in) || value.expires_in <= 0 ||
    !Number.isFinite(value.interval) || value.interval <= 0
  ) {
    throw new Error("Microsoft returned an invalid device-code response.");
  }
}

function writeState(filePath: string, state: DeviceLoginState): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const temporaryPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  try {
    fs.writeFileSync(temporaryPath, JSON.stringify(state), { encoding: "utf8", mode: 0o600 });
    fs.renameSync(temporaryPath, filePath);
  } finally {
    fs.rmSync(temporaryPath, { force: true });
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
