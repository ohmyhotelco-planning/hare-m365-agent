import fs from "node:fs";
import path from "node:path";
import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration
} from "@azure/msal-node";
import type { AppConfig } from "./config.js";

const scopes = [
  "User.Read",
  "Mail.Read",
  "Chat.Read",
  "Chat.ReadWrite",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "Files.Read.All",
  "Sites.Read.All",
  "offline_access"
];

function cachePath(config: AppConfig): string {
  return path.join(config.cacheDir, "msal-cache.json");
}

async function buildPca(config: AppConfig): Promise<PublicClientApplication> {
  const msalConfig: Configuration = {
    auth: {
      clientId: config.clientId,
      authority: config.authority
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (cacheContext) => {
          const file = cachePath(config);
          if (fs.existsSync(file)) {
            cacheContext.tokenCache.deserialize(fs.readFileSync(file, "utf8"));
          }
        },
        afterCacheAccess: async (cacheContext) => {
          if (cacheContext.cacheHasChanged) {
            fs.mkdirSync(config.cacheDir, { recursive: true });
            fs.writeFileSync(cachePath(config), cacheContext.tokenCache.serialize(), {
              encoding: "utf8",
              mode: 0o600
            });
          }
        }
      }
    }
  };

  return new PublicClientApplication(msalConfig);
}

export async function login(config: AppConfig): Promise<AuthenticationResult> {
  const pca = await buildPca(config);
  const result = await pca.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => {
      console.log(response.message);
    }
  });

  if (!result) throw new Error("Login failed: Microsoft returned no authentication result.");
  return result;
}

export async function getAccount(config: AppConfig): Promise<AccountInfo | null> {
  const pca = await buildPca(config);
  const accounts = await pca.getTokenCache().getAllAccounts();
  return accounts[0] ?? null;
}

export async function getAccessToken(config: AppConfig): Promise<string> {
  const pca = await buildPca(config);
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0] ?? null;
  if (!account) {
    throw new Error(
      "Not logged in. Run: hare-m365 auth login. If using npx, run: npx @ohmyhotel/hare-m365-agent auth login. After login, rerun doctor/auth status and retry the original read command."
    );
  }

  const result = await pca.acquireTokenSilent({
    account,
    scopes
  });

  if (!result?.accessToken) throw new Error("Could not acquire access token.");
  return result.accessToken;
}

export function logout(config: AppConfig): void {
  const file = cachePath(config);
  if (fs.existsSync(file)) fs.rmSync(file, { force: true });
}

export function getScopeList(): string[] {
  return [...scopes];
}
