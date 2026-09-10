import { ProxyAwareNetworkClient } from "./msal-network.js";
import { fetchWithProxy } from "./proxy.js";

export type ExecutionEnvironment = "codex" | "cowork" | "unknown";

const endpoint = "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration";
const timeoutMs = 3_000;

export const networkExecutionGuidance = `Before startup/auth status or the first requested M365 read in a new execution environment, run the same Hare command prefix and exact --data-dir followed by network check --environment <actual-host> (codex, cowork, or unknown). This check does not read the authentication cache or modify the data directory. Do not repeat it before every page or Graph request while the approved execution context remains unchanged.
- REACHABLE: check authentication in that same environment. Only loggedIn=true and tokenUsable=true permit reads; connectivity alone does not establish login, Graph permissions, or SharePoint access.
- EXECUTION_PERMISSION_REQUIRED in Codex: invoke the host's standard execution-permission request for the exact check. Do not simply report that Hare cannot work or ask the user to force another attempt. Only after permission is granted, repeat the check once, then run startup/auth status and the originally requested read with the same executable and dataDir in that approved context. The --environment option does not grant permission. If denied, unavailable, or still blocked, stop. Never automatically retry a write.
- NETWORK_PERMISSION_REQUIRED: an explicit allowlist block is not an elevation/retry case, including in Codex. Report the blocked domain and request the policy change. In Cowork, open a new task with the same project after the change takes effect.
- NETWORK_CHECK_BLOCKED: report the safe cause and stop; do not assume that all network failures are allowlist problems or token expiry. Cowork must remain in its allowlist-governed session shell, even for EACCES. Unknown hosts must not be treated as Codex.
- Never change the dataDir, use another authentication cache, restart sign-in, or bypass a denied permission/policy to solve a connectivity failure. When a later read returns AUTH_CHECK_BLOCKED, use this bounded check once; preserve genuine login-required and data-permission errors as distinct failures.`;

export async function checkMicrosoftConnectivity(
  environment: ExecutionEnvironment,
  network = new ProxyAwareNetworkClient((url, init) => fetchWithProxy(url, { ...init, redirect: "manual" }))
) {
  const base = {
    environment,
    hostname: "login.microsoftonline.com",
    timeoutMs,
    authenticationChecked: false,
    cacheAccessed: false
  };
  let code: string;
  let httpStatus: number | undefined;
  try {
    // Public metadata only: never attach an account, token, or tenant-specific URL.
    const response = await network.sendGetHeadersAsync(endpoint, timeoutMs);
    httpStatus = response.status;
    if (response.status === 200) {
      return {
        ...base, ok: true, state: "REACHABLE", httpStatus,
        nextAction: "CHECK_AUTH_IN_SAME_ENVIRONMENT",
        instruction: "Microsoft public login metadata is reachable in this execution environment. This does not verify authentication or Graph/SharePoint access. Run startup/auth status with the same Hare executable and exact dataDir here; only READY permits the requested read. Keep using this approved environment for this task while its permission grant remains valid."
      };
    }
    const proxyError = Object.entries(response.headers).find(([name]) => name.toLowerCase() === "x-proxy-error")?.[1];
    code = response.status === 403 && proxyError === "blocked-by-allowlist"
      ? "PROXY_ALLOWLIST_BLOCKED" : "HTTP_ERROR";
  } catch {
    code = network.takeNetworkFailure()?.code ?? "NETWORK_ERROR";
  }

  const allowlistBlocked = code === "PROXY_ALLOWLIST_BLOCKED";
  const permissionRequired = !allowlistBlocked && environment === "codex" && (code === "EACCES" || code === "EPERM");
  return {
    ...base, ok: false,
    state: allowlistBlocked ? "NETWORK_PERMISSION_REQUIRED"
      : permissionRequired ? "EXECUTION_PERMISSION_REQUIRED" : "NETWORK_CHECK_BLOCKED",
    code, httpStatus,
    nextAction: permissionRequired ? "REQUEST_EXECUTION_PERMISSION" : "REPORT_BLOCKER",
    instruction: permissionRequired
      ? "In Codex only, use the host's standard execution-permission request for this exact network check instead of asking the user to force a retry. Do not run outside the sandbox until permission is granted. After approval, repeat the check once with the same Hare executable and exact dataDir; if REACHABLE, check auth and perform the requested read in that approved environment. If denied, unavailable, or still blocked, stop. Never change dataDir, reset the cache, restart login, or automatically retry a write."
      : "Report the blocked hostname and safe error code. Keep the existing dataDir and cache; do not restart login or infer token expiry. In Cowork, stay in the session shell governed by its domain allowlist; never move to a local shell to bypass it. For an explicit allowlist block, request the domain policy change and use a new Cowork task with the same project after it is applied. Otherwise ask for network/environment diagnosis, not automatic allowlisting or repeated retries."
  };
}
