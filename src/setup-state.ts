export type SetupState =
  | "SETUP_REQUIRED"
  | "FOLDER_REQUIRED"
  | "LOGIN_START_REQUIRED"
  | "LOGIN_COMPLETE_REQUIRED"
  | "READY"
  | "BLOCKED";

export type SetupSnapshot = {
  configured: boolean;
  dataDirPersistent: boolean;
  loggedIn: boolean;
  tokenUsable: boolean;
  authMigrationRequired: boolean;
  authReason?: string;
  pendingLoginStateExists: boolean;
};

export type SetupContract = {
  state: SetupState;
  nextAction:
    | "CHECK_CONFIGURATION"
    | "SELECT_PROJECT_FOLDER"
    | "RUN_LOGIN_START"
    | "WAIT_FOR_USER_THEN_RUN_LOGIN_COMPLETE"
    | "WAIT_FOR_USER_REQUEST"
    | "REPORT_BLOCKER";
  nextCommand?: string;
  stopAfterAction: true;
  instruction: string;
};

export function determineSetupState(snapshot: SetupSnapshot): SetupState {
  if (!snapshot.configured) return "SETUP_REQUIRED";
  if (!snapshot.dataDirPersistent) return "FOLDER_REQUIRED";
  if (snapshot.loggedIn && snapshot.tokenUsable) return "READY";
  if (snapshot.authMigrationRequired) return "LOGIN_START_REQUIRED";
  if (snapshot.pendingLoginStateExists) return "LOGIN_COMPLETE_REQUIRED";
  const authFailure = classifyAuthFailure(snapshot.authReason);
  if (authFailure === "NETWORK_BLOCKED" || authFailure === "UNKNOWN_BLOCKED") {
    return "BLOCKED";
  }
  return "LOGIN_START_REQUIRED";
}

export type AuthFailureClass =
  | "NONE"
  | "LOGIN_REQUIRED"
  | "NETWORK_BLOCKED"
  | "UNKNOWN_BLOCKED";

export function classifyAuthFailure(reason: string | undefined): AuthFailureClass {
  if (!reason || reason === "NO_ACCOUNT_IN_CACHE" || reason === "NO_ACCESS_TOKEN") {
    return reason ? "LOGIN_REQUIRED" : "NONE";
  }
  if (reason === "AUTH_APP_CHANGED") return "LOGIN_REQUIRED";
  if (!reason.startsWith("TOKEN_ACQUISITION_FAILED:")) return "UNKNOWN_BLOCKED";
  if (
    /network_error|fetch failed|eai_again|etimedout|econnreset|econnrefused|enotfound|blocked-by-allowlist/i.test(
      reason
    )
  ) {
    return "NETWORK_BLOCKED";
  }
  if (/interaction_required|invalid_grant|login_required|consent_required/i.test(reason)) {
    return "LOGIN_REQUIRED";
  }
  return "UNKNOWN_BLOCKED";
}

export function buildSetupContract(
  snapshot: SetupSnapshot,
  selfCommand: string
): SetupContract {
  const state = determineSetupState(snapshot);

  switch (state) {
    case "SETUP_REQUIRED":
      return {
        state,
        nextAction: "CHECK_CONFIGURATION",
        stopAfterAction: true,
        instruction: "Hare configuration is missing. Report SETUP_REQUIRED and stop."
      };
    case "FOLDER_REQUIRED":
      return {
        state,
        nextAction: "SELECT_PROJECT_FOLDER",
        stopAfterAction: true,
        instruction:
          "This Cowork task was not started with a project folder. Report FOLDER_REQUIRED and stop. Tell the user to open a new Cowork task with their existing Hare project or folder selected, then paste the same prompt there."
      };
    case "LOGIN_START_REQUIRED":
      return {
        state,
        nextAction: "RUN_LOGIN_START",
        nextCommand: `${selfCommand} auth login-start`,
        stopAfterAction: true,
        instruction: snapshot.authMigrationRequired
          ? "Hare M365 Agent authentication permissions or application changed. Tell the user that one Microsoft sign-in is required, then run nextCommand unchanged once in the foreground, show the returned Microsoft URL and user code, and stop until the user finishes sign-in. Tell the user to sign in with their own company Microsoft account that they will use with Hare. Never name, recommend, or preselect a specific email address. Never start a background or detached poller."
          : "Run nextCommand unchanged once in the foreground, show the returned Microsoft URL and user code, then stop and wait for the user to finish sign-in. Tell the user to sign in with their own company Microsoft account that they will use with Hare. Never name, recommend, or preselect a specific email address. Never start a background or detached poller."
      };
    case "LOGIN_COMPLETE_REQUIRED":
      return {
        state,
        nextAction: "WAIT_FOR_USER_THEN_RUN_LOGIN_COMPLETE",
        nextCommand: `${selfCommand} auth login-complete`,
        stopAfterAction: true,
        instruction:
          "Keep the existing device code. After the user says login is complete, run nextCommand unchanged once in the foreground; it must finish within 25 seconds. COMPLETE is valid only after the persisted cache is verified."
      };
    case "READY":
      return {
        state,
        nextAction: "WAIT_FOR_USER_REQUEST",
        stopAfterAction: true,
        instruction: "Hare M365 Agent is ready. Wait for the user's Microsoft 365 lookup request."
      };
    case "BLOCKED":
      return buildBlockedSetupContract(
        snapshot.authReason ?? "Microsoft token validation failed. Do not start a new login flow."
      );
  }
}

export function buildBlockedSetupContract(reason: string): SetupContract {
  return {
    state: "BLOCKED",
    nextAction: "REPORT_BLOCKER",
    stopAfterAction: true,
    instruction: `Report this blocker in one sentence and stop. Do not start a new Microsoft sign-in or replace the existing cache: ${reason}`
  };
}
