export const FIXED_HOST_DATA_DIRS = {
  windows: "%USERPROFILE%\\HareM365Agent",
  mac: "~/HareM365Agent"
} as const;

export type SetupState =
  | "SETUP_REQUIRED"
  | "FOLDER_REQUIRED"
  | "LOGIN_START_REQUIRED"
  | "LOGIN_COMPLETE_REQUIRED"
  | "READY"
  | "BLOCKED";

type EvaluatedSetupState = Exclude<SetupState, "BLOCKED">;

export type SetupSnapshot = {
  configured: boolean;
  dataDirPersistent: boolean;
  loggedIn: boolean;
  tokenUsable: boolean;
  pendingLoginStateExists: boolean;
};

export type SetupContract = {
  state: SetupState;
  nextAction:
    | "CHECK_CONFIGURATION"
    | "CONNECT_FIXED_FOLDER"
    | "RUN_LOGIN_START"
    | "WAIT_FOR_USER_THEN_RUN_LOGIN_COMPLETE"
    | "WAIT_FOR_USER_REQUEST"
    | "REPORT_BLOCKER";
  nextCommand?: string;
  stopAfterAction: true;
  fixedHostDataDirs?: typeof FIXED_HOST_DATA_DIRS;
  instruction: string;
};

export function determineSetupState(snapshot: SetupSnapshot): EvaluatedSetupState {
  if (!snapshot.configured) return "SETUP_REQUIRED";
  if (!snapshot.dataDirPersistent) return "FOLDER_REQUIRED";
  if (snapshot.loggedIn && snapshot.tokenUsable) return "READY";
  if (snapshot.pendingLoginStateExists) return "LOGIN_COMPLETE_REQUIRED";
  return "LOGIN_START_REQUIRED";
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
        nextAction: "CONNECT_FIXED_FOLDER",
        stopAfterAction: true,
        fixedHostDataDirs: FIXED_HOST_DATA_DIRS,
        instruction:
          "The current data directory is not persistent. A Cowork Linux OS-default path such as /root/.local/share is container-local and must not be used for login. Use the OS-specific fixed HareM365Agent host path, request access to that folder, and rerun startup with --data-dir set to the mounted folder root."
      };
    case "LOGIN_START_REQUIRED":
      return {
        state,
        nextAction: "RUN_LOGIN_START",
        nextCommand: `${selfCommand} auth login-start`,
        stopAfterAction: true,
        instruction:
          "Run nextCommand unchanged once in the foreground, show the returned Microsoft URL and user code, then stop and wait for the user to finish sign-in. Never start a background or detached poller."
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
  }
}

export function buildBlockedSetupContract(reason: string): SetupContract {
  return {
    state: "BLOCKED",
    nextAction: "REPORT_BLOCKER",
    stopAfterAction: true,
    instruction: `Report this blocker in one sentence and stop: ${reason}`
  };
}
