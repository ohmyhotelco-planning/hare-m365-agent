# OMH M365 Agent Operation Model

## Objective

Provide a local, policy-controlled execution layer that lets an LLM answer Microsoft 365 work questions using Microsoft Graph delegated permissions.

The CLI is not intended to be a human-first tool. It is a stable command surface for LLMs, scripts, and installers.

## Operating Principle

Human attention is reserved for authentication, approval, policy decisions, and irreversible or external-facing actions.

Safe reads, lookups, summaries, validation checks, retries, and smoke tests should be automated through the CLI or by the LLM that calls it.

## Permission Model

- Azure Enterprise Application permissions define the maximum Graph permission scope.
- Delegated login means the operation runs as the signed-in user.
- The CLI policy and LLM instructions define which operations may actually be invoked.
- The initial POC is read-focused. Write, send, delete, upload, and share actions remain closed.

## Human Responsibilities

- Approve the Azure Application and Graph permission scope.
- Confirm `.env` delegated public-client configuration is present.
- Complete Microsoft login and consent flows.
- Approve any future write/send/delete/share operation before it is implemented or run.
- Rotate or revoke secrets if exposure is suspected.

## LLM Responsibilities

- Use the documented CLI, not ad hoc Graph scripts.
- Run diagnostics before asking the human for help.
- Use small limits and summarize results.
- Avoid raw dumps of mail, chats, files, tokens, or cache contents.
- Ask the human only when authentication, approval, missing policy, or irreversible work is required.

## Current Verified Capabilities

- Local configuration diagnostics
- Login status checks
- Outlook Inbox lookup
- Teams team lookup
- Teams chat lookup
- Teams chat message lookup
- OneDrive/SharePoint-visible file search
- Explicit file download

## Operational Gaps Before Pilot

- Replace file-based token cache with OS-backed credential storage.
- Add a stronger `auth status` check that verifies silent token usability.
- Decide whether `Chat.ReadWrite` should be removed from the requested scope list.
- Expand SharePoint and Teams file search beyond `/me/drive/root/search`.
- Define release ownership, versioning, update, rollback, and audit expectations.
