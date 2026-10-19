import type { Command } from "commander";
import { requireConfigured, type AppConfig } from "./config.js";
import { resolveGroup } from "./groups-directory.js";
import { requireMicrosoft365Group } from "./groups-common.js";
import { listGroupPosts } from "./groups-posts.js";
import { readGroupCursorInput } from "./groups-cursor-input.js";
import { listGroupPostAttachments, downloadGroupPostAttachment } from "./groups-attachments.js";

type EmitJson = (data: unknown, out?: string) => void;
type GroupOption = { group: string; out?: string };
type AttachmentOptions = GroupOption & { threadId: string; postId: string };

export function registerGroupCommands(program: Command, config: AppConfig, emitJson: EmitJson): void {
  const groups = program.command("groups").description("M365 group directory, post bodies and attachments (delegated read only; not shared mailboxes)");
  groups.command("resolve")
    .description("Resolve an exact SMTP address, SMTP proxy alias, mailNickname or directory group ID; return type evidence")
    .requiredOption("--group <address-alias-or-id>", "exact group SMTP address, mailNickname or directory object ID")
    .option("--out <path>", "write JSON under Hare resultsDir")
    .action(async (options: GroupOption) => {
      requireConfigured(config);
      emitJson({ group: await resolveGroup(config, options.group) }, options.out);
    });
  groups.command("posts")
    .description("Traverse conversations/threads and read complete posts; date is post createdDateTime, not summary time")
    .requiredOption("--group <address-alias-or-id>", "exact M365 group identity; no self-mailbox fallback")
    .option("--since <YYYY-MM-DD>", "first creation date, Asia/Seoul; default recent 90 days")
    .option("--until <YYYY-MM-DD>", "last creation date inclusive, Asia/Seoul; default today")
    .option("--limit <number>", "maximum returned posts per call (1-100); traversal order is not newest-first", "20")
    .option("--cursor <cursor>", "resume the same group and date range using page.nextCursor")
    .option("--cursor-file <path>", "read a UTF-8 cursor or previous groups posts result JSON; avoids command-line length limits")
    .option("--include-subject", "read each returned post's PidTagSubject within the existing budget; never substitute thread topic")
    .option("--out <path>", "write JSON under Hare resultsDir")
    .action(async (options: GroupOption & { since?: string; until?: string; limit: string; cursor?: string; cursorFile?: string; includeSubject?: boolean }) => {
      const cursor = readGroupCursorInput(options);
      requireConfigured(config);
      const group = await resolveGroup(config, options.group);
      emitJson(await listGroupPosts(config, group, Number(options.limit), { ...options, cursor }), options.out);
    });
  const attachments = groups.command("attachments").description("Group post attachments; Graph application permissions are unsupported for these endpoints");
  const common = (command: Command) => command
    .requiredOption("--group <address-alias-or-id>", "M365 group owning this thread and post")
    .requiredOption("--thread-id <id>", "threadId from groups posts")
    .requiredOption("--post-id <id>", "post id from groups posts; never a mailbox message ID");
  common(attachments.command("list").description("Read attachment metadata for the exact group/thread/post"))
    .option("--limit <number>", "maximum returned attachments", "20")
    .option("--cursor <cursor>", "continue the same group/thread/post attachment list")
    .option("--out <path>", "write JSON under Hare resultsDir")
    .action(async (options: AttachmentOptions & { limit: string; cursor?: string }) => {
      requireConfigured(config);
      const group = await resolveGroup(config, options.group);
      requireMicrosoft365Group(group);
      emitJson(await listGroupPostAttachments(config, { groupId: group.id, threadId: options.threadId, postId: options.postId }, Number(options.limit), { cursor: options.cursor }), options.out);
    });
  common(attachments.command("download").description("Download raw file/item bytes within existing size/approval policy; reference attachments unsupported"))
    .requiredOption("--attachment-id <id>", "attachment ID from groups attachments list")
    .option("--name <filename>", "safe output filename; existing files are never overwritten")
    .option("--approval-token <token>", "one-time token after explicit approval of the exact large-download preview")
    .action(async (options: AttachmentOptions & { attachmentId: string; name?: string; approvalToken?: string }) => {
      requireConfigured(config);
      const group = await resolveGroup(config, options.group);
      requireMicrosoft365Group(group);
      emitJson(await downloadGroupPostAttachment(config, { groupId: group.id, threadId: options.threadId, postId: options.postId }, options.attachmentId, options.name, options.approvalToken));
    });
}
