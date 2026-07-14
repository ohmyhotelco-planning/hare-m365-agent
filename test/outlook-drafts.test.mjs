import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  createApprovedDraft,
  prepareDraft
} from "../dist/outlook-drafts.js";

function configFor(dataDir) {
  return {
    dataDir,
    policy: {
      allowDraftActions: true,
      maxDraftAttachmentBytes: 150 * 1024 * 1024,
      maxDraftTotalAttachmentBytes: 150 * 1024 * 1024
    }
  };
}

function fakeClient(overrides = {}) {
  const calls = [];
  const client = {
    calls,
    async get(url) {
      calls.push(["GET", url]);
      throw new Error(`Unexpected GET ${url}`);
    },
    async post(url, body) {
      calls.push(["POST", url, body]);
      if (url === "/me/messages") {
        return {
          id: "draft-new",
          subject: body.subject,
          isDraft: true,
          hasAttachments: false
        };
      }
      if (url.endsWith("/attachments")) return { id: "attachment" };
      throw new Error(`Unexpected POST ${url}`);
    },
    async patch(url, body) {
      calls.push(["PATCH", url, body]);
      return { id: "draft-thread", subject: "thread", isDraft: true, ...body };
    },
    async delete(url) {
      calls.push(["DELETE", url]);
    },
    async uploadChunk(url, bytes, start, total) {
      calls.push(["PUT", url, bytes.byteLength, start, total]);
    },
    ...overrides
  };
  return client;
}

test("new draft preview is deterministic and performs no write", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-draft-preview-"));
  const attachment = path.join(dataDir, "approval.txt");
  fs.writeFileSync(attachment, "approved attachment", "utf8");
  const client = fakeClient();
  const input = {
    kind: "new",
    to: ["one@example.com,two@example.com"],
    cc: ["copy@example.com"],
    subject: "Approved subject",
    body: "Approved body",
    contentType: "text",
    attachmentPaths: [attachment]
  };

  const first = await prepareDraft(configFor(dataDir), input, client);
  const second = await prepareDraft(configFor(dataDir), input, client);

  assert.equal(first.approvalToken, second.approvalToken);
  assert.deepEqual(first.preview.to, ["one@example.com", "two@example.com"]);
  assert.equal(first.preview.attachments[0].uploadMode, "direct");
  assert.equal(first.preview.warning, "This action creates an Outlook draft only. Hare cannot send mail.");
  assert.equal(client.calls.length, 0);
  fs.rmSync(dataDir, { recursive: true, force: true });
});

test("approval token is bound to the exact body", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-draft-approval-"));
  const client = fakeClient();
  const input = {
    kind: "new",
    to: ["one@example.com"],
    subject: "Subject",
    body: "Approved body",
    contentType: "text"
  };
  const prepared = await prepareDraft(configFor(dataDir), input, client);

  await assert.rejects(
    () => createApprovedDraft(
      configFor(dataDir),
      { ...input, body: "Changed body" },
      prepared.approvalToken,
      client
    ),
    /DRAFT_APPROVAL_REQUIRED/
  );
  assert.equal(client.calls.length, 0);
  fs.rmSync(dataDir, { recursive: true, force: true });
});

test("approved new draft creates the draft and adds a small attachment", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-draft-new-"));
  const attachment = path.join(dataDir, "small.txt");
  fs.writeFileSync(attachment, "small", "utf8");
  const client = fakeClient();
  const input = {
    kind: "new",
    to: ["one@example.com"],
    subject: "Subject",
    body: "Body",
    contentType: "text",
    attachmentPaths: [attachment]
  };
  const prepared = await prepareDraft(configFor(dataDir), input, client);
  const result = await createApprovedDraft(
    configFor(dataDir),
    input,
    prepared.approvalToken,
    client
  );

  assert.equal(result.stage, "DRAFT_CREATED");
  assert.equal(result.sendAvailable, false);
  assert.deepEqual(client.calls.map(([method, url]) => [method, url]), [
    ["POST", "/me/messages"],
    ["POST", "/me/messages/draft-new/attachments"]
  ]);
  fs.rmSync(dataDir, { recursive: true, force: true });
});

test("reply-all preview excludes the signed-in user and uses createReplyAll", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-draft-reply-all-"));
  const client = fakeClient({
    async get(url) {
      this.calls.push(["GET", url]);
      if (url.startsWith("/me/messages/")) {
        return {
          id: "source",
          subject: "Status",
          from: { emailAddress: { address: "sender@example.com" } },
          toRecipients: [
            { emailAddress: { address: "me@example.com" } },
            { emailAddress: { address: "peer@example.com" } }
          ],
          ccRecipients: [{ emailAddress: { address: "copy@example.com" } }]
        };
      }
      if (url.startsWith("/me?")) {
        return { mail: "me@example.com", userPrincipalName: "me@example.com" };
      }
      throw new Error(`Unexpected GET ${url}`);
    },
    async post(url, body) {
      this.calls.push(["POST", url, body]);
      if (url.endsWith("/createReplyAll")) {
        return {
          id: "draft-thread",
          subject: "RE: Status",
          isDraft: true,
          body: { contentType: "HTML", content: "<div>Original</div>" }
        };
      }
      throw new Error(`Unexpected POST ${url}`);
    }
  });
  const input = {
    kind: "replyAll",
    sourceMessageId: "source",
    body: "Reply body",
    contentType: "text"
  };
  const prepared = await prepareDraft(configFor(dataDir), input, client);

  assert.deepEqual(prepared.preview.to, ["sender@example.com", "peer@example.com"]);
  assert.deepEqual(prepared.preview.cc, ["copy@example.com"]);
  await createApprovedDraft(configFor(dataDir), input, prepared.approvalToken, client);
  assert.ok(client.calls.some(([method, url]) => method === "POST" && url.endsWith("/createReplyAll")));
  const patch = client.calls.find(([method]) => method === "PATCH");
  assert.match(patch[2].body.content, /Reply body/);
  assert.match(patch[2].body.content, /Original/);
  fs.rmSync(dataDir, { recursive: true, force: true });
});

test("forward draft uses createForward and preserves the original thread", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-draft-forward-"));
  const client = fakeClient({
    async get(url) {
      this.calls.push(["GET", url]);
      return {
        id: "source-forward",
        subject: "Travel plan",
        from: { emailAddress: { address: "sender@example.com" } }
      };
    },
    async post(url, body) {
      this.calls.push(["POST", url, body]);
      if (url.endsWith("/createForward")) {
        return {
          id: "draft-forward",
          subject: "FW: Travel plan",
          isDraft: true,
          body: { contentType: "HTML", content: "<div>Original forward</div>" }
        };
      }
      throw new Error(`Unexpected POST ${url}`);
    }
  });
  const input = {
    kind: "forward",
    sourceMessageId: "source-forward",
    to: ["recipient@example.com"],
    body: "Forward note",
    contentType: "text"
  };
  const prepared = await prepareDraft(configFor(dataDir), input, client);

  assert.equal(prepared.preview.subject, "FW: Travel plan");
  assert.equal(prepared.preview.originalThreadWillBeIncluded, true);
  await createApprovedDraft(configFor(dataDir), input, prepared.approvalToken, client);
  assert.ok(client.calls.some(([method, url]) => method === "POST" && url.endsWith("/createForward")));
  const patch = client.calls.find(([method]) => method === "PATCH");
  assert.match(patch[2].body.content, /Forward note/);
  assert.match(patch[2].body.content, /Original forward/);
  fs.rmSync(dataDir, { recursive: true, force: true });
});

test("attachment failure deletes the incomplete draft", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-draft-cleanup-"));
  const attachment = path.join(dataDir, "failure.txt");
  fs.writeFileSync(attachment, "failure", "utf8");
  const client = fakeClient({
    async post(url, body) {
      this.calls.push(["POST", url, body]);
      if (url === "/me/messages") return { id: "draft-to-clean", isDraft: true };
      if (url.endsWith("/attachments")) throw new Error("attachment rejected");
      throw new Error(`Unexpected POST ${url}`);
    }
  });
  const input = {
    kind: "new",
    to: ["one@example.com"],
    subject: "Subject",
    body: "Body",
    contentType: "text",
    attachmentPaths: [attachment]
  };
  const prepared = await prepareDraft(configFor(dataDir), input, client);

  await assert.rejects(
    () => createApprovedDraft(configFor(dataDir), input, prepared.approvalToken, client),
    /DRAFT_CREATION_FAILED_CLEANED_UP/
  );
  assert.ok(client.calls.some(([method, url]) => method === "DELETE" && url.includes("draft-to-clean")));
  fs.rmSync(dataDir, { recursive: true, force: true });
});

test("large attachments use an Outlook upload session in sequential chunks", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-draft-large-"));
  const attachment = path.join(dataDir, "large.bin");
  fs.writeFileSync(attachment, Buffer.alloc(5 * 1024 * 1024, 7));
  const client = fakeClient({
    async post(url, body) {
      this.calls.push(["POST", url, body]);
      if (url === "/me/messages") return { id: "draft-large", isDraft: true };
      if (url.endsWith("/createUploadSession")) return { uploadUrl: "https://outlook.office.com/upload" };
      throw new Error(`Unexpected POST ${url}`);
    }
  });
  const input = {
    kind: "new",
    to: ["one@example.com"],
    subject: "Large attachment",
    body: "Body",
    contentType: "text",
    attachmentPaths: [attachment]
  };
  const prepared = await prepareDraft(configFor(dataDir), input, client);
  assert.equal(prepared.preview.attachments[0].uploadMode, "uploadSession");
  await createApprovedDraft(configFor(dataDir), input, prepared.approvalToken, client);

  const uploadCalls = client.calls.filter(([method]) => method === "PUT");
  assert.equal(uploadCalls.length, 2);
  assert.equal(uploadCalls[0][3], 0);
  assert.equal(uploadCalls[0][4], 5 * 1024 * 1024);
  assert.equal(uploadCalls[1][3], uploadCalls[0][2]);
  assert.equal(uploadCalls[1][4], 5 * 1024 * 1024);
  fs.rmSync(dataDir, { recursive: true, force: true });
});
