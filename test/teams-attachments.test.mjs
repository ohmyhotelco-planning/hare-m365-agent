import assert from "node:assert/strict";
import test from "node:test";
import {
  buildTeamsMessagePath,
  downloadTeamsMessageAttachment,
  listTeamsMessageAttachments
} from "../dist/teams-attachments.js";

function config() {
  return {
    policy: {
      allowDownloads: true
    }
  };
}

test("Teams message path encodes chat and message IDs", () => {
  assert.equal(
    buildTeamsMessagePath("19:chat/id@thread.v2", "message/id"),
    "/chats/19%3Achat%2Fid%40thread.v2/messages/message%2Fid"
  );
});

test("attachment listing returns safe metadata without the sharing URL", async () => {
  const client = {
    async get() {
      return {
        attachments: [{
          id: "attachment-1",
          name: "Payment Notification.zip",
          contentType: "reference",
          contentUrl: "https://ohmylab.sharepoint.com/sites/acc/Shared%20Documents/Payment%20Notification.zip"
        }]
      };
    },
    async downloadDriveItem() {
      throw new Error("not used");
    }
  };

  const result = await listTeamsMessageAttachments(config(), "chat-1", "message-1", client);
  assert.equal(result.list.returnedCount, 1);
  assert.equal(result.attachments[0].downloadable, true);
  assert.equal("contentUrl" in result.attachments[0], false);
});

test("Teams reference attachment resolves a site path and reuses the existing downloader", async () => {
  const calls = [];
  const client = {
    async get(url) {
      calls.push(url);
      if (url === "/chats/chat-1/messages/message-1") {
        return {
          attachments: [{
            id: "attachment-1",
            name: "VAT-to-FAST folder.zip",
            contentType: "reference",
            contentUrl: "https://ohmylab.sharepoint.com/sites/acc/Shared%20Documents/Tax/VAT-to-FAST%20folder.zip?web=1"
          }]
        };
      }
      if (url.startsWith("/sites/ohmylab.sharepoint.com:/sites/acc")) {
        return { id: "site-1", displayName: "Accounting" };
      }
      if (url.startsWith("/sites/site-1/drives")) {
        return {
          value: [{
            id: "drive-1",
            name: "Documents",
            webUrl: "https://ohmylab.sharepoint.com/sites/acc/Shared%20Documents"
          }]
        };
      }
      if (url.startsWith("/drives/drive-1/root:/Tax/VAT-to-FAST%20folder.zip:")) {
        return {
          id: "item-1",
          name: "VAT-to-FAST folder.zip",
          size: 26_278_538,
          file: {},
          parentReference: { driveId: "drive-1" }
        };
      }
      throw new Error(`Unexpected GET ${url}`);
    },
    async downloadDriveItem(driveId, itemId, name, approvalToken) {
      calls.push({ driveId, itemId, name, approvalToken });
      return "C:\\downloads\\VAT-to-FAST folder.zip";
    }
  };

  const result = await downloadTeamsMessageAttachment(
    config(),
    "chat-1",
    "message-1",
    "attachment-1",
    undefined,
    undefined,
    client
  );

  assert.equal(result.stage, "DOWNLOADED");
  assert.equal(result.file.size, 26_278_538);
  assert.deepEqual(calls.at(-1), {
    driveId: "drive-1",
    itemId: "item-1",
    name: "VAT-to-FAST folder.zip",
    approvalToken: undefined
  });
});

test("unapproved hosts are never resolved or downloaded", async () => {
  let calls = 0;
  const client = {
    async get() {
      calls += 1;
      return {
        attachments: [{
          id: "attachment-1",
          name: "external.zip",
          contentType: "reference",
          contentUrl: "https://example.com/external.zip"
        }]
      };
    },
    async downloadDriveItem() {
      calls += 1;
      throw new Error("must not download");
    }
  };

  await assert.rejects(
    () => downloadTeamsMessageAttachment(
      config(),
      "chat-1",
      "message-1",
      "attachment-1",
      undefined,
      undefined,
      client
    ),
    /not a downloadable SharePoint reference file/
  );
  assert.equal(calls, 1);
});

test("large-file approval previews pass through without downloading", async () => {
  const pending = {
    stage: "AWAITING_USER_APPROVAL",
    preview: { source: "sharepoint", name: "large.zip" },
    approval: { token: "test-token" }
  };
  const client = {
    async get(url) {
      if (url.startsWith("/chats/")) {
        return {
          attachments: [{
            id: "attachment-1",
            name: "large.zip",
            contentType: "reference",
            contentUrl: "https://ohmylab-my.sharepoint.com/personal/user/Documents/large.zip"
          }]
        };
      }
      if (url.startsWith("/sites/ohmylab-my.sharepoint.com:/personal/user")) {
        return { id: "site-1" };
      }
      if (url.startsWith("/sites/site-1/drives")) {
        return {
          value: [{
            id: "drive-1",
            webUrl: "https://ohmylab-my.sharepoint.com/personal/user/Documents"
          }]
        };
      }
      return {
        id: "item-1",
        name: "large.zip",
        size: 200_000_000,
        file: {},
        parentReference: { driveId: "drive-1" }
      };
    },
    async downloadDriveItem() {
      return pending;
    }
  };

  const result = await downloadTeamsMessageAttachment(
    config(),
    "chat-1",
    "message-1",
    "attachment-1",
    undefined,
    undefined,
    client
  );
  assert.equal(result, pending);
});

test("opaque sharing links and non-default ports are not treated as downloadable", async () => {
  const client = {
    async get() {
      return {
        attachments: [
          {
            id: "opaque",
            name: "opaque.zip",
            contentType: "reference",
            contentUrl: "https://ohmylab.sharepoint.com/:u:/s/acc/opaque"
          },
          {
            id: "port",
            name: "port.zip",
            contentType: "reference",
            contentUrl: "https://ohmylab.sharepoint.com:444/sites/acc/Shared%20Documents/port.zip"
          }
        ]
      };
    },
    async downloadDriveItem() {
      throw new Error("must not download");
    }
  };

  const result = await listTeamsMessageAttachments(config(), "chat-1", "message-1", client);
  assert.deepEqual(result.attachments.map((item) => item.downloadable), [false, false]);
});

test("folder attachments are rejected before the downloader is called", async () => {
  let downloadCalls = 0;
  const client = {
    async get(url) {
      if (url.startsWith("/chats/")) {
        return {
          attachments: [{
            id: "attachment-1",
            name: "Folder",
            contentType: "reference",
            contentUrl: "https://ohmylab.sharepoint.com/sites/acc/Shared%20Documents/Folder"
          }]
        };
      }
      if (url.startsWith("/sites/ohmylab.sharepoint.com:/sites/acc")) {
        return { id: "site-1" };
      }
      if (url.startsWith("/sites/site-1/drives")) {
        return {
          value: [{
            id: "drive-1",
            webUrl: "https://ohmylab.sharepoint.com/sites/acc/Shared%20Documents"
          }]
        };
      }
      return {
        id: "folder-1",
        name: "Folder",
        folder: {},
        parentReference: { driveId: "drive-1" }
      };
    },
    async downloadDriveItem() {
      downloadCalls += 1;
      throw new Error("must not download");
    }
  };

  await assert.rejects(
    () => downloadTeamsMessageAttachment(
      config(),
      "chat-1",
      "message-1",
      "attachment-1",
      undefined,
      undefined,
      client
    ),
    /does not resolve to a file/
  );
  assert.equal(downloadCalls, 0);
});
