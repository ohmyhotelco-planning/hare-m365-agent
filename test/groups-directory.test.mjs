import assert from "node:assert/strict";
import test from "node:test";
import { resolveGroup } from "../dist/groups-directory.js";
import { requireMicrosoft365Group } from "../dist/groups-common.js";

const firstId = "11111111-1111-4111-8111-111111111111";
const secondId = "22222222-2222-4222-8222-222222222222";
const address = "team@example.test";
const root = "https://graph.microsoft.com/v1.0";

function group(overrides = {}) {
  return {
    id: firstId,
    displayName: "Synthetic Team",
    mail: address,
    mailNickname: "team",
    proxyAddresses: [`SMTP:${address}`, "smtp:alias@example.test"],
    groupTypes: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    ...overrides
  };
}

function mockPages(...pages) {
  const calls = [];
  return {
    calls,
    async get(url, options) {
      calls.push({ url, options });
      assert.ok(calls.length <= pages.length, "Unexpected extra directory request");
      const result = pages[calls.length - 1];
      if (result instanceof Error) throw result;
      return result;
    }
  };
}

function requestUrl(call) {
  return new URL(call.url.startsWith("/") ? `${root}${call.url}` : call.url);
}

test("primary SMTP resolves only through groups with ordinary exact filters", async () => {
  const client = mockPages({ value: [group()] });
  const result = await resolveGroup({}, " TEAM@example.test ", client);
  assert.equal(result.id, firstId);
  assert.equal(result.mail, address);
  assert.equal(result.classification, "microsoft365");
  assert.deepEqual(result.groupTypes, ["Unified"]);
  assert.equal(result.mailEnabled, true);
  assert.equal(result.securityEnabled, false);
  assert.equal(result.exchangeRecipientType, "unverified");
  assert.ok(result.evidence.includes("mail: exact SMTP match"));
  requireMicrosoft365Group(result);
  const url = requestUrl(client.calls[0]);
  assert.equal(url.pathname, "/v1.0/groups");
  assert.equal(url.searchParams.get("$filter"), "mail eq 'TEAM@example.test' or proxyAddresses/any(p:p eq 'smtp:TEAM@example.test') or proxyAddresses/any(p:p eq 'SMTP:TEAM@example.test')");
  assert.equal(url.searchParams.has("$search"), false);
  assert.equal(url.searchParams.has("$count"), false);
  assert.equal(client.calls[0].options.headers, undefined);
  assert.ok(client.calls[0].options.totalTimeoutMs > 0);
  assert.ok(client.calls[0].options.totalTimeoutMs <= 15_000);
  for (const field of ["id", "mail", "proxyAddresses", "mailNickname", "groupTypes", "mailEnabled", "securityEnabled"]) {
    assert.ok(url.searchParams.get("$select").split(",").includes(field));
  }
});

for (const prefix of ["smtp:", "SMTP:"]) {
  test(`SMTP alias matches exact ${prefix} proxy address and retains primary mail`, async () => {
    const client = mockPages({ value: [group({ proxyAddresses: [`${prefix}ALIAS@example.test`] })] });
    const result = await resolveGroup({}, "alias@example.test", client);
    assert.equal(result.mail, address);
    assert.ok(result.evidence.includes("proxyAddresses: exact SMTP match"));
    assert.equal(result.evidence.includes("mail: exact SMTP match"), false);
  });
}

test("bare alias uses exact mailNickname only, not display name or SMTP inference", async () => {
  const client = mockPages({ value: [group({ mail: null, proxyAddresses: [] })] });
  const result = await resolveGroup({}, "TEAM", client);
  assert.equal(requestUrl(client.calls[0]).searchParams.get("$filter"), "mailNickname eq 'TEAM'");
  assert.equal(result.mail, null);
  assert.ok(result.evidence.includes("mailNickname: exact bare alias match (not SMTP evidence)"));
  await assert.rejects(resolveGroup({}, "team", mockPages({ value: [group({ mailNickname: "different", displayName: "team" })] })), /GROUP_DIRECTORY_NO_MATCH/);
});

test("OData string literals are escaped in SMTP and bare-alias filters", async () => {
  for (const input of ["o'hare@example.test", "o'hare"]) {
    const client = mockPages({ value: [group({ mail: input, mailNickname: input })] });
    await resolveGroup({}, input, client);
    const filter = requestUrl(client.calls[0]).searchParams.get("$filter");
    assert.ok(filter.includes("o''hare"));
    assert.equal(filter.includes("o'hare"), false);
  }
});

test("bare alias ambiguity or unavailable candidate evidence prevents selection", async () => {
  await assert.rejects(resolveGroup({}, "team", mockPages({ value: [group(), group({ id: secondId })] })), /GROUP_DIRECTORY_AMBIGUOUS/);
  for (const missing of [undefined, null]) {
    await assert.rejects(resolveGroup({}, "team", mockPages({
      value: [group(), group({ id: secondId, mailNickname: missing })]
    })), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE/);
  }
});

test("GUID resolution directly reads the matching group ID", async () => {
  const id = "abcdefab-1234-4321-8123-abcdefabcdef";
  const client = mockPages(group({ id, mail: null }));
  const result = await resolveGroup({}, id.toUpperCase(), client);
  assert.equal(requestUrl(client.calls[0]).pathname, `/v1.0/groups/${id.toUpperCase()}`);
  assert.equal(requestUrl(client.calls[0]).searchParams.has("$filter"), false);
  assert.equal(result.id, id);
  assert.equal(result.mail, null);
  assert.ok(result.evidence.includes("id: exact GUID match"));
  await assert.rejects(resolveGroup({}, firstId, mockPages(group({ id: secondId }))), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE/);
});

test("zero visible matches are not proof of group absence or a mailbox fallback", async () => {
  const client = mockPages({ value: [] });
  await assert.rejects(resolveGroup({}, address, client), /GROUP_DIRECTORY_NO_MATCH:.*not proof.*does not exist/);
  assert.equal(client.calls.length, 1);
  assert.equal(requestUrl(client.calls[0]).pathname, "/v1.0/groups");
});

test("partial addresses and non-SMTP proxies never match", async () => {
  const client = mockPages({ value: [group({ mail: "other@example.test", proxyAddresses: [`sip:${address}`, `smtp:prefix-${address}`] })] });
  await assert.rejects(resolveGroup({}, address, client), /GROUP_DIRECTORY_NO_MATCH/);
});

test("ambiguity includes primary versus alias matches without preferring either", async () => {
  const client = mockPages({ value: [group(), group({ id: secondId, mail: "other@example.test", proxyAddresses: [`smtp:${address}`] })] });
  await assert.rejects(resolveGroup({}, address, client), /GROUP_DIRECTORY_AMBIGUOUS/);
});

for (const [status, code] of [[403, "ACCESS_DENIED"], [401, "AUTH_BLOCKED"], [404, "NOT_FOUND_OR_UNAVAILABLE"], [400, "API_UNSUPPORTED"]]) {
  test(`HTTP ${status} is a DIRECTORY-stage error without fallback or consent advice`, async () => {
    for (const input of [address, firstId]) {
      const client = mockPages(new Error(`Graph GET failed (${status} Synthetic): private-response-marker`));
      await assert.rejects(resolveGroup({}, input, client), (error) => {
        assert.ok(error.message.startsWith(`GROUP_DIRECTORY_${code}:`));
        assert.equal(error.message.includes("private-response-marker"), false);
        assert.equal(/sign in again|grant .*All/i.test(error.message), false);
        return true;
      });
      assert.equal(client.calls.length, 1);
    }
  });
}

for (const [classification, mailEnabled, securityEnabled] of [
  ["distribution", true, false],
  ["mail-enabled-security", true, true],
  ["security", false, true]
]) {
  test(`${classification} resolves but the common read guard rejects it`, async () => {
    const result = await resolveGroup({}, firstId, mockPages(group({ groupTypes: [], mailEnabled, securityEnabled })));
    assert.equal(result.classification, classification);
    assert.equal(result.exchangeRecipientType, "unverified");
    assert.throws(() => requireMicrosoft365Group(result), /GROUP_TYPE_UNSUPPORTED/);
  });
}

test("missing classification fields without Unified remain unknown instead of treating missing as false", async () => {
  for (const field of ["groupTypes", "mailEnabled", "securityEnabled"]) {
    for (const missing of [undefined, null]) {
      const result = await resolveGroup({}, firstId, mockPages(group({ groupTypes: [], [field]: missing })));
      assert.equal(result[field], null);
      assert.equal(result.classification, "unknown");
      assert.throws(() => requireMicrosoft365Group(result), /GROUP_TYPE_UNSUPPORTED/);
    }
  }
  const minimal = await resolveGroup({}, firstId, mockPages({ id: firstId }));
  assert.equal(minimal.mail, null);
  assert.equal(minimal.classification, "unknown");
  assert.equal(minimal.exchangeRecipientType, "unverified");
});

test("actual Unified evidence classifies Microsoft 365 with unavailable boolean fields", async () => {
  for (const missing of [undefined, null]) {
    for (const fields of [
      { mailEnabled: missing },
      { securityEnabled: missing },
      { mailEnabled: missing, securityEnabled: missing }
    ]) {
      const result = await resolveGroup({}, firstId, mockPages(group(fields)));
      assert.equal(result.classification, "microsoft365");
      assert.deepEqual(result.groupTypes, ["Unified"]);
      for (const field of Object.keys(fields)) assert.equal(result[field], null);
      assert.ok(result.evidence.includes('groupTypes: ["Unified"]'));
      assert.equal(result.exchangeRecipientType, "unverified");
      requireMicrosoft365Group(result);
    }
  }
});

test("Unified must occur as an exact groupTypes element", async () => {
  const result = await resolveGroup({}, firstId, mockPages(group({ groupTypes: ["unified"], mailEnabled: false, securityEnabled: false })));
  assert.equal(result.classification, "unknown");
  assert.throws(() => requireMicrosoft365Group(result), /GROUP_TYPE_UNSUPPORTED/);
});

test("complete paging finds a later match and does not select before the last page", async () => {
  const next = `${root}/groups?$skiptoken=synthetic`;
  const client = mockPages(
    { value: [group({ id: secondId, mail: "other@example.test", proxyAddresses: [] })], "@odata.nextLink": next, "@odata.count": 2 },
    { value: [group()] }
  );
  assert.equal((await resolveGroup({}, address, client)).id, firstId);
  assert.equal(client.calls[1].url, next);
  assert.equal(client.calls[1].options.headers, undefined);
  assert.ok(client.calls[1].options.totalTimeoutMs <= client.calls[0].options.totalTimeoutMs);
  await assert.rejects(resolveGroup({}, address, mockPages(
    { value: [group()], "@odata.nextLink": next },
    { value: [group({ id: secondId })] }
  )), /GROUP_DIRECTORY_AMBIGUOUS/);
});

test("empty pages with valid continuations are followed to an exact match", async () => {
  const next = `${root}/groups?$skiptoken=empty-first`;
  const last = `${root}/groups?$skiptoken=empty-second`;
  const client = mockPages(
    { value: [], "@odata.nextLink": next },
    { value: [], "@odata.nextLink": last },
    { value: [group()] }
  );
  assert.equal((await resolveGroup({}, address, client)).id, firstId);
  assert.equal(client.calls.length, 3);
  assert.equal(client.calls[1].url, next);
  assert.equal(client.calls[2].url, last);
});

test("empty-page continuations still enforce loop and page limits", async () => {
  const next = `${root}/groups?$skiptoken=empty-loop`;
  const loop = mockPages(
    { value: [], "@odata.nextLink": next },
    { value: [], "@odata.nextLink": next }
  );
  await assert.rejects(resolveGroup({}, address, loop), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE:.*loops/);
  assert.equal(loop.calls.length, 2);
  const bounded = mockPages(...Array.from({ length: 10 }, (_, index) => ({
    value: [], "@odata.nextLink": `${root}/groups?$skiptoken=empty-${index}`
  })));
  await assert.rejects(resolveGroup({}, address, bounded), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE:.*page limit/);
  assert.equal(bounded.calls.length, 10);
});

test("a later permission error rejects even an already observed exact match", async () => {
  await assert.rejects(resolveGroup({}, address, mockPages(
    { value: [group()], "@odata.nextLink": "/groups?$skiptoken=synthetic" },
    new Error("Graph GET failed (403 Forbidden): synthetic")
  )), /GROUP_DIRECTORY_ACCESS_DENIED/);
});

for (const next of [
  "https://example.test/v1.0/groups?$skiptoken=x",
  "https://graph.microsoft.com.example.test/v1.0/groups",
  "http://graph.microsoft.com/v1.0/groups",
  "https://user:password@graph.microsoft.com/v1.0/groups",
  "https://graph.microsoft.com:444/v1.0/groups",
  `${root}/users?$skiptoken=x`,
  `${root}/groups/${firstId}/members`,
  "https://graph.microsoft.com/beta/groups",
  `${root}/groups#fragment`,
  "//example.test/groups",
  "not a URL",
  ` ${root}/groups`,
  `${root}/groups\n?$skiptoken=x`,
  `${root}/groups\\?$skiptoken=x`
]) {
  test(`unsafe continuation is never requested: ${JSON.stringify(next)}`, async () => {
    const client = mockPages({ value: [group()], "@odata.nextLink": next });
    await assert.rejects(resolveGroup({}, address, client), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE/);
    assert.equal(client.calls.length, 1);
  });
}

test("looping continuations and repeated IDs reject rather than return a candidate", async () => {
  const next = `${root}/groups?$skiptoken=x`;
  const client = mockPages(
    { value: [group()], "@odata.nextLink": next },
    { value: [group({ id: secondId, mail: "other@example.test", proxyAddresses: [] })], "@odata.nextLink": next }
  );
  await assert.rejects(resolveGroup({}, address, client), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE:.*loops/);
  assert.equal(client.calls.length, 2);
  await assert.rejects(resolveGroup({}, address, mockPages(
    { value: [group()], "@odata.nextLink": next }, { value: [group()] }
  )), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE:.*repeated/);
});

test("page limit rejects even a matching group on the first page", async () => {
  const pages = Array.from({ length: 10 }, (_, index) => ({
    value: [group({ id: `synthetic-${index}` })],
    "@odata.nextLink": `${root}/groups?$skiptoken=${index}`
  }));
  const client = mockPages(...pages);
  await assert.rejects(resolveGroup({}, address, client), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE:.*page limit/);
  assert.equal(client.calls.length, 10);
});

for (const page of [
  undefined, null, [], {}, { value: null }, { value: {} },
  { value: [null] }, { value: [{}] }, { value: [group({ id: "" })] },
  { value: [group({ "@odata.type": "#microsoft.graph.user" })] },
  { value: [group({ groupTypes: "Unified" })] },
  { value: [group({ groupTypes: [1] })] },
  { value: [group({ mailEnabled: "true" })] },
  { value: [group({ proxyAddresses: "smtp:alias@example.test" })] },
  { value: [group({ "proxyAddresses@odata.nextLink": "/groups" })] },
  { value: [group()], "@odata.nextLink": "" },
  { value: [group()], "@odata.nextLink": null },
  { value: [group()], "@odata.nextLink": 1 },
  { value: [group()], "@odata.deltaLink": "/groups/delta" },
  { value: [group()], "@odata.count": 2 },
  { value: [group()], "@odata.count": 0 },
  { value: [group()], "@odata.count": "1" },
  { value: [group()], "@odata.count": 251 },
  { value: Array.from({ length: 26 }, (_, index) => group({ id: `synthetic-${index}` })) },
  { value: [group(), group({ id: secondId, mail: undefined, proxyAddresses: undefined })] }
]) {
  test(`malformed or truncated page rejects: ${JSON.stringify(page)}`, async () => {
    await assert.rejects(resolveGroup({}, address, mockPages(page)), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE/);
  });
}

test("inconsistent count metadata between pages fails closed", async () => {
  await assert.rejects(resolveGroup({}, address, mockPages(
    { value: [group()], "@odata.count": 2, "@odata.nextLink": "/groups?$skiptoken=x" },
    { value: [group({ id: secondId })], "@odata.count": 3 }
  )), /GROUP_DIRECTORY_LOOKUP_INCOMPLETE/);
});

test("elapsed overall time rejects a response without issuing a follow-up", async (t) => {
  let now = 1_000;
  t.mock.method(Date, "now", () => now);
  let calls = 0;
  await assert.rejects(resolveGroup({}, address, {
    async get(_url, options) {
      calls += 1;
      assert.equal(options.totalTimeoutMs, 15_000);
      now += 15_001;
      return { value: [group()], "@odata.nextLink": "/groups?$skiptoken=x" };
    }
  }), /GROUP_DIRECTORY_NETWORK_BLOCKED/);
  assert.equal(calls, 1);
});

test("a stalled client is bounded by the outer timeout", async (t) => {
  const timers = [];
  t.mock.method(globalThis, "setTimeout", (callback, delay) => {
    assert.ok(delay > 0 && delay <= 15_000);
    const timer = { callback };
    timers.push(timer);
    queueMicrotask(callback);
    return timer;
  });
  t.mock.method(globalThis, "clearTimeout", (timer) => { timer.cleared = true; });
  await assert.rejects(resolveGroup({}, firstId, { get() { return new Promise(() => {}); } }), /GROUP_DIRECTORY_NETWORK_BLOCKED/);
  assert.equal(timers.length, 1);
  assert.equal(timers[0].cleared, true);
});

test("invalid identifiers fail before any client call", async () => {
  for (const input of [undefined, "", " ", "a\nb", "a\u007fb", "a".repeat(65), "Display Name", "smtp:team@example.test", "a@@example.test"]) {
    const client = mockPages();
    await assert.rejects(resolveGroup({}, input, client), /GROUP_INVALID_IDENTIFIER/);
    assert.equal(client.calls.length, 0);
  }
});
