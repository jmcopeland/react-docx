import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { test } from "node:test";

import {
  assertExpectedFontFingerprint,
  fingerprintFontRecords,
} from "../../scripts/word-oracle/font-fingerprint.mjs";

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

test("font fingerprints bind normalized face metadata to actual font bytes", async () => {
  const fontBytes = new Map([
    ["/fonts/a.ttf", Buffer.from("font-a")],
    ["/fonts/b.ttf", Buffer.from("font-b")],
  ]);
  const records = [
    {
      postScriptName: "Fixture-Bold",
      version: "2",
      filePath: "/fonts/b.ttf",
    },
    {
      postScriptName: "Fixture-Regular",
      version: "1",
      filePath: "/fonts/a.ttf",
    },
  ];
  const result = await fingerprintFontRecords(records, {
    readFile: async (filePath) => fontBytes.get(filePath),
  });
  const expectedInventory = [
    `Fixture-Bold\t2\t${sha256(fontBytes.get("/fonts/b.ttf"))}`,
    `Fixture-Regular\t1\t${sha256(fontBytes.get("/fonts/a.ttf"))}`,
    "",
  ].join("\n");
  assert.equal(result.normalizedInventory, expectedInventory);
  assert.equal(result.fingerprintSha256, sha256(expectedInventory));
  assert.equal(result.fontFaceCount, 2);
  assert.equal(result.fontFileCount, 2);

  assert.doesNotThrow(() =>
    assertExpectedFontFingerprint(
      result.fingerprintSha256,
      result.fingerprintSha256
    )
  );
  assert.throws(
    () => assertExpectedFontFingerprint(result.fingerprintSha256, "f".repeat(64)),
    /Installed font fingerprint mismatch/
  );
});

test("font fingerprints are stable across fontconfig output order", async () => {
  const records = [
    {
      postScriptName: "Fixture-Regular",
      version: "1",
      filePath: "/fonts/a.ttf",
    },
    {
      postScriptName: "Fixture-Bold",
      version: "2",
      filePath: "/fonts/b.ttf",
    },
  ];
  const readFile = async (filePath) => Buffer.from(filePath);
  const forward = await fingerprintFontRecords(records, { readFile });
  const reverse = await fingerprintFontRecords(records.toReversed(), {
    readFile,
  });
  assert.equal(forward.fingerprintSha256, reverse.fingerprintSha256);
});
