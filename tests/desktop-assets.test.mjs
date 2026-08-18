import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("desktop package contains every script and stylesheet referenced by HTML", async () => {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("desktop-assets-test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  const response = await worker.fetch(
    new Request("http://127.0.0.1/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
  assert.equal(response.status, 200);
  const html = await response.text();
  const assetUrls = [...html.matchAll(/<(?:script|link)\b[^>]*(?:src|href)="(\/assets\/[^"]+)"/gi)]
    .map((match) => match[1]);
  assert.ok(assetUrls.length >= 2, "expected built JavaScript and CSS assets");
  for (const assetUrl of new Set(assetUrls)) {
    const assetPath = path.join(projectRoot, "dist", "client", assetUrl);
    await access(assetPath);
    assert.ok((await readFile(assetPath)).byteLength > 0, `${assetUrl} must not be empty`);
  }
  assert.doesNotMatch(html, /\/workspace\/.*\.woff2/i, "HTML must not contain build-machine font paths");
});
