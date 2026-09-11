#!/usr/bin/env node
/**
 * Live MCP handshake check. Start the Flask app first, then run:
 * CONNECTOR_BASE_URL=http://127.0.0.1:5011 CONNECTOR_API_TOKEN=... node tests/connector_api_mcp_handshake.mjs
 *
 * The SDK is intentionally resolved from nyharness during this shared
 * development stage; production callers provide their own MCP SDK.
 */
import { spawn } from "node:child_process";
import { readdirSync } from "node:fs";
import { createServer } from "node:net";
import path from "node:path";
import { pathToFileURL } from "node:url";
import process from "node:process";

const token = process.env.CONNECTOR_API_TOKEN;
if (!token) throw new Error("CONNECTOR_API_TOKEN 必填");

const nanyaRoot = process.env.NANYA_HARNESS_ROOT || path.resolve(process.cwd(), "../../../ny-harness/nyharness");
const pnpmDir = path.join(nanyaRoot, "node_modules", ".pnpm");
const sdkEntry = readdirSync(pnpmDir).find((item) => item.startsWith("@modelcontextprotocol+sdk@"));
if (!sdkEntry) throw new Error("nyharness 中未找到 @modelcontextprotocol/sdk；请先安装其依赖或设置 NANYA_HARNESS_ROOT");
const sdkRoot = path.join(pnpmDir, sdkEntry, "node_modules", "@modelcontextprotocol", "sdk");
const { Client } = await import(pathToFileURL(path.join(sdkRoot, "dist", "esm", "client", "index.js")).href);
const { StreamableHTTPClientTransport } = await import(pathToFileURL(path.join(sdkRoot, "dist", "esm", "client", "streamableHttp.js")).href);

let server = null;
let baseUrl = process.env.CONNECTOR_BASE_URL;
if (!baseUrl) {
  const port = await new Promise((resolve, reject) => {
    const probe = createServer();
    probe.once("error", reject);
    probe.listen(0, "127.0.0.1", () => {
      const selected = probe.address().port;
      probe.close((error) => error ? reject(error) : resolve(selected));
    });
  });
  const python = process.env.CONNECTOR_PYTHON || "./.venv/bin/python";
  server = spawn(python, ["tests/connector_api_test_server.py"], {
    cwd: process.cwd(), env: { ...process.env, CONNECTOR_API_TOKEN: token, CONNECTOR_TEST_PORT: String(port) },
    stdio: ["ignore", "pipe", "pipe"],
  });
  const ready = await new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error("connector test server did not start")), 15_000);
    server.stdout.on("data", (chunk) => {
      if (chunk.toString().includes("READY")) { clearTimeout(timeout); resolve(true); }
    });
    server.on("error", (error) => { clearTimeout(timeout); reject(error); });
    server.stderr.on("data", (chunk) => process.stderr.write(chunk));
  });
  if (!ready) throw new Error("connector test server unavailable");
  baseUrl = `http://127.0.0.1:${port}`;
}

const transport = new StreamableHTTPClientTransport(new URL(`${baseUrl.replace(/\/$/, "")}/api/connector/v1/mcp`), {
  requestInit: { headers: { Authorization: `Bearer ${token}`, "X-Nanya-Employee-Id": "23582" } },
});
const client = new Client({ name: "nanya-marketing-connector-contract", version: "1.0.0" }, { capabilities: {} });
try {
  await client.connect(transport);
  const { tools } = await client.listTools();
  const names = tools.map((item) => item.name).sort();
  const expected = ["marketing.job.get", "marketing.mail.sync_orders", "marketing.material.query", "marketing.order.get_case", "marketing.order.list_cases", "marketing.order.prepare_entry", "marketing.order.reply_draft", "marketing.order.submit_entry"];
  if (JSON.stringify(names) !== JSON.stringify(expected)) throw new Error(`unexpected tools: ${names.join(",")}`);
  console.log(JSON.stringify({ ok: true, transport: "streamable-http", tools: names }, null, 2));
} finally {
  await transport.close().catch(() => {});
  server?.kill("SIGTERM");
}
