import { describe, expect, it } from "vitest";
// Vite's `?raw` suffix inlines the file's text content at bundle time, so this
// needs no runtime filesystem access — vitest-pool-workers' sandboxed workerd
// runtime doesn't give test code a real host filesystem (an earlier version of
// this test used `node:fs` `readFileSync` directly and failed with "no such file
// or directory" even though the file plainly exists, for exactly that reason).
import wranglerJsoncRaw from "../wrangler.jsonc?raw";

/**
 * Guards the worst failure mode in this two-tenant build: the future-comms and
 * total-telco environments in wrangler.jsonc silently sharing a resource. Wrangler
 * named environments do NOT inherit `vars`, `kv_namespaces`, `unsafe.bindings` or
 * `define` from the top level — every one of those must be redeclared in full in
 * both env blocks, and if one is copy-pasted without changing the values, the
 * Worker deploys cleanly and just serves one business's data (or name) under both.
 *
 * Parsed with a small string-aware comment stripper (not the `typescript` package's
 * JSONC parser — that needs `node:os`, which fails to load in the sandboxed workerd
 * runtime these tests run in). The state machine walks the raw text so `//` inside a
 * quoted string (e.g. `"https://..."` in WORKER_URL) is never mistaken for a comment.
 */

function stripJsonComments(input: string): string {
	let out = "";
	let inString = false;
	let inLineComment = false;
	let inBlockComment = false;
	let escaped = false;

	for (let i = 0; i < input.length; i++) {
		const char = input[i];
		const next = input[i + 1];

		if (inLineComment) {
			if (char === "\n") {
				inLineComment = false;
				out += char;
			}
			continue;
		}
		if (inBlockComment) {
			if (char === "*" && next === "/") {
				inBlockComment = false;
				i++;
			}
			continue;
		}
		if (inString) {
			out += char;
			if (escaped) escaped = false;
			else if (char === "\\") escaped = true;
			else if (char === '"') inString = false;
			continue;
		}
		if (char === '"') {
			inString = true;
			out += char;
			continue;
		}
		if (char === "/" && next === "/") {
			inLineComment = true;
			i++;
			continue;
		}
		if (char === "/" && next === "*") {
			inBlockComment = true;
			i++;
			continue;
		}
		out += char;
	}
	return out;
}

function loadWranglerConfig(): any {
	return JSON.parse(stripJsonComments(wranglerJsoncRaw));
}

const config = loadWranglerConfig();
const envs = config.env ?? {};
const futureComms = envs["future-comms"];
const totalTelco = envs["total-telco"];

describe("wrangler.jsonc future-comms vs total-telco disjointness", () => {
	it("both environments are actually defined", () => {
		expect(futureComms, "env.future-comms missing from wrangler.jsonc").toBeDefined();
		expect(totalTelco, "env.total-telco missing from wrangler.jsonc").toBeDefined();
	});

	it("has different worker names", () => {
		expect(futureComms.name).toBeTruthy();
		expect(totalTelco.name).toBeTruthy();
		expect(futureComms.name).not.toBe(totalTelco.name);
	});

	it("has different KV namespace IDs", () => {
		const fcId = futureComms.kv_namespaces?.[0]?.id;
		const ttId = totalTelco.kv_namespaces?.[0]?.id;
		expect(fcId, "future-comms kv_namespaces[0].id missing").toBeTruthy();
		expect(ttId, "total-telco kv_namespaces[0].id missing").toBeTruthy();
		// This is the failure that matters most: a shared KV namespace means one
		// business's mailbox tokens serve the other worker's requests, with no error.
		expect(fcId).not.toBe(ttId);
	});

	it("has different rate-limit namespace IDs for every matching binding", () => {
		const toMap = (env: any) => {
			const map = new Map<string, string>();
			for (const binding of env.unsafe?.bindings ?? []) {
				map.set(binding.name, binding.namespace_id);
			}
			return map;
		};
		const fcBindings = toMap(futureComms);
		const ttBindings = toMap(totalTelco);

		const bindingNames = new Set([...fcBindings.keys(), ...ttBindings.keys()]);
		expect(bindingNames.size).toBeGreaterThan(0);

		for (const name of bindingNames) {
			const fcNamespaceId = fcBindings.get(name);
			const ttNamespaceId = ttBindings.get(name);
			expect(fcNamespaceId, `future-comms missing rate-limit binding ${name}`).toBeTruthy();
			expect(ttNamespaceId, `total-telco missing rate-limit binding ${name}`).toBeTruthy();
			expect(fcNamespaceId).not.toBe(ttNamespaceId);
		}
	});

	it("has a SERVER_NAME_CONST define present on each environment, and different between them", () => {
		const fcServerName = futureComms.define?.SERVER_NAME_CONST;
		const ttServerName = totalTelco.define?.SERVER_NAME_CONST;
		// Missing define here is the silent-500s trap: the build succeeds with only
		// a warning, and every request throws ReferenceError at module load.
		expect(fcServerName, "future-comms define.SERVER_NAME_CONST missing").toBeTruthy();
		expect(ttServerName, "total-telco define.SERVER_NAME_CONST missing").toBeTruthy();
		expect(fcServerName).not.toBe(ttServerName);
	});
});
