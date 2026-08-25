import type { Env } from "../src/types.js";

declare module "cloudflare:test" {
	interface ProvidedEnv extends Env {}
}

// Vite's `?raw` import suffix — used by wrangler-env-disjointness.spec.ts to read
// wrangler.jsonc's text at bundle time without needing filesystem access at runtime.
declare module "*?raw" {
	const content: string;
	export default content;
}
