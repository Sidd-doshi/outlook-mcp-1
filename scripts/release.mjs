#!/usr/bin/env node
// Cut a release: bump, changelog, commit, tag — in that order, atomically.
//
// ── Why this exists ───────────────────────────────────────────────────────────
// Versions used to be bumped whenever it occurred to someone. 2.4.0 was bumped
// at the *start* of its work and then collected two months of fixes under the
// same number; 2.5.0 and 2.6.0 were bumped at the *end*. Mixing the two makes
// "which version is this commit in?" unanswerable from the tree: the transcript
// fix in b450955 sits in a commit whose package.json still reads 2.4.0, even
// though it shipped in 2.5.0. Only `git tag --contains` can tell you, and that
// only works once someone has tagged correctly.
//
// The rule this enforces: the version bump IS the release commit, and the tag
// points at it. So `git show <tag>:package.json` always agrees with the tag,
// and every commit belongs to the next tag that contains it.
//
// Usage:  npm run release 2.7.0
//
// Deliberately does not push. On this repo a push to main is a deploy — see
// .github/workflows/deploy.yml — so the last step stays a human decision.

import { execFileSync } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";

const run = (cmd, args) => execFileSync(cmd, args, { encoding: "utf8" }).trim();
const git = (...args) => run("git", args);

function fail(message, hint) {
	console.error(`\n  ✗ ${message}`);
	if (hint) console.error(`    ${hint}`);
	console.error("");
	process.exit(1);
}

// ── 1. Argument ───────────────────────────────────────────────────────────────

const version = process.argv[2];
if (!version) fail("No version given.", "Usage: npm run release 2.7.0");
if (!/^\d+\.\d+\.\d+$/.test(version)) {
	// SERVER_VERSION is displayed verbatim by MCP clients, and version.spec.ts
	// rejects anything that isn't a plain triple.
	fail(`"${version}" is not a plain semver triple.`, "No v prefix, no -rc suffix — e.g. 2.7.0");
}

const pkgPath = new URL("../package.json", import.meta.url);
const pkg = JSON.parse(readFileSync(pkgPath, "utf8"));
const current = pkg.version;

const asNumbers = (v) => v.split(".").map(Number);
const [a, b, c] = asNumbers(version);
const [x, y, z] = asNumbers(current);
if (a * 1e6 + b * 1e3 + c <= x * 1e6 + y * 1e3 + z) {
	fail(`${version} is not ahead of the current ${current}.`);
}

// ── 2. Preconditions ──────────────────────────────────────────────────────────

const branch = git("rev-parse", "--abbrev-ref", "HEAD");
if (branch !== "main") {
	fail(`On branch "${branch}", not main.`, "Releases are cut from main.");
}

if (git("status", "--porcelain")) {
	fail("Working tree is dirty.", "Commit or stash first — the release commit should contain only the bump.");
}

git("fetch", "origin", "main");
if (git("rev-list", "--count", "origin/main..HEAD") !== "0") {
	fail("Local main has unpushed commits.", "Push them first so the release commit sits on top of what others see.");
}
if (git("rev-list", "--count", "HEAD..origin/main") !== "0") {
	fail("Local main is behind origin.", "Pull first.");
}

if (git("tag", "--list", `v${version}`)) fail(`Tag v${version} already exists.`);

// ── 3. Changelog ──────────────────────────────────────────────────────────────
// An "## Unreleased" heading is required. Writing the entry while the work is
// fresh is the whole point — everything before 2.5.0 had to be reconstructed
// from commit subjects months later, and two fixes ended up filed under the
// wrong release until the dates were checked.

const changelogPath = new URL("../CHANGELOG.md", import.meta.url);
let changelog = readFileSync(changelogPath, "utf8");

if (!changelog.includes("## Unreleased")) {
	fail(
		"CHANGELOG.md has no `## Unreleased` section.",
		"Add one describing this release before cutting it.",
	);
}

const unreleasedBody = changelog.split("## Unreleased")[1].split(/\n## /)[0].trim();
if (!unreleasedBody) {
	fail("The `## Unreleased` section is empty.", "Say what changed — someone updating needs to know.");
}

const today = new Date().toISOString().slice(0, 10);
changelog = changelog.replace("## Unreleased", `## ${version} — ${today}`);

// ── 4. Verify before writing anything ─────────────────────────────────────────

console.log(`\n  Releasing ${current} → ${version}\n`);
console.log("  Running typecheck and tests...");
try {
	execFileSync("npm", ["run", "typecheck"], { stdio: "pipe" });
	execFileSync("npx", ["vitest", "run"], { stdio: "pipe" });
} catch (e) {
	fail("Tests or typecheck failed — nothing was changed.", "Fix them, then re-run.");
}
console.log("  ✓ green\n");

// ── 5. Bump, commit, tag ──────────────────────────────────────────────────────

pkg.version = version;
writeFileSync(pkgPath, `${JSON.stringify(pkg, null, "\t")}\n`);

const indexPath = new URL("../src/index.ts", import.meta.url);
const index = readFileSync(indexPath, "utf8");
const bumped = index.replace(
	/export const SERVER_VERSION = "\d+\.\d+\.\d+";/,
	`export const SERVER_VERSION = "${version}";`,
);
if (bumped === index) {
	fail("Could not find SERVER_VERSION in src/index.ts.", "The declaration may have been reshaped.");
}
writeFileSync(indexPath, bumped);
writeFileSync(changelogPath, changelog);

git("add", "package.json", "src/index.ts", "CHANGELOG.md");
git("commit", "-m", `Release ${version}\n\n${unreleasedBody}\n\nCo-Authored-By: Claude Opus 5 <noreply@anthropic.com>`);
git("tag", "-a", `v${version}`, "-m", `v${version}`);

console.log(`  ✓ Committed and tagged v${version}\n`);
console.log("  Push when ready — note that pushing main deploys:\n");
console.log(`    git push origin main v${version}\n`);
