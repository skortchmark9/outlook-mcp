# Onboarding

This is a fork of [ryaker/outlook-mcp](https://github.com/ryaker/outlook-mcp). The upstream
[README](./README.md) covers Azure app registration and general setup — **read that first**.

This document covers only what differs in this fork, plus the setup gotchas that aren't
obvious from the upstream docs.

## Quick start

```bash
git clone git@github.com:skortchmark9/outlook-mcp.git
cd outlook-mcp
npm install
cp .env.example .env     # fill in your own Azure credentials
npm run auth-server      # then use the `authenticate` tool in Claude
```

You need your **own** Azure app registration (or to be added to an existing one) — credentials
are never committed. See the README's "Azure App Registration & Configuration" section.

## Gotcha: two different env var prefixes

The codebase reads Azure credentials under **two** naming schemes, and different entry points
read different ones. This trips people up.

| Variable | Read by | Notes |
|---|---|---|
| `MS_CLIENT_ID` / `MS_CLIENT_SECRET` | `outlook-auth-server.js:49-50` | **Required.** The auth server reads *only* these. |
| `MS_TENANT_ID` | `outlook-auth-server.js:51`, `auth/token-storage.js:8` | Defaults to `common`. Set to your tenant GUID for single-tenant apps. |
| `OUTLOOK_CLIENT_ID` / `OUTLOOK_CLIENT_SECRET` | `config.js:20-21` | Alternative naming, intended for Claude Desktop config. |
| either prefix | `auth/token-storage.js:12-13` | Token refresh accepts `MS_*` **or** `OUTLOOK_*`. |

**Use `MS_*` in `.env`** (what `.env.example` ships). That is what the auth server needs, and
token refresh accepts it too.

One harmless quirk: with only `MS_*` set, the URL from the `authenticate` tool ends in an empty
`client_id=` (`auth/tools.js:42` reads `config.AUTH_CONFIG.clientId`, which is `OUTLOOK_*`-only).
This does **not** break anything — the auth server ignores that query param and uses its own
`MS_CLIENT_ID` when building the real Microsoft URL (`outlook-auth-server.js:230`). Don't chase it.

## Fork-specific: multiple mailboxes (`OUTLOOK_TOKEN_PATH`)

Upstream hardcodes the token file to `~/.outlook-mcp-tokens.json`, so one machine = one mailbox.
This fork makes that path configurable, so you can run several instances against different
accounts (e.g. a work mailbox and a personal one) side by side.

**Authenticate each mailbox into its own token file:**

```bash
node outlook-auth-server.js --token-path=~/.outlook-mcp-tokens-work.json
# complete the browser flow, then repeat for the next account:
node outlook-auth-server.js --token-path=~/.outlook-mcp-tokens-personal.json
```

Use `node outlook-auth-server.js --token-path=...` rather than `npm run auth-server` — passing
args through npm needs an extra `--` and is easy to get wrong.

**Point each MCP server entry at the matching token file** via the `OUTLOOK_TOKEN_PATH` env var:

```json
{
  "mcpServers": {
    "outlook-work": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": { "OUTLOOK_TOKEN_PATH": "/Users/you/.outlook-mcp-tokens-work.json" }
    },
    "outlook-personal": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": { "OUTLOOK_TOKEN_PATH": "/Users/you/.outlook-mcp-tokens-personal.json" }
    }
  }
}
```

Resolution order is CLI arg → `OUTLOOK_TOKEN_PATH` → the `~/.outlook-mcp-tokens.json` default
(`outlook-auth-server.js:14-19`), so omitting it everywhere reproduces upstream's behavior.

Note that `.mcp.json` is gitignored in this fork — your local server config stays out of git.

## Fork-specific: `Mail.ReadWrite` scope

The auth server now requests `Mail.ReadWrite` (`outlook-auth-server.js:58`), which is needed for
drafts, replies, moves, and deletes.

**If you have tokens from before this change, you must re-authenticate.** Scopes are baked into
the token at issue time; a refresh will not add them. The symptom is `ErrorAccessDenied` on
write operations while reads keep working. Delete your token file and run the auth flow again.

Make sure `Mail.ReadWrite` is also granted on the Azure app registration itself, under
API permissions → Microsoft Graph → Delegated.

## Fork-specific: `reply-email` tool

New tool (`email/reply.js`) that replies **in-thread** via Graph's `/reply` and `/replyAll`
endpoints, preserving `In-Reply-To`, `References`, and `conversationId`. Using `send-email` for
a reply starts a new thread instead — use this tool when replying.

| Param | Required | Notes |
|---|---|---|
| `id` | yes | ID of the message being replied to |
| `body` | yes | Your text; Graph quotes the original beneath it |
| `replyAll` | no | Default `false` |
| `to` / `cc` | no | Comma-separated overrides; defaults to the original sender |
| `isHtml` | no | Auto-detected from the body if omitted |

## Fork-specific: `search-emails` behavior changes

Three upstream bugs are fixed here, and the third changes results you may be used to:

1. **KQL quoting** (`email/search.js`, `buildKqlTerm`). Upstream wrapped values in nested double
   quotes, which Graph rejects with `400 Syntax error: character '"' is not valid`. Multi-word
   values are now parenthesized — `subject:(quarterly report)` — because without the parens only
   the first word is scoped to the field and the rest matches *any* field as free text.

2. **Folder aliases** (`email/folder-utils.js`). `"Sent Items"`, `"sentitems"`, `"sent"` etc. all
   normalize to the same well-known folder. Upstream silently fell back to the **inbox** when a
   folder name didn't resolve, returning wrong-folder results that looked like a broken search.
   `drafts`/`draft`, `deleted`/`deleteditems`/`trash`, and `junk`/`junkemail`/`spam` are aliased too.

3. **Honest zero results** (`email/search.js`). When you supply search terms and nothing matches,
   you now get an empty result. Upstream fell back to dumping recent emails and labelled them
   "matching your search criteria" — pure false positives. The recent-emails fallback only
   applies when no search terms were given at all.

Also: `email/draft.js` detects HTML from any common tag (`<p>`, `<div>`, `<br>`, `<table>`, …)
rather than requiring a literal `<html>`, so drafts written as HTML fragments no longer send as
plain text showing raw markup.

## Staying in sync with upstream

```bash
git remote add upstream https://github.com/ryaker/outlook-mcp.git   # once
git fetch upstream
git merge upstream/main    # or: git rebase upstream/main
```

Our changes are additive and touch few files, so upstream merges are usually clean. The most
likely conflict points are `email/search.js` and `outlook-auth-server.js`.

## Troubleshooting

| Symptom | Cause |
|---|---|
| `AADSTS7000215: Invalid client secret` | You used the Secret **ID** instead of the secret **VALUE** from Azure. |
| Auth server shows "Configuration Error" | `MS_CLIENT_ID` / `MS_CLIENT_SECRET` aren't set — check `.env` is in the repo root. |
| `EADDRINUSE` on port 3333 | `npx kill-port 3333` |
| `ErrorAccessDenied` on write ops, reads fine | Token predates the `Mail.ReadWrite` scope — re-authenticate. |
| Wrong mailbox's mail returned | `OUTLOOK_TOKEN_PATH` isn't set on that MCP server entry, so it fell back to the default token file. |
| Search returns unrelated emails | Should be fixed here; if it recurs, check whether the folder name resolved (the server logs the folder path it used). |
