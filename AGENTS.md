# AI Assistant Instructions: D365 PCF Document Manager

Read this file first. It is the single source of instructions for AI coding agents (Claude Code reads it via
`CLAUDE.md`; Cursor, Copilot and Windsurf read `AGENTS.md` directly). Details live in `docs/`:

| File | Read it when |
|---|---|
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Changing folders, metadata, auth, or anything crossing the PCF ↔ Function ↔ Graph boundary |
| [docs/PCF_STANDARDS.md](docs/PCF_STANDARDS.md) | Touching anything under `pcf/` |
| [docs/API_SPECS.md](docs/API_SPECS.md) | Calling or changing an endpoint (ours or Microsoft Graph) |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | A build, packaging, auth or runtime error |
| [docs/PREREQUISITES.md](docs/PREREQUISITES.md) | What must exist in SharePoint / D365 / Entra / Azure (columns, lookups, settings, roles) |
| [docs/SETUP.md](docs/SETUP.md) | Step-by-step setup and smoke test |
| [docs/CONFIGURATION.md](docs/CONFIGURATION.md) | Categories, per-table config (contact lookup column), environment variables, `dataverse/provision.mjs` |

## Project overview

A Power Apps Component Framework (PCF) **virtual** control (`EnterpriseDocManagement.CategorizedDocManager`)
on Dynamics 365 model-driven forms (Opportunity, Advisor Onboarding). It keeps every document for an advisor
in **one SharePoint folder per Contact**, tags each file with category and source-record metadata, and shows
"This Record's Documents" / "All Contact Documents" tabs with an inline preview pane.

```
D365 form ─► PCF control ──(user token via MSAL SSO)──► Azure Function /api/documents/*
                                                          │ verifies the JWT itself
                                                          └─(app-only token, Sites.Selected)─► Microsoft Graph ─► SharePoint
```

The **browser never calls Microsoft Graph or SharePoint directly**. All Graph access is in `backend/`.

## Tech stack (exact; do not "upgrade" without being asked)

| Area | Choice | Why |
|---|---|---|
| Control | PCF virtual control (`control-type="virtual"`), field template, bound to a Contact lookup | Uses the platform's React/Fluent instead of bundling them |
| UI | **React 16.14.0** + **Fluent UI v9 9.46.2** as PCF *platform libraries* | React 18 platform libraries are still behind a preview flag in `pcf-scripts` (see TROUBLESHOOTING) |
| Icons | Local inline SVGs in `pcf/.../components/Icons.tsx` (copied from `@fluentui/react-icons`, MIT) | The icons package bundles a second Griffel that fails to build against React 16 |
| Auth (client) | `@azure/msal-browser` 3.x, `ssoSilent` with a popup fallback | — |
| Language | TypeScript 4.9 (pcf), TypeScript 5.x (backend) | — |
| Backend | Azure Functions v4, Node 20, `@azure/identity`, `jose` | — |
| Data | Microsoft Graph v1.0 (drives, list items, preview, upload sessions); Dataverse via `context.webAPI` / `context.utils` only | — |
| Tooling | `pcf-scripts` (webpack), PAC CLI, `dotnet build`/MSBuild for the solution | — |

## Core rules

1. **No heavy packages in `pcf/`.** No moment/dayjs (use `Intl`/`toLocaleDateString`), no lodash, no Fluent v8
   (`@fluentui/react`), no `@fluentui/react-icons` (add icons to `Icons.tsx` instead). `react`, `react-dom` and
   `@fluentui/react-components` stay in `devDependencies` and pinned to the platform-library versions in
   `ControlManifest.Input.xml`. Check: `npm run build:prod` and keep `bundle.js` around 330 KB (about 284 KB is MSAL).
2. **Stable props from `updateView`.** `index.ts` caches `host`, `categories` and the `DocumentApi` and replaces
   them only when their inputs change. Never build new objects/arrays/functions per `updateView` call, never fetch
   from `updateView`, and never use `notifyOutputChanged` to force a re-render (the control has no outputs).
3. **Memoize.** Filtering and sorting go in `React.useMemo`. Handlers passed to cards go in `React.useCallback`.
   `DocumentCard` and `DocumentManager` are `React.memo`. Keep it that way.
4. **Every network call has a timeout, and retries must not duplicate files.** Browser: `fetchWithTimeout` /
   `withTimeout` (`pcf/.../services/http.ts`), never bare `fetch`. File uploads use `xhrUpload` (same file), which is
   the only allowed upload transport: it gives byte progress and the same timeout contract. Don't switch uploads
   back to `fetch` and don't add other raw `XMLHttpRequest` calls. Backend: Graph calls go through `graph()`
   (`graph.ts`) and upload-session traffic through `uploadToSession()` (`uploadSession.ts`), which also times out the
   cancel request. Only idempotent requests are retried after ambiguous failures (`retryDecision`). Timeouts surface
   as a message with Retry, never an endless spinner. Timeout values are named constants; see PCF_STANDARDS §4 and
   API_SPECS Part B.
5. **Entity context comes from `context.mode.contextInfo`** (`entityTypeName`, `entityId`). `context.page` is not used.
   **Configuration is data, not code:** categories, folder paths, which tables host the control, each table's
   **Contact lookup column** and upload limits live in Dataverse (`ce_documentcategory`, `ce_documentmanagerconfig`).
   The connection settings live in the environment variables `ce_DocManager*`. Never hard-code a category, table or
   connection value. See docs/CONFIGURATION.md.
6. **No DOM access outside the control.** React state and Fluent components only. Don't touch the D365 page DOM
   or `Xrm`. (Exceptions: the temporary `<a>` used to start a download, and submitting the control's own
   React-rendered preview `<form>`.)
7. **Destructive actions need confirmation on both sides.** Delete goes through `DeleteDialog` (checkbox
   acknowledgement). `DocumentApi.remove(id, name)` sends `X-Confirm-Delete`, and the backend refuses folders and
   name mismatches. Never add a delete or bulk-delete path that skips either check.
8. **Server is the source of truth** for enforcing configuration (`backend/src/lib/tableConfig.ts`: table configured,
   category allowed, folder path, size limit), file safety
   (`fileRules.ts`: blocked types, executable content, content/extension match), file-name sanitising, input
   validation and authorization. Client checks are for UX only.
9. **Never trust `X-MS-CLIENT-PRINCIPAL`.** Callers are authenticated by verifying the bearer JWT in
   `backend/src/lib/auth.ts`. **Every document endpoint authorizes per record** through `accessFor(caller)`
   (`backend/src/lib/access.ts`, Dataverse On-Behalf-Of as the user) **before** touching SharePoint: read the
   contact, filter or verify source records, write access for changes. Never add a bypass, never use the app-only
   `dataverse()` for user authorization decisions, and keep `access.test.ts` covering any new rule.
10. **Secrets never go in the PCF control or the repo.** Control properties hold only public IDs and URLs.
    `GRAPH_CLIENT_SECRET` lives in Function App settings / Key Vault. `local.settings.json` is gitignored.
11. **Files are untrusted input.** Never store a file before `assertFileAllowed()` passes. Never hand the browser a
    URL that hasn't passed `assertTrustedUrl()`. Never weaken the post-upload size check or the metadata rollback.

## Commands

| Task | Command (run from) |
|---|---|
| Install | `npm ci` (`pcf/`, `backend/`) |
| Build control (dev) | `npm run build` (`pcf/`) |
| Build control (prod, check size) | `npm run build:prod` (`pcf/`) |
| Test harness | `npm start watch` (`pcf/`) — harness has no real API; see PCF_STANDARDS §6 |
| Backend build + unit tests | `npm test` (`backend/`) |
| Control unit tests (helpers in `utils.ts`) | `npm test` (`pcf/`) |
| Dataverse provisioning script tests | `node --test dataverse/provision.test.mjs` (repo root) |
| Provision / configure Dataverse | `node dataverse/provision.mjs all --org <url> [--dry-run\|--offline]`, then `set-env`, `app-user` |
| Run backend locally | `cp local.settings.sample.json local.settings.json && npm start` (`backend/`, needs Functions Core Tools) |
| Package solution | `pac solution init …` → `pac solution add-reference --path ../pcf` → `dotnet build -c Release` (or `msbuild /t:build /restore /p:Configuration=Release`) |
| Push to a dev org | `pac pcf push --publisher-prefix ce` (`pcf/`) |

Linting is skipped during build (`pcfconfig.json` → `skipBuildLinting`). Run `npx tsc --noEmit` in `pcf/` as the
type check.

## Definition of done for a change

- `npm run build`, `npm run build:prod` and `npm test` succeed in `pcf/`. `npm test` passes in `backend/`. Anything touching
  uploads, retries or file safety keeps `uploadSession.test.ts` and `safety.test.ts` green and extends them.
- New backend logic that can be tested without Graph (parsing, naming, validation, auth rejection) has a test in
  `backend/test/`.
- If you changed a contract, update `docs/API_SPECS.md`. If you changed a folder, column or flow, update
  `docs/ARCHITECTURE.md`. If you changed a setting, update `docs/SETUP.md`. If you changed a SharePoint column, Dataverse requirement or
  permission, update `docs/PREREQUISITES.md` and `sharepoint/Test-ContactsLibrary.ps1`.
- Adding a category or a hosting table is **configuration, not code** (docs/CONFIGURATION.md §3). If you change the
  configuration **schema**, update `dataverse/provision.mjs` (+ its tests), `dataverse/seed.json`,
  `backend/src/lib/tableConfig.ts`, `pcf/.../services/SettingsService.ts` and docs/CONFIGURATION.md together.
  `node --test dataverse/provision.test.mjs` must pass.
