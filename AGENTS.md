# AI Assistant Instructions: D365 PCF Document Manager

Read this file first. It is the single source of instructions for AI coding agents (Claude Code reads it via
`CLAUDE.md`; Cursor, Copilot and Windsurf read `AGENTS.md` directly). Details live in `docs/`:

| File | Read it when |
|---|---|
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Changing folders, metadata, auth, or anything crossing the PCF ↔ Function ↔ Graph boundary |
| [docs/PCF_STANDARDS.md](docs/PCF_STANDARDS.md) | Touching anything under `pcf/` |
| [docs/API_SPECS.md](docs/API_SPECS.md) | Calling or changing an endpoint (ours or Microsoft Graph) |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | A build, packaging, auth or runtime error |
| [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md) | Entra, SharePoint, Azure or form setup |

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
4. **Every network call has a timeout.** In the browser, use `fetchWithTimeout` / `withTimeout` from
   `pcf/.../services/http.ts`. Do not call `fetch` directly. Defaults: 8 s for API calls, 15 s for preview,
   60 s plus file size for uploads, 10 s for silent token calls. In the backend, all Graph calls go through
   `graph()` in `backend/src/lib/graph.ts` (30 s, retries 429/503/504).
   Timeouts must surface as a message with a Retry button, never as an endless spinner.
5. **Entity context comes from `context.mode.contextInfo`** (`entityTypeName`, `entityId`). Category lists come
   from `categoriesFor(entityName, categoryConfig)` in `categories.ts`. `context.page` is not used.
6. **No DOM access outside the control.** React state and Fluent components only. Don't touch the D365 page DOM
   or `Xrm`. (The temporary `<a>` element used to start a download is the only exception.)
7. **Destructive actions need confirmation on both sides.** Delete goes through `DeleteDialog` (checkbox
   acknowledgement). `DocumentApi.remove(id, name)` sends `X-Confirm-Delete`, and the backend refuses folders and
   name mismatches. Never add a delete or bulk-delete path that skips either check.
8. **Server is the source of truth** for category → folder mapping (`backend/src/lib/categories.ts`), input
   validation and authorization. Client checks are for UX only.
9. **Never trust `X-MS-CLIENT-PRINCIPAL`.** Callers are authenticated by verifying the bearer JWT in
   `backend/src/lib/auth.ts`.
10. **Secrets never go in the PCF control or the repo.** Control properties hold only public IDs and URLs.
    `GRAPH_CLIENT_SECRET` lives in Function App settings / Key Vault. `local.settings.json` is gitignored.

## Commands

| Task | Command (run from) |
|---|---|
| Install | `npm ci` (`pcf/`, `backend/`) |
| Build control (dev) | `npm run build` (`pcf/`) |
| Build control (prod, check size) | `npm run build:prod` (`pcf/`) |
| Test harness | `npm start watch` (`pcf/`) — harness has no real API; see PCF_STANDARDS §6 |
| Backend build + unit tests | `npm test` (`backend/`) |
| Run backend locally | `cp local.settings.sample.json local.settings.json && npm start` (`backend/`, needs Functions Core Tools) |
| Package solution | `pac solution init …` → `pac solution add-reference --path ../pcf` → `dotnet build -c Release` (or `msbuild /t:build /restore /p:Configuration=Release`) |
| Push to a dev org | `pac pcf push --publisher-prefix ce` (`pcf/`) |

Linting is skipped during build (`pcfconfig.json` → `skipBuildLinting`). Run `npx tsc --noEmit` in `pcf/` as the
type check.

## Definition of done for a change

- `npm run build` and `npm run build:prod` succeed in `pcf/`. `npm test` passes in `backend/`.
- New backend logic that can be tested without Graph (parsing, naming, validation, auth rejection) has a test in
  `backend/test/`.
- If you changed a contract, update `docs/API_SPECS.md`. If you changed a folder, column or flow, update
  `docs/ARCHITECTURE.md`. If you changed a setting, update `docs/DEPLOYMENT.md`.
- When adding a category, update all three places: `backend/src/lib/categories.ts`, `ALL_CATEGORIES` in
  `pcf/.../categories.ts`, and the SharePoint choice values (`sharepoint/Setup-ContactsLibrary.ps1`).
