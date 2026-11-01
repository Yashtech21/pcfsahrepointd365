# PCF & React Coding Standards

Applies to everything under `pcf/`. Current code already follows these rules. Keep it that way.

## 1. Lifecycle and request deduplication

- Dynamics calls `updateView(context)` many times during form load and on every field change.
- `updateView` **only builds props**. It must:
  - reuse the cached `host` object unless `entityName`, `recordId`, `contactId` or `contactName` changed;
  - reuse the cached `categories` array unless `entityName` or `categoryConfig` changed;
  - reuse the single `DocumentApi` unless one of `apiBaseUrl`, `tenantId`, `clientId` or `apiScope` changed.
- **Never fetch in `updateView`.** Data loading happens in `DocumentManager` in a `useEffect` keyed on
  `[api, host.contactId]`.
- `DocumentApi.listByContact` shares one in-flight request per `contactId`. `DocumentManager` uses a
  request sequence number so only the latest response updates state, even when the lookup changes quickly.
- Async values that arrive later (entity display names, UPN) go into React state or a promise passed down.
  **Do not** call `notifyOutputChanged()` to force a re-render; the control has no outputs.

## 2. Rendering performance

- Derived lists (filter + sort for the active tab) use `React.useMemo`.
- Every handler passed to a child uses `React.useCallback`, and `DocumentCard` is `React.memo`. Opening a preview
  must re-render only the previously and newly selected cards.
- Don't create inline object or array props on memoized children in hot paths. Inline arrow handlers are fine
  on unmemoized leaf elements such as the sort buttons.
- `key` = the Graph driveItem `id`. Never use the array index.

## 3. UI and Fluent UI v9

- Use only `@fluentui/react-components` (platform library 9.46.2). APIs added after 9.46 are **not available** at
  runtime, so check the installed `.d.ts` before using a component.
- Style with `makeStyles` and `tokens.*` for spacing, colour, typography and radius. Griffel does not allow CSS shorthands
  such as `borderColor`, `padding: "a b"` or `border` on some properties. Use `shorthands.*()` or longhands.
  Fixed colours are allowed only in `categoryColor()` (category badges) and `FileTypeIcon` (file-type colours).
  Everything else, including the preview background, uses tokens so dark themes work.
- Icons: import from `./Icons`. To add one, copy the 20px Regular path from `@fluentui/react-icons` into
  `Icons.tsx` with the same `make()` helper. Do **not** add `@fluentui/react-icons` as a dependency.
- Wrap the tree in `FluentProvider` with `context.fluentDesignLanguage?.tokenTheme ?? webLightTheme` (done in `index.ts`).

## 4. Loading, errors and timeouts

- First load (no documents yet): show `<SkeletonCards>` in the same grid class as the real cards, so there's no
  layout shift. Later refreshes keep the current cards and only set `aria-busy`.
- Preview loading: `<SkeletonPreview>`, not a spinner.
- All HTTP goes through `DocumentApi.call()` → `fetchWithTimeout()` (`services/http.ts`). Timeouts are the
  constants `DEFAULT_TIMEOUT_MS`, `PREVIEW_TIMEOUT_MS` and `uploadTimeoutMs()` in `DocumentApi.ts`, plus
  `SILENT_TIMEOUT_MS` in `AuthService.ts` for silent token calls. The popup fallback is not timed. Change the
  constants, not the call sites.
- A failed list shows an error `MessageBar` with a **Retry** action. Action errors (upload, download, delete) show
  an error `MessageBar`. Never leave a spinner or skeleton with no way out.

## 5. User actions

- **Upload:** the button locks on the first click (`submittingRef`, then `submitting` state) and stays disabled
  until the upload finishes. Files go up one at a time. Files that fail **stay staged** for retry, and successful ones
  are cleared. The list reloads if any file succeeded or timed out.
- **Upload timeouts can't create duplicates:** a client timeout doesn't mean the server failed. After the reload,
  `landedDespiteTimeout()` treats a timed-out file as uploaded when a matching document (same size, this record,
  recent, same base name allowing for SharePoint's ` 1` rename) exists, and removes it from the staged list.
- **Server rejections are final:** `415` (blocked or disguised file) and `400` messages are shown per file. Don't
  add client-side copies of the blocked-type list; the server (`fileRules.ts`) is the single source of truth.
- **Delete:** `DeleteDialog` (`modalType="alert"`) shows file, category, source and date. Delete stays disabled
  until the acknowledgement checkbox is ticked, and the dialog can't be dismissed while deleting. On success the
  card is removed from local state and the list is **not** refetched. The update is **confirmed, not
  optimistic**: the server can refuse with 400 or 409, so the card stays until the server says 204.
- **Preview:** the iframe is created only after the user clicks **Preview** and the preview target has arrived. It is
  sandboxed (`allow-scripts allow-same-origin allow-popups allow-forms`), uses `referrerPolicy="no-referrer"`, and
  is removed when the pane closes. A `{ url }` target becomes the iframe `src`. A `{ url, postParameters }` target is
  sent by a hidden React-rendered `<form method="post" target={frameName}>` submitted from an effect. Non-https
  URLs are refused on the client as well as the server.
- **Download:** ask the API for the short-lived `downloadUrl`, then click a temporary `<a>`. Never link the
  SharePoint `webUrl` as the download.

## 5a. Indicators and counts

Every indicator is derived with `useMemo` from state that already exists. Don't add extra fetches to compute counts.

| Indicator | Where | Source |
|---|---|---|
| Tab counts (`CounterBadge`) | Both tabs. Hidden during first load | `recordCount`, `documents.length` |
| Category chips with counts (`ToggleButton` + `CounterBadge`) | Above the grid when the tab has 2+ categories; click to filter, click again to clear | `categoryCounts` (from `scopeDocs`) |
| "N items · size" / "Showing X of Y items · size" | Grid footer (`aria-live`) | `displayed`, `scopeDocs`, `displayedBytes` |
| Staged summary "N files selected · size" | Drop zone | `staged` in `UploadBar` |
| Upload progress bar + "Uploading i of n: name · sent of total (pct)" | Under the upload bar; button shows `Uploading i/n…`, then `Finishing…` while the list refreshes | `IUploadProgress` from `xhrUpload` progress events |
| Refresh bar (indeterminate `ProgressBar`) | Under the grid header on reloads (not first load, which uses skeletons) | `loading && !firstLoad` |
| "New" badge | Cards uploaded in this session | `newIds` |
| "Deleting…" | Delete dialog button | `deleting` |

Uploads use `xhrUpload` (`services/http.ts`) because `fetch` can't report upload progress. It keeps the same
timeout contract (rejects with `TimeoutError`). Counts include all files in the scope, not just the filtered ones.
Chip counts follow the active tab.

## 6. Testing the UI

`npm start watch` runs the PCF harness, but the harness has no backend or MSAL, so you'll only see the
"not configured" message or auth errors. To check the UI, render `DocumentManager` with a fake `DocumentApi`
(an object with `listByContact`, `upload`, `previewUrl`, `downloadUrl` and `remove`) bundled with esbuild, as was
done for the screenshots in `docs/`. Check at least: skeleton → cards, both tabs, preview open/close, delete
dialog gating, and a double-clicked Upload sending one request.

## 7. Bundle hygiene

- `npm run build:prod` → `out/controls/CategorizedDocManager/bundle.js` is roughly 330 KB (about 80 KB gzipped),
  and MSAL accounts for most of it. Treat a jump of more than about 50 KB as a regression to explain. A big jump usually means React, Fluent or Griffel got bundled. Check with
  `grep -c griffel bundle.js` (it should be 0).
- No date or format libraries. Use `formatDate` / `formatSize` in `utils.ts`.
