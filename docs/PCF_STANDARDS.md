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
  Fixed brand colours are allowed only in `categoryColor()` (category badges) and `FileTypeIcon` (file-type colours).
- Icons: import from `./Icons`. To add one, copy the 20px Regular path from `@fluentui/react-icons` into
  `Icons.tsx` with the same `make()` helper. Do **not** add `@fluentui/react-icons` as a dependency.
- Wrap the tree in `FluentProvider` with `context.fluentDesignLanguage?.tokenTheme ?? webLightTheme` (done in `index.ts`).

## 4. Loading, errors and timeouts

- First load (no documents yet): show `<SkeletonCards>` in the same grid class as the real cards, so there's no
  layout shift. Later refreshes keep the current cards and only set `aria-busy`.
- Preview loading: `<SkeletonPreview>`, not a spinner.
- All HTTP goes through `DocumentApi.call()` → `fetchWithTimeout()` (`services/http.ts`). Timeouts: 8 s default,
  15 s preview, `60 s + bytes/256 ms` upload. Token acquisition uses `withTimeout(…, 10 s)`. The popup fallback
  is not timed.
- A failed list shows an error `MessageBar` with a **Retry** action. Action errors (upload, download, delete) show
  an error `MessageBar`. Never leave a spinner or skeleton with no way out.

## 5. User actions

- **Upload:** the button locks on the first click (`submittingRef`, then `submitting` state) and stays disabled
  until the upload finishes. Files go up one at a time. Files that fail **stay staged** for retry, and successful ones
  are cleared. The list reloads only if at least one file succeeded.
- **Delete:** `DeleteDialog` (`modalType="alert"`) shows file, category, source and date. Delete stays disabled
  until the acknowledgement checkbox is ticked, and the dialog can't be dismissed while deleting. On success the
  card is removed from local state and the list is **not** refetched. The update is **confirmed, not
  optimistic**: the server can refuse with 400 or 409, so the card stays until the server says 204.
- **Preview:** the iframe is created only after the user clicks **Preview** and the embed URL has arrived. It is
  sandboxed (`allow-scripts allow-same-origin allow-popups allow-forms`) and removed when the pane closes.
- **Download:** ask the API for the short-lived `downloadUrl`, then click a temporary `<a>`. Never link the
  SharePoint `webUrl` as the download.

## 6. Testing the UI

`npm start watch` runs the PCF harness, but the harness has no backend or MSAL, so you'll only see the
"not configured" message or auth errors. To check the UI, render `DocumentManager` with a fake `DocumentApi`
(an object with `listByContact`, `upload`, `previewUrl`, `downloadUrl` and `remove`) bundled with esbuild, as was
done for the screenshots in `docs/`. Check at least: skeleton → cards, both tabs, preview open/close, delete
dialog gating, and a double-clicked Upload sending one request.

## 7. Bundle hygiene

- `npm run build:prod` → `out/controls/CategorizedDocManager/bundle.js` should stay around **330 KB** (about 80 KB
  gzipped). MSAL is about 284 KB of that. A big jump usually means React, Fluent or Griffel got bundled. Check with
  `grep -c griffel bundle.js` (it should be 0).
- No date or format libraries. Use `formatDate` / `formatSize` in `utils.ts`.
