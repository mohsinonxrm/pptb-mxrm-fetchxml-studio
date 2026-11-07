# Dev harness: run FetchXML Studio against a live org

`npm run dev` opens FXS in a plain browser tab, but there is no PPTB host there, so
`window.dataverseAPI` / `window.toolboxAPI` are missing and nothing loads. The **dev harness**
(`npm run dev:harness`) runs the real UI in a browser against a live Dataverse environment
instead. Use it for fast iteration, live-org checks of fixes, screenshots and scripted smoke
tests. It is dev-only: nothing under `dev/` is part of `vite build` or the npm package.

> Use a **development / sandbox** environment. The harness makes real requests as you, so saving
> views, deleting records, bulk delete and so on change real data.

## Prerequisites

1. Node and `npm ci`, as for normal development.
2. The [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli), signed in with an
   account that has access to the environment:

    ```bash
    az login                      # add --tenant <tenant-id> if you have several tenants
    az account get-access-token --resource https://<org>.crm.dynamics.com -o none   # sanity check
    ```

    On Windows the harness also finds `az` in its default install folder
    (`C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd`) when it isn't on `PATH`. Set
    `AZURE_CLI_PATH` to use another location.

## Run it

Set the org URL, either in the shell or in a git-ignored `.env.local` file
(`DATAVERSE_URL=https://<org>.crm.dynamics.com`), then start the harness:

```bash
# bash / zsh
DATAVERSE_URL=https://<org>.crm.dynamics.com npm run dev:harness
```

```powershell
# PowerShell
$env:DATAVERSE_URL = "https://<org>.crm.dynamics.com"; npm run dev:harness
```

This opens `http://localhost:5173/`. The terminal logs each proxied call with its method, path,
status and duration only.

| Variable                     | Required | Description                                                                                                          |
| ---------------------------- | -------- | -------------------------------------------------------------------------------------------------------------------- |
| `DATAVERSE_URL`              | yes      | Org URL, e.g. `https://contoso.crm.dynamics.com`.                                                                    |
| `DATAVERSE_CALLER_OBJECT_ID` | no       | Microsoft Entra object id of a user to impersonate. Sent as `CallerObjectId` on every request, as PPTB does.         |
| `AZURE_CLI_PATH`             | no       | Full path to the `az` executable, if it isn't on `PATH` or in the default Windows location.                          |

Query-string options:

- `?theme=dark` or `?theme=light` forces the host theme. Otherwise it follows `prefers-color-scheme`.
- `?fetchXml=<url-encoded FetchXML>` or `?prefill=<url-encoded JSON>` simulates a Tool-to-Tool
  launch with that prefill (see `pptb.config.json` for the prefill contract).

## How it works

- `npm run dev:harness` is `vite --mode harness --open`. Only in that mode does `vite.config.ts`
  load `dev/harness/vite-plugin-dataverse-proxy.ts`, which:
    - injects `dev/harness/pptbStub.ts` into `index.html` ahead of `/src/main.tsx`, so the app
      itself is unchanged;
    - proxies `/__dv/*` to `<org>/api/data/v9.2/*` and adds a bearer token from
      `az account get-access-token --resource <org>`. **The token stays in the Node process.** It
      is cached in memory, refreshed about 5 minutes before it expires, and never logged, written
      to disk or sent to the browser.
- `pptbStub.ts` installs `window.toolboxAPI` and `window.dataverseAPI` with the members FXS uses.
  `dataverseStub.ts` shapes requests the way the PPTB host does (desktop-app v1.2.7,
  `dataverseManager.ts`): the same entity-set pluralization, `execute()` URL and parameter
  encoding, `Prefer: return=representation` (plus `odata.include-annotations="*"` for
  `queryData` / `fetchXmlQuery`) and the same error text, so host behaviour reproduces here.
- The proxy refuses to start if the dev server listens beyond localhost. It answers only loopback
  requests that carry the `x-fxs-harness` header and a localhost `Host` / `Origin`, and forwards
  only to the org's `/api/data/v9.2/` path.

## Limits

- Only the host APIs FXS uses are implemented: `dataverseAPI.queryData`, `fetchXmlQuery`,
  `getEntityMetadata`, `execute`, `create`, `update`, `delete`; `toolboxAPI.connections`,
  `utils` (theme, clipboard, notifications), `settings` (backed by `localStorage`),
  `fileSystem.saveFile` (a browser download), `events` (no events are emitted), `getToolContext`
  and `invocation` (no other tools are installed, so Send to Tool lists none).
- Notifications are written to the browser console.
- It doesn't replace testing in PPTB itself. Do a final smoke test in PPTB before a release.
