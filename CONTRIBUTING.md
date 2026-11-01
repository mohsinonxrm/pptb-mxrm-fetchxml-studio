# Contributing to FetchXML Studio

Thanks for helping. This guide covers setup, the branch and PR workflow, and what a PR
needs before it can merge. Releases are described in [RELEASING.md](RELEASING.md).

## Prerequisites

- **Node.js 22.22.2+, 24.15.0+ or 26+** (`engines` in `package.json`, the range supported by jsdom, the Vitest environment; CI uses Node 24) and npm.
- **Power Platform ToolBox 1.2.7 or later** to run the tool for real (`minAPI` 1.2.7).
- For the browser dev harness: the **Azure CLI** (`az login`) and a Dataverse dev org.

## Setup

```bash
npm ci                 # exact versions from package-lock.json
npm run dev            # Vite dev server (no Dataverse)
npm run dev:harness    # the real UI in a browser against a live org; see dev/README.md
```

Before opening a PR, run what CI runs:

```bash
npm run typecheck
npm run lint           # zero warnings
npm test
npm run validate       # PPTB manifest rules
npm run build
```

CI also fails on any **high or critical** `npm audit --omit=dev` finding and checks the
package contents (`.github/scripts/check-pack.mjs`).

To try a build inside PPTB, load the repository folder as a local tool after
`npm run build`.

## Branches and PRs

- Branch from `dev`: `feat/<issue>-<slug>`, `fix/<issue>-<slug>`, `chore/<issue>-<slug>`
  or `docs/<issue>-<slug>`.
- Open the PR into `dev`. A fix for the release being stabilized goes into `release/X.Y`,
  and a fix for the current stable release goes into `main` as a `hotfix/*` branch. See
  [RELEASING.md](RELEASING.md) → *Where a fix goes*.
- Link the issue in the description: `Closes #123`, or `Part of #123` while manual checks
  are pending.
- The `check` status must pass. PRs into `dev` are squash-merged.
- Commit messages follow [Conventional Commits](https://www.conventionalcommits.org/):
  `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`.

## What every PR includes

- **CHANGELOG:** user-facing changes go under `## [Unreleased]` in `CHANGELOG.md`.
- **No version change.** The version, the README version badge and the dated CHANGELOG
  section change only in release-prep PRs, so PRs stay independent.
- **README:** update the feature docs when behaviour changes, and the
  "What this tool changes" section when an operation writes data.
- **Tests:** unit tests for changes to the FetchXML model, parser, serializer or rules.

## Product rules

- **Ground behaviour in Microsoft Learn, then verify it live.** Cite the Learn page in the
  code or PR. When the docs and Dataverse disagree, Dataverse's actual behaviour wins, and
  the PR says what was verified and how.
- **Never refuse supported FetchXML.** Block only what is known to be invalid; warn only
  when the outcome can't be determined. A Dataverse error is an acceptable outcome.
- **Data-changing operations** (delete, bulk delete, state changes, workflows, saving
  views) confirm with the **environment** and the **scope** (entity and record count), per
  the PPTB Marketplace policy.

## UI changes

- **Agree on the design first.** Share a prototype or screenshots in the issue before
  building.
- **Fluent UI v9 only** (`@fluentui/react-components`), following the PPTB theme in light
  and dark. PR screenshots show both themes.
- **Wording:** sentence case, "…" for ellipses, and "Entity" / "Attribute" as the nouns.
  The same wording is shared with Command Bar Studio and Dataverse Request Studio.

## Dependencies

- Prefer the platform and existing dependencies. New dependencies need a reason in the PR.
- **Only versions published at least 14 days ago.** Dependabot's cooldown applies the same
  rule to its version updates.
- After changing dependencies, run `npm install` and commit `package-lock.json`.

## Security

Report vulnerabilities privately; see [SECURITY.md](SECURITY.md). Never commit tokens,
org URLs with credentials, or `.env.local`.
