# Releasing FetchXML Studio

The npm package `@mohsinonxrm/pptb-fetchxml-studio` is published by CI
(`.github/workflows/release.yml`) when a `v*` tag is pushed. Nobody runs `npm publish`
locally. The same process is used by Command Bar Studio and Dataverse Request Studio.

## Branches

| Branch | Holds | Tags | npm |
|---|---|---|---|
| `main` | The current **stable** release, plus hotfixes | `vX.Y.Z` | `latest` |
| `release/X.Y` | The next minor release while it **stabilizes**; exists from the feature freeze until `X.Y.0` ships | `vX.Y.0-beta.N`, `vX.Y.0-rc.N` | `beta` |
| `dev` | Development of the release after that; **never frozen** | none | — |
| `support/X.Y` | Critical fixes for the previous minor release; created only when needed (see *Supported versions*) | `vX.Y.Z` | `support-X.Y` |

- Work happens on branches cut from `dev`, merged into `dev` by PR (squash is fine).
- Merges **between** long-lived branches (`main`, `release/*`, `support/*`, `dev`) always
  use a **merge commit**, so history is shared and fixes are never lost.
- CI (`ci.yml`) runs on every PR and on pushes to `dev`, `main`, `release/**` and
  `support/**`. A ruleset on those branches requires PRs and the `check` status.

## Every PR

- Add your user-facing changes under `## [Unreleased]` in `CHANGELOG.md`, and update the
  README's feature docs if behaviour changes.
- Don't change the version: it's set when a release is cut. That keeps PRs independent,
  so they can merge in any order.
- Put `Closes #N` in the PR description. PRs into `dev` close the issue on merge. For PRs
  into `main`, `release/*` or `support/*`, the issue closes when the squash commit (which
  includes the description) reaches `dev` through the forward merge.

## Where a fix goes

**Fix on the oldest supported branch that has the bug, then merge forward:**
`support/X.Y` → `main` → `release/X.Y` → `dev`. Never copy a fix backwards from `dev`.

| The bug exists in | Fix PR into | Release | Then merge forward |
|---|---|---|---|
| The current stable release (and later) | `hotfix/X.Y.Z` → `main` | tag `vX.Y.Z` on `main` → `latest` | `main` → `release/*` (if open) → `dev` |
| Only the release being stabilized | `release/X.Y` | tag `vX.Y.0-beta.N` → `beta` | `release/X.Y` → `dev` |
| Only `dev` | `dev` | with the next release | — |
| The previous minor only (support window) | `support/X.Y` | tag `vX.Y.Z` → `support-X.Y` | to `main` if it applies there too |

When a forward merge conflicts badly because the code has moved on, re-apply the fix on
the newer branch in a small companion PR instead.

**Resolving forward-merge conflicts** in `CHANGELOG.md` and `package.json`: keep the
target branch's version and its `[Unreleased]` section on top, and add the incoming dated
sections in version order.

## Release lifecycle

### 1. Feature freeze: cut the release branch

Once the release's scope is in `dev`:

```bash
git switch dev && git pull && git switch -c release/1.5 && git push -u origin release/1.5
```

`dev` carries on with the next release straight away.

### 2. Betas (from `release/X.Y`)

1. On a branch from `release/1.5`: `npm version 1.5.0-beta.13 --no-git-tag-version`,
   rename `## [Unreleased]` to `## [1.5.0-beta.13] - YYYY-MM-DD` (add a new empty
   `## [Unreleased]` above it), update the README version badge, and PR it into
   `release/1.5`.
2. Tag that commit on `release/1.5`: `git switch release/1.5 && git pull && git tag v1.5.0-beta.13 && git push origin v1.5.0-beta.13`.
3. Merge `release/1.5` into `dev` (merge commit) so `dev` has the fixes.

`-beta.N` and `-rc.N` publish to the npm `beta` tag as a GitHub pre-release; `latest`
doesn't change. The build bundles the gold Insider icon for pre-releases.

### 3. Stabilizing (beta → stable)

Before promoting a beta to `latest`:

1. **Milestone clear:** every issue in the milestone is closed or moved.
2. **Smoke test** the latest beta, installed from npm, in the PPTB version the tool
   targets (`minAPI`). Run the manual checks the tracking issue lists.
3. **Soak:** 7 days with no new P0/P1 bug against the latest beta. Each fix ships as a
   new `beta.N`, and the clock restarts.
4. **Health:** CI green on `release/X.Y`; `npm audit --omit=dev --audit-level=high`
   reports 0.

### 4. Stable release (to `main`)

1. Release-prep PR into `release/1.5`: `npm version 1.5.0 --no-git-tag-version`, a
   `[1.5.0]` CHANGELOG section that summarises the betas above their history (see
   `[1.4.2]`), and the README badge.
2. PR `release/1.5` → `main`, merged with a merge commit.
3. Tag the merge commit on `main`: `git switch main && git pull && git tag v1.5.0 && git push origin v1.5.0`.
4. Merge `main` into `dev` (merge commit), then delete `release/1.5`.

## Hotfixes (to the current stable release)

1. Branch `hotfix/1.4.3` from `main`; fix, set the version (`1.4.3`), date its CHANGELOG
   section, update the badge; PR into `main` (merge commit).
2. Tag `v1.4.3` on `main` → npm `latest`.
3. Merge `main` forward into `release/*` (if one is open) and `dev`, keeping each target's
   version and `[Unreleased]` entries.

## Supported versions

- **The current stable release** (`main`) gets fixes as hotfixes.
- **The previous minor release** gets **critical and security fixes for 30 days** after
  the next minor ships. FetchXML Studio 1.5 requires PPTB 1.2.7, so 1.4.x stays
  available for users who can't update PPTB yet. When such a fix is needed, create
  `support/1.4` from the last `v1.4.*` tag
  (`git switch -c support/1.4 v1.4.2 && git push -u origin support/1.4`), PR the fix into
  it, and tag `v1.4.4` there. It publishes to the npm `support-1.4` tag, never `latest`,
  and isn't marked as the latest GitHub Release.
- Older versions aren't supported. `SECURITY.md` lists the supported versions.

## Milestones and tracking

- Each release has a milestone (`v1.5.0`, `v1.4.3`). It's the release that **ships** the
  change, and its closed issues are the list of what the release contains. Close issues
  as they're done; that doesn't lose track of them.
- What each beta contains is its `[X.Y.Z-beta.N]` CHANGELOG section and its GitHub
  pre-release notes.
- A **"Release X.Y.Z" tracking issue** in the minor release's milestone holds the
  stabilizing checklist. The milestone closes when the stable release ships.

## What `release.yml` does

Checks the tag matches `package.json` and chooses the channel from the version and the
tagged commit:

| Tag | Tagged commit must be on | npm tag | GitHub Release |
|---|---|---|---|
| `vX.Y.Z-beta.N`, `vX.Y.Z-rc.N` | any branch (normally `release/X.Y`) | `beta` | pre-release |
| `vX.Y.Z` | `support/X.Y` | `support-X.Y` | not marked latest |
| `vX.Y.Z` | `main`, with `X.Y` = main's own version line | `latest` | latest |

`support/X.Y` is checked first: it merges forward into `main`, so an older line's commit can
be on `main` too, and it must never move `latest` back.

Any other tag fails before publishing. Afterwards, replace the generated release notes with
the CHANGELOG section.

It runs in separate jobs, so the code that installs and runs dependencies never holds
publishing rights (#188):

| Job | Runs on | Permissions | What it does |
|---|---|---|---|
| `build` | always | read the repository | Checks the tag and chooses the channel; `npm ci --ignore-scripts` (no cache); the runtime audit (no high or critical findings), lint, tests, validation, build (with the type check), the DOMPurify check; packs the tarball and checks its contents. |
| `publish` | tag push only | npm OIDC token (`id-token: write`) | Installs no project dependencies: downloads the tarball and publishes it with provenance. |
| `rehearse` | manual run only | none | `npm publish --dry-run` of the tarball, without provenance or an npm token. |
| `github-release` | tag push only | write the repository | Creates the GitHub Release. |

Node (24.22.0), npm (11.21.1) and the actions (by full commit SHA, the tag in a comment) are
pinned to versions at least 14 days old. Dependabot proposes newer actions on `dev`; other
branches get them when `dev` reaches them through a release. Move the Node and npm pins by
hand. Only one release run per ref runs at a time. If a run fails after publishing, use
*Re-run failed jobs*: a full re-run stops at `publish`, since npm never accepts a version twice.

**Rehearse a release** after the version bump and before the tag, on a branch whose workflow
changed: `gh workflow run release.yml --ref <branch>` (or Actions → *Release* → *Run workflow*
once `dev` has the manual trigger). It runs every check and `npm publish --dry-run`, and
creates no GitHub Release. npm's dry run refuses a version that's already published, so on an
unbumped branch the rehearsal stops with a notice after the checks. It doesn't exercise the
npm token exchange or provenance signing. A manual run never publishes, even on a tag.

## One-time setup

- **npm trusted publisher**: npmjs.com → package → *Settings* → *Trusted publishing* →
  *GitHub Actions*: user `mohsinonxrm`, repository `pptb-mxrm-fetchxml-studio`, workflow
  `release.yml`, no environment, *Allow npm publish*. No npm token is needed.
- `repository.url` in `package.json` must point at this GitHub repository, or provenance
  and trusted publishing fail.
- A ruleset on `dev`, `main`, `release/**` and `support/**` requiring PRs and the `check`
  status.
- Repository *Settings* → *Pull Requests* → squash merging → default commit message
  **"Pull request title and description"**, so `Closes #N` travels with the commit.
