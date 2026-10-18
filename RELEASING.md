# Releasing FetchXML Studio

The npm package `@mohsinonxrm/pptb-fetchxml-studio` is published by CI
(`.github/workflows/release.yml`) when a `v*` tag is pushed. Nobody runs `npm publish`
locally.

## Branches

- Work happens on branches cut from `dev` and merged into `dev` by PR.
- A release is a PR from `dev` into `main`. Merge it with a **merge commit**, so `main`
  and `dev` keep the same history.
- CI (`ci.yml`) runs on every PR and on pushes to `dev` and `main`.

## Cutting a release

1. On a branch from `dev`, set the version and date the changelog:

   ```bash
   npm version 1.4.2 --no-git-tag-version          # or e.g. 1.4.2-beta.1
   ```

   In `CHANGELOG.md`, rename `## [Unreleased]` to `## [1.4.2] - YYYY-MM-DD` and add a
   new empty `## [Unreleased]` above it, and update the version badge at the top of
   `README.md`. Open a PR into `dev`.

2. Open the release PR `dev → main` and merge it with a merge commit.

3. Tag the merge commit on `main` and push the tag:

   ```bash
   git switch main && git pull
   git tag v1.4.2 && git push origin v1.4.2
   ```

4. `release.yml` checks that the tag matches `package.json`, runs the tests, validation
   and build, publishes to npm, and creates the GitHub Release. Versions with a
   pre-release part (`-beta.1`) go to the `beta` npm tag and are marked as pre-releases;
   everything else goes to `latest`. The build bundles the gold Insider icon for
   pre-release versions automatically.

## One-time setup

- **npm trusted publisher**: on npmjs.com → package → *Settings* → *Trusted publishing*
  → *GitHub Actions*: user `mohsinonxrm`, repository `pptb-mxrm-fetchxml-studio`,
  workflow `release.yml`. No npm token is needed.
- `repository.url` in `package.json` must point at this GitHub repository, or
  provenance and trusted publishing fail.
- Branch protection (or a ruleset) on `dev` and `main` requiring the `CI / check` job.
