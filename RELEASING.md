# Releasing

Releases are fully automated. Every push to `main` runs `.github/workflows/release.yml`, which:

1. Runs the test suite.
2. Runs [python-semantic-release](https://python-semantic-release.readthedocs.io/), which parses
   [conventional commits](https://www.conventionalcommits.org/) since the last tag.
   `feat:` bumps minor, `fix:`/`perf:` bump patch; other types release nothing.
3. If a new version is due: bumps `pyproject.toml` and `src/excel_explorer/__init__.py`,
   refreshes `uv.lock`, commits `chore(release): {version}`, tags `v{version}`,
   creates a GitHub Release, and builds the sdist + wheel.
4. Publishes the artifacts to PyPI via [trusted publishing](https://docs.pypi.org/trusted-publishers/) (OIDC, no API token).

## One-time setup (required before the first release)

### 1. GitHub `release` environment

The release job runs in an environment named `release`. Create it under
**Settings → Environments → New environment**, and add yourself as a
**required reviewer** so every PyPI publish needs a human approval.

### 2. `GH_TOKEN` repository secret

semantic-release pushes the version commit and tag back to `main`, which the
default `GITHUB_TOKEN` cannot do if branch protection is enabled (and pushes made
with it don't trigger other workflows). Create a fine-grained PAT:

- Repository access: `pressw-llc/excel-explorer`
- Permissions: **Contents: Read and write**
- If `main` is branch-protected, allow the token's owner to bypass push protection.

Save it as a repository secret named `GH_TOKEN`.

### 3. PyPI trusted publisher

The project name `excel-explorer` must be registered as a *pending publisher*
before the first upload. On PyPI: **Account → Publishing → Add a new pending publisher**:

| Field | Value |
|---|---|
| PyPI project name | `excel-explorer` |
| Owner | `pressw-llc` |
| Repository | `excel-explorer` |
| Workflow name | `release.yml` |
| Environment | `release` |

After the first successful publish the pending publisher converts to a normal
trusted publisher automatically.

### 4. Restrict fork PR workflow runs (when the repo goes public)

This setting only exists on public repositories — while the repo is private,
only invited collaborators can open PRs, so it doesn't apply. If/when the repo
is made public, go to **Settings → Actions → General → Fork pull request
workflows** and select **"Require approval for all outside collaborators"**
(GitHub's default is the weaker "first-time contributors") so unknown
contributors cannot execute CI without a maintainer approving the run.

## Supply-chain notes

- All actions in `.github/workflows/` are pinned to full commit SHAs with a
  `# vX.Y.Z` comment. Dependabot (`.github/dependabot.yml`) opens PRs to bump
  them; review the diff of the upstream action before merging.
- The release job runs with `enable-cache: false` — the publish path never
  restores a GitHub Actions cache, which closes the cache-poisoning route to
  PyPI. Keep it that way.

## Cutting a release

Merge conventional commits to `main`. That's it — a `feat:` or `fix:` commit
triggers a release; `docs:`, `chore:`, `ci:`, etc. do not.

To verify what would be released without releasing, run locally:

```bash
uvx --from python-semantic-release semantic-release version --print
```
