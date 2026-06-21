#!/usr/bin/env bash
#
# release.sh — cut a pptchartengine release via GitHub Actions + PyPI Trusted Publishing (path A).
#
# What it does:
#   1. Preflight: clean tree, on main, version sync (pyproject == __version__ == CHANGELOG),
#      CHANGELOG entry is dated (not "Unreleased"), and — for a first publish — the name is
#      still free on PyPI.
#   2. Local verification mirroring docs/release.md: pytest, clean build, twine check, pip check.
#   3. Tag vX.Y.Z, push it, and create the GitHub Release.
#
# The published GitHub Release triggers .github/workflows/publish.yml, which uploads to PyPI via
# OIDC Trusted Publishing. This script NEVER runs `twine upload` and needs no PyPI token.
#
# Prereq (one-time, manual on the PyPI website — cannot be scripted):
#   Configure a Trusted Publisher / pending publisher on PyPI as documented in docs/release.md
#   (project pptchartengine, owner hanlinlibham, repo pptchartengine, workflow publish.yml,
#   environment pypi).
#
# Usage:
#   ./scripts/release.sh            # full release
#   ./scripts/release.sh --dry-run  # run all checks, stop before tag/push/release
#
set -euo pipefail

DRY_RUN=0
[[ "${1:-}" == "--dry-run" ]] && DRY_RUN=1

cd "$(dirname "$0")/.."
ROOT="$(pwd)"
echo "==> repo: $ROOT"

fail() { echo "ERROR: $*" >&2; exit 1; }

# --- 1. preflight ---------------------------------------------------------
command -v gh >/dev/null   || fail "gh CLI not found (needed to create the GitHub Release)."
command -v python3 >/dev/null || fail "python3 not found."

BRANCH="$(git rev-parse --abbrev-ref HEAD)"
[[ "$BRANCH" == "main" ]] || fail "not on main (on '$BRANCH'). Releases publish from main."
[[ -z "$(git status --porcelain)" ]] || fail "working tree is dirty. Commit or stash first."

PYPROJECT_VERSION="$(python3 - <<'PY'
import re, pathlib
m = re.search(r'^version\s*=\s*"([^"]+)"', pathlib.Path("pyproject.toml").read_text(), re.M)
print(m.group(1) if m else "")
PY
)"
[[ -n "$PYPROJECT_VERSION" ]] || fail "could not read version from pyproject.toml."

PKG_VERSION="$(python3 -c "import sys; sys.path.insert(0,'src'); import pptchartengine; print(pptchartengine.__version__)")"
[[ "$PYPROJECT_VERSION" == "$PKG_VERSION" ]] \
  || fail "version mismatch: pyproject=$PYPROJECT_VERSION vs __version__=$PKG_VERSION."

VERSION="$PYPROJECT_VERSION"
TAG="v$VERSION"
echo "==> releasing version: $VERSION  (tag $TAG)"

# CHANGELOG must have a dated entry for this version (not "Unreleased").
grep -qE "^## ${VERSION//./\\.}[[:space:]]+-[[:space:]]+[0-9]{4}-[0-9]{2}-[0-9]{2}" CHANGELOG.md \
  || fail "CHANGELOG.md has no dated '## $VERSION - YYYY-MM-DD' entry (still 'Unreleased'?)."

# tag must not already exist
git rev-parse -q --verify "refs/tags/$TAG" >/dev/null \
  && fail "tag $TAG already exists." || true

# first-publish name-availability guard (informational on later releases)
HTTP="$(curl -s -o /dev/null -w '%{http_code}' "https://pypi.org/pypi/pptchartengine/json" || echo 000)"
if [[ "$HTTP" == "404" ]]; then
  echo "==> PyPI: pptchartengine not yet published (first release will claim the name)."
elif [[ "$HTTP" == "200" ]]; then
  echo "==> PyPI: project exists (this is a follow-up release)."
else
  echo "WARN: could not check PyPI (HTTP $HTTP); continuing."
fi

# --- 2. local verification (mirrors docs/release.md) ----------------------
echo "==> pytest"
python3 -m pytest -q
echo "==> clean build"
rm -rf dist build
python3 -m build --sdist --wheel
echo "==> twine check"
python3 -m twine check dist/*
echo "==> pip check"
python3 -m pip check || echo "WARN: pip check reported issues (review above)."

if [[ "$DRY_RUN" == "1" ]]; then
  echo "==> --dry-run: all checks passed; stopping before tag/push/release."
  exit 0
fi

# --- 3. tag + GitHub Release (triggers publish.yml -> PyPI) ----------------
printf '==> create tag %s, push, and publish GitHub Release? [y/N] ' "$TAG"
read -r ans
[[ "$ans" == "y" || "$ans" == "Y" ]] || { echo "aborted."; exit 1; }

git tag -a "$TAG" -m "pptchartengine $VERSION"
git push origin "$TAG"

# Extract this version's CHANGELOG section as the release notes.
NOTES="$(python3 - "$VERSION" <<'PY'
import sys, re, pathlib
ver = sys.argv[1]
text = pathlib.Path("CHANGELOG.md").read_text()
blocks = re.split(r'(?m)^## ', text)
for b in blocks:
    if b.startswith(ver):
        sys.stdout.write("## " + b.strip() + "\n")
        break
PY
)"

gh release create "$TAG" \
  --title "pptchartengine $VERSION" \
  --notes "${NOTES:-Release $VERSION}" \
  --verify-tag

echo
echo "==> GitHub Release $TAG published."
echo "==> Trusted-Publishing workflow is now running. Watch it with:"
echo "      gh run watch \$(gh run list --workflow=publish.yml --limit=1 --json databaseId -q '.[0].databaseId')"
echo "==> When green, verify: https://pypi.org/project/pptchartengine/$VERSION/"
