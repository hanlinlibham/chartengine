# Release Checklist

This project publishes with GitHub Actions and PyPI Trusted Publishing.

## Current GitHub Setup

- Repository: `hanlinlibham/pptchartengine`
- CI workflow: `.github/workflows/ci.yml`
- Publish workflow: `.github/workflows/publish.yml`
- GitHub environments:
  - `testpypi`
  - `pypi`

The publish workflow does not use a long-lived PyPI token. It requests an OIDC
token through GitHub Actions and uses `pypa/gh-action-pypi-publish`.

## TestPyPI Trusted Publisher

Configure a pending publisher on TestPyPI with:

- PyPI project name: `pptchartengine`
- Owner: `hanlinlibham`
- Repository name: `pptchartengine`
- Workflow filename: `publish.yml`
- Environment name: `testpypi`

Then run the `Publish Python Package` workflow manually from GitHub Actions.
Manual dispatch publishes only to TestPyPI.

## PyPI Trusted Publisher

Configure a pending publisher on PyPI with:

- PyPI project name: `pptchartengine`
- Owner: `hanlinlibham`
- Repository name: `pptchartengine`
- Workflow filename: `publish.yml`
- Environment name: `pypi`

Publishing to production PyPI is tied to a published GitHub Release. Do not
publish directly from a regular branch workflow run.

## Before Publishing

Run locally:

```bash
python -m pytest -q
rm -rf dist build
python -m build --sdist --wheel
python -m twine check dist/*
python -m pip check
```

Also verify:

- `README.md` states the alpha support scope accurately.
- `CHANGELOG.md` has an entry for the version being released.
- `pyproject.toml` version matches `pptchartengine.__version__`.
- The package name is still available on PyPI before first publish.

## References

- PyPI Trusted Publishers: https://docs.pypi.org/trusted-publishers/
- Publishing with a Trusted Publisher: https://docs.pypi.org/trusted-publishers/using-a-publisher/
- Creating a project through OIDC: https://docs.pypi.org/trusted-publishers/creating-a-project-through-oidc/
