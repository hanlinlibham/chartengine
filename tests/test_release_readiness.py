"""Release-readiness contracts for the public package surface.

These tests guard publishing details that are easy to regress while editing
docs or packaging metadata. They are intentionally lightweight and local-only;
full wheel/sdist build checks still belong in release CI.
"""

from __future__ import annotations

try:
    import tomllib  # Python 3.11+
except ModuleNotFoundError:  # Python 3.10
    import tomli as tomllib
from pathlib import Path

import pytest

import ablechart


ROOT = Path(__file__).resolve().parents[1]


def _read_text(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


def test_pyproject_metadata_is_ready_for_public_distribution():
    project = tomllib.loads(_read_text("pyproject.toml"))["project"]

    assert project["name"] == "ablechart"
    assert project["version"] == ablechart.__version__
    assert project["readme"] == "README.md"
    assert project["requires-python"] == ">=3.10"
    assert project["license"] == "MIT"
    assert project["authors"] == [{"name": "ablechart contributors"}]

    urls = project["urls"]
    assert urls["Homepage"] == "https://github.com/hanlinlibham/ablechart"
    assert urls["Repository"] == "https://github.com/hanlinlibham/ablechart"
    assert urls["Issues"] == "https://github.com/hanlinlibham/ablechart/issues"
    assert urls["Changelog"].endswith("/CHANGELOG.md")

    dependencies = set(project["dependencies"])
    assert {
        "python-pptx>=1.0.2,<1.1",
        "pandas>=2.2,<3",
        "numpy>=1.26,<2",
        "lxml>=5,<6",
        "openpyxl>=3.1,<4",
    }.issubset(dependencies)


def test_license_file_is_mit_and_not_stale_apache_text():
    license_text = _read_text("LICENSE")

    assert license_text.startswith("MIT License")
    assert "Permission is hereby granted, free of charge" in license_text
    assert "Copyright (c) 2026 ablechart contributors" in license_text
    assert "Apache License" not in license_text
    assert "Apache-2.0" not in license_text


def test_readme_states_positioning_install_and_comparison():
    readme = _read_text("README.md")
    normalized = " ".join(readme.split())

    required_phrases = [
        "pip install ablechart",
        "chart asset kernel",
        "not a full presentation generation framework",
        "inspect_pptx_charts",
        "replace_pptx_chart_data",
        "python-pptx",
        "PptxGenJS",
        "mschart",
        "Aspose.Slides",
        "Spire.Presentation",
        "MIT License",
    ]
    for phrase in required_phrases:
        assert phrase in normalized


def test_manifest_includes_release_docs_and_pptx_fixtures():
    manifest = _read_text("MANIFEST.in").splitlines()

    assert "include LICENSE" in manifest
    assert "include README.md" in manifest
    assert "include CHANGELOG.md" in manifest
    assert "include pyproject.toml" in manifest
    assert "recursive-include tests *.py *.pptx" in manifest


def test_github_workflows_cover_ci_and_trusted_publishing():
    workflows = ROOT / ".github" / "workflows"
    if not workflows.exists():
        pytest.skip("GitHub workflow files are not included in this source package")

    ci = (workflows / "ci.yml").read_text(encoding="utf-8")
    publish = (workflows / "publish.yml").read_text(encoding="utf-8")

    assert "python -m pytest -q" in ci
    assert "python -m build --sdist --wheel" in ci
    assert "python -m twine check dist/*" in ci
    assert "python-version: [\"3.10\", \"3.11\", \"3.12\"]" in ci

    assert "id-token: write" in publish
    assert "environment:" in publish
    assert "name: testpypi" in publish
    assert "name: pypi" in publish
    assert "pypa/gh-action-pypi-publish@v1.14.0" in publish
    assert "repository-url: https://test.pypi.org/legacy/" in publish
