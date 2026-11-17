"""Keep temporary test files under the repository's ignored output directory."""

import tempfile
from pathlib import Path


def temporary_directory():
    directory = Path(__file__).resolve().parents[1] / "output" / "tests"
    directory.mkdir(parents=True, exist_ok=True)
    return tempfile.TemporaryDirectory(dir=directory)
