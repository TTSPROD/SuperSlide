"""Repository resources and the common destination for generated artifacts.

Paths are anchored to this checkout, so CLI defaults do not depend on the
directory from which SuperSlide is started. Explicit user paths stay relative
to the caller's working directory.
"""

from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = PROJECT_ROOT / "output"
DEFAULT_TEMPLATE = PROJECT_ROOT / "templates" / "severstal_source.pptx"
DEFAULT_MANIFEST = PROJECT_ROOT / "templates" / "severstal_template.json"
DEFAULT_ASSET_CATALOG = PROJECT_ROOT / "assets" / "images" / "catalog.json"
SOURCE_PHOTO_CATALOG = PROJECT_ROOT / "assets" / "photos" / "severstal" / "catalog.json"
SOURCE_ICON_CATALOG = PROJECT_ROOT / "assets" / "icons" / "severstal" / "catalog.json"
LEGACY_TEMPLATE = PROJECT_ROOT / "templates" / "legacy" / "corporate_base.pptx"
LEGACY_MANIFEST = PROJECT_ROOT / "templates" / "legacy" / "template_manifest.json"
LEGACY_EXAMPLES = PROJECT_ROOT / "examples" / "legacy"
