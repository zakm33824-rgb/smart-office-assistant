"""Template library package for the PPT generator."""

from .catalog import build_source_catalog, build_template_catalog
from .page_engine import build_library_overview, build_page_level_outline, build_page_preview_sheet, build_demo_page_records, extract_pptx_to_slide_catalog, infer_request, search_components, search_pages, seed_demo_page_library
from .paths import ensure_library_structure, library_paths
from .slide_storage import load_component_catalog, load_slide_catalog, load_slide_summary, save_component_catalog, save_slide_catalog
from .storage import load_source_catalog, load_template_catalog, save_source_catalog
