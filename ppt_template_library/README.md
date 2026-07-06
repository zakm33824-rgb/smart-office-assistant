# PPT Template Library

This package provides the resource-library layer for the PPT generator.

## What it does

- Maintains a large source inventory for open and free presentation resources
- Builds a scaffolded template catalog from the source inventory
- Seeds a reusable slide/page library and a component library
- Assigns conservative license, usage, and quality hints
- Stores source/template/page/component metadata in SQLite
- Generates preview images and contact-sheet previews
- Supports deduplication, routing, export, and incremental refresh

## Directory layout

```text
ppt_template_library/
  templates/
  components/
  preview/
  metadata/
  database/
  commercial_allowed/
  personal_use_only/
  license_uncertain/
  premium_quality/
  industry/
  scenario/
  style/
  color/
  layout/
  logs/
  source_records/
```

## Bootstrap

```bash
python scripts/bootstrap_template_library.py
```

This command creates the directory structure, initializes the SQLite database, writes the source catalog, builds template blueprints, seeds the slide/component libraries, and exports JSON indexes.

## Update flow

1. Refresh the source catalog
2. Build template scaffolds from the source inventory
3. Seed slide and component catalogs
4. Generate preview assets
5. Export searchable JSON indexes
6. Deduplicate by file hash, structure hash, and preview hash
7. Route records into the appropriate license and quality buckets

## Search helpers

The library exposes search and filter helpers for:

- Sources
- Template blueprints
- Slides / page components
- Component snippets

These functions are used by the Streamlit page and can be reused by the PPT generator later.
