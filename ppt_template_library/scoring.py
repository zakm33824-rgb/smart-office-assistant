from __future__ import annotations

from .models import SourceEntry


LICENSE_SCORE = {
    "MIT": 24,
    "Apache-2.0": 24,
    "CC0-1.0": 24,
    "Public-Domain": 24,
    "CC-BY": 18,
    "CC-BY-SA": 15,
    "LGPL-2.1": 12,
    "GPL-3.0": 10,
    "CC-BY-NC-SA": 4,
    "review required": 8,
    "site_terms": 8,
    "unknown": 5,
}


SOURCE_TYPE_SCORE = {
    "repository": 28,
    "theme_gallery": 22,
    "collection": 20,
    "category_page": 18,
    "topic_index": 14,
    "engine_home": 20,
}


ACCESS_SCORE = {"high": 16, "medium": 11, "low": 6}
COMMERCIAL_SCORE = {"allowed": 18, "restricted": 5, "uncertain": 9, "n/a": 0}
MOD_SCORE = {"allowed": 10, "restricted": 0, "uncertain": 4, "n/a": 0}


def score_source(entry: SourceEntry) -> int:
    score = 0
    score += SOURCE_TYPE_SCORE.get(entry.source_type, 10)
    score += LICENSE_SCORE.get(entry.license_hint, 5)
    score += COMMERCIAL_SCORE.get(entry.commercial_use_hint, 5)
    score += MOD_SCORE.get(entry.modification_allowed_hint, 4)
    score += ACCESS_SCORE.get(entry.programmatic_access, 5)
    score += min(max(entry.quality_hint, 0), 30)
    return max(0, min(100, score))


def score_template(
    design_score: int = 0,
    usability_score: int = 0,
    quality_score: int = 0,
    richness_bonus: int = 0,
) -> int:
    score = design_score * 0.35 + usability_score * 0.25 + quality_score * 0.3 + richness_bonus * 0.1
    return max(0, min(100, round(score)))

