from __future__ import annotations

from .models import SourceEntry
from .scoring import score_source


def _slugify(value: str) -> str:
    text = value.lower()
    out = []
    previous_dash = False
    for ch in text:
        if ch.isalnum():
            out.append(ch)
            previous_dash = False
        else:
            if not previous_dash:
                out.append("-")
                previous_dash = True
    slug = "".join(out).strip("-")
    return slug or "source"


def _make_source(
    name: str,
    url: str,
    source_type: str,
    license_hint: str,
    commercial_use_hint: str,
    modification_allowed_hint: str,
    download_method: str,
    programmatic_access: str,
    quality_hint: int,
    language: str,
    tags: tuple[str, ...],
    notes: str = "",
) -> SourceEntry:
    source_id = _slugify(url)
    entry = SourceEntry(
        source_id=source_id,
        name=name,
        source_url=url,
        source_website=url.split("/")[2],
        source_type=source_type,
        license_hint=license_hint,
        commercial_use_hint=commercial_use_hint,
        modification_allowed_hint=modification_allowed_hint,
        download_method=download_method,
        programmatic_access=programmatic_access,
        quality_hint=quality_hint,
        language=language,
        tags=tags,
        notes=notes or f"Seed source for {name}.",
    )
    return SourceEntry(**{**entry.__dict__, "score": score_source(entry)})


def build_seed_sources() -> list[SourceEntry]:
    entries: list[SourceEntry] = []

    repo_sources = [
        ("GBIF PowerPoint Template", "https://github.com/gbif/ppt-template", "repository", "review required", "restricted", "allowed", "git", "high", 84, "en", ("biology", "research", "institutional"), "Open template repository with toolkit slides and branded layouts."),
        ("Slidev", "https://github.com/slidevjs/slidev", "repository", "MIT", "allowed", "allowed", "git", "high", 92, "en", ("markdown", "developer", "web"), "Markdown-first slide engine with themes and PPTX export."),
        ("Marp", "https://github.com/marp-team/marp", "repository", "MIT", "allowed", "allowed", "git", "high", 91, "en", ("markdown", "presentation", "theme"), "Open-source Markdown presentation ecosystem."),
        ("reveal.js", "https://github.com/hakimel/reveal.js/", "repository", "MIT", "allowed", "allowed", "git", "high", 93, "en", ("html", "web", "interactive"), "Popular HTML presentation framework."),
        ("pptx-template", "https://github.com/m3dev/pptx-template", "repository", "Apache-2.0", "allowed", "allowed", "git", "high", 89, "en", ("python", "pptx", "data"), "Template-driven PPTX builder for dynamic content."),
        ("Chuk MCP PPTX", "https://github.com/IBM/chuk-mcp-pptx", "repository", "Apache-2.0", "allowed", "allowed", "git", "high", 89, "en", ("mcp", "pptx", "components"), "Template-first PowerPoint system with component registry."),
        ("pptx-masters", "https://github.com/anotb/pptx-masters", "repository", "MIT", "allowed", "allowed", "git", "high", 86, "en", ("pptx", "masters", "corporate"), "Extracts corporate slide masters into code."),
        ("Presentation Studio MCP", "https://github.com/vicmaster/presentation-studio-mcp", "repository", "NOASSERTION", "uncertain", "allowed", "git", "high", 82, "en", ("mcp", "renderer", "slides"), "Local production-grade PPTX engine exposed over MCP."),
        ("Claude Skill Slide Kit", "https://github.com/PHY041/claude-skill-slide-kit", "repository", "MIT", "allowed", "allowed", "git", "high", 88, "en", ("python", "pitch-deck", "academic"), "Native editable .pptx deck builder with charts and diagrams."),
        ("Metropolis / mtheme", "https://github.com/matze/mtheme", "repository", "CC-BY-SA-4.0", "allowed", "allowed", "git", "high", 84, "en", ("beamer", "latex", "modern"), "Modern LaTeX Beamer theme."),
        ("Moloch", "https://github.com/jolars/moloch", "repository", "CC-BY-SA-4.0", "allowed", "allowed", "git", "high", 84, "en", ("beamer", "latex", "minimal"), "Minimalist Beamer theme with dark/light support."),
        ("Pure Minimalistic Beamer", "https://github.com/kai-tub/latex-beamer-pure-minimalistic/", "repository", "GPL-3.0", "allowed", "allowed", "git", "high", 79, "en", ("beamer", "latex", "minimal"), "Truly minimal Beamer template for 16:9 and 4:3."),
        ("FancyBeamer", "https://github.com/SEatUPB/FancyBeamer", "repository", "CC0-1.0", "allowed", "allowed", "git", "high", 83, "en", ("beamer", "university", "animations"), "University-oriented Beamer template with animations."),
        ("Fake Beamer", "https://github.com/sbryngelson/Fake-Beamer", "repository", "MIT", "allowed", "allowed", "git", "high", 78, "en", ("beamer", "powerpoint", "classic"), "PowerPoint template that imitates classic Beamer styling."),
        ("Marp Slides Template", "https://github.com/codebytes/marp-slides-template", "repository", "MIT", "allowed", "allowed", "git", "high", 82, "en", ("marp", "github-pages", "template"), "Minimal Marp template with GitHub Pages workflow."),
        ("Marp Slides Template", "https://github.com/christophdb/marp-slides-template", "repository", "MIT", "allowed", "allowed", "git", "high", 82, "en", ("marp", "github-pages", "template"), "Marp starter with pages workflow and custom themes."),
        ("Awesome Marp Template", "https://github.com/yKicchan/awesome-marp-template", "repository", "MIT", "allowed", "allowed", "git", "high", 84, "en", ("marp", "utility", "theme"), "Feature-rich Marp template with utility classes."),
        ("Slide Templates", "https://github.com/Echtzeitsysteme/slide-templates", "repository", "review required", "uncertain", "allowed", "git", "high", 76, "en", ("university", "pptx", "latex"), "PowerPoint and LaTeX templates for student presentations."),
        ("THU-PPT-Theme", "https://github.com/atomiechen/THU-PPT-Theme", "repository", "CC-BY-NC-SA-4.0", "restricted", "allowed", "git", "high", 78, "zh", ("academic", "thesis", "university"), "Chinese university-themed PPT template."),
        ("HHU Slides Template", "https://github.com/Mizera-Mondo/HHU-Slides-Template", "repository", "LGPL-2.1", "allowed", "allowed", "git", "high", 80, "en", ("university", "pptx", "beamer"), "Unofficial PPT and Beamer template."),
        ("ShanghaiTech PPT Template", "https://github.com/HypoxanthineOvO/SHTU-PPT-Template", "repository", "review required", "uncertain", "allowed", "git", "high", 81, "zh", ("university", "marp", "beamer"), "Three template systems maintained in one repository."),
        ("GordenPPTSkill", "https://github.com/GordenSun/GordenPPTSkill", "repository", "personal use only", "restricted", "allowed", "git", "high", 92, "zh", ("ai", "pptx", "templates"), "AI-friendly PPT builder with 17 hand-polished Chinese PPTX templates."),
        ("SAP Presentation Template", "https://github.com/jung-thomas/sap-presentation-template", "repository", "review required", "uncertain", "allowed", "git", "high", 90, "en", ("corporate", "slidev", "brand"), "GitHub template repo for SAP-branded presentations."),
        ("NYPL Marp Template", "https://github.com/NYPL/nypl-marp-template", "repository", "review required", "uncertain", "allowed", "git", "high", 83, "en", ("marp", "institutional", "brand"), "Marp template themed with New York Public Library guidelines."),
        ("Quarto RevealJS Clean", "https://github.com/grantmcdermott/quarto-revealjs-clean", "repository", "MIT", "allowed", "allowed", "git", "high", 84, "en", ("quarto", "revealjs", "minimal"), "Minimalist and elegant Quarto Reveal.js theme."),
        ("RevealJS Template", "https://github.com/ktkv419/revealjs-template", "repository", "review required", "uncertain", "allowed", "git", "high", 76, "en", ("revealjs", "boilerplate", "theme"), "Reveal.js boilerplate with Tokyo Night theming."),
        ("Tidy RevealJS", "https://github.com/julie-ng/tidy-revealjs", "repository", "review required", "uncertain", "allowed", "git", "high", 74, "en", ("revealjs", "template", "automation"), "Reveal.js theme and build system."),
        ("Reveal JS Themes", "https://github.com/AnneTee/reveal-js-themes", "repository", "review required", "uncertain", "allowed", "git", "high", 73, "en", ("revealjs", "themes", "gallery"), "Collection of reveal.js templates and custom themes."),
        ("Awesome Presentations", "https://github.com/aspose-slides/Awesome-Presentations", "repository", "MIT", "allowed", "allowed", "git", "high", 86, "en", ("presentation", "tools", "resources"), "Curated list of tools and resources for presentations."),
        ("Primer Presentations", "https://primer.style/presentations/", "repository", "review required", "uncertain", "allowed", "git", "high", 85, "en", ("github", "brand", "guidelines"), "GitHub brand presentation system and assets."),
        ("ppt_themes", "https://github.com/hornhuang/ppt_themes", "repository", "MIT", "allowed", "allowed", "git", "high", 84, "zh", ("template", "pptx", "collection"), "Large personal PPT template collection."),
        ("PPT Agent", "https://github.com/icip-cas/pptagent", "repository", "review required", "uncertain", "allowed", "git", "high", 88, "en", ("agent", "pptx", "generation"), "Agentic framework for reflective PPT generation."),
        ("PPTX Masters", "https://github.com/anotb/pptx-masters", "repository", "MIT", "allowed", "allowed", "git", "high", 86, "en", ("pptx", "master", "code"), "Slide master extraction for AI-friendly slide generation."),
        ("PPT Master", "https://github.com/hugohe3/ppt-master", "repository", "MIT", "allowed", "allowed", "git", "high", 96, "en", ("ai", "pptx", "editable"), "AI system that generates natively editable PowerPoint decks."),
        ("McKinsey PPTX", "https://github.com/seulee26/mckinsey-pptx", "repository", "MIT", "allowed", "allowed", "git", "high", 94, "ko", ("consulting", "mckinsey", "pptx"), "Consulting-style PPTX generator with 40 slide templates."),
        ("BIT PPT Template", "https://github.com/yang-kun-long/bit-ppt-template", "repository", "MIT", "allowed", "allowed", "git", "high", 90, "zh", ("university", "pptx", "cli"), "YAML-driven editable PPTX generator with web UI and MCP entrypoints."),
        ("GitHub Social Preview PPTX Template", "https://github.com/europanite/github_social_preview_pptx_template", "repository", "Apache-2.0", "allowed", "allowed", "git", "high", 86, "en", ("github", "social-image", "pptx"), "PPTX template for GitHub social preview images."),
        ("pptxtemplates", "https://github.com/EmilHvitfeldt/pptxtemplates", "repository", "NOASSERTION", "uncertain", "allowed", "git", "high", 72, "en", ("r", "rmarkdown", "template"), "R Markdown PowerPoint template collection."),
    ]

    site_sources = [
        ("PresentationGO", "https://www.presentationgo.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 86, "en", ("business", "diagram", "timeline"), "Large free PPT and Google Slides library."),
        ("Slidesgo Business", "https://slidesgo.com/business", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 84, "en", ("business", "pitch-deck", "report"), "Business category with free and premium templates."),
        ("Slidesgo Finances", "https://www.slidesgo.com/finances", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 82, "en", ("finance", "report", "data"), "Finance category with editable templates."),
        ("SlidesCarnival", "https://www.slidescarnival.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 82, "en", ("free", "education", "business"), "Free PowerPoint and Google Slides templates."),
        ("PresentationMagazine", "https://www.presentationmagazine.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 77, "en", ("backgrounds", "business", "theme"), "Long-running free PPT template collection."),
        ("Microsoft Templates", "https://powerpoint.cloud.microsoft/create/en/templates/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 84, "en", ("microsoft", "business", "education"), "Microsoft template gallery."),
        ("SlideStart Corporate", "https://www.slidestart.com/templates/corporate-template", "category_page", "commercial", "allowed", "allowed", "manual", "medium", 90, "en", ("corporate", "dashboard", "strategy"), "Large professional slide library with advanced filters."),
        ("SlideGeeks Open Source", "https://www.slidegeeks.com/open-source-powerpoint-presentation-templates", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 72, "en", ("template", "collection", "open-source"), "Open-source themed collection page."),
        ("Canva Business Templates", "https://www.canva.com/presentations/templates/business/", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 86, "en", ("business", "presentation", "design"), "Business presentation category in Canva."),
        ("Canva Finance Templates", "https://www.canva.com/presentations/templates/finance/", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 84, "en", ("finance", "presentation", "design"), "Finance presentation category in Canva."),
        ("LibreTemplates Financial Theme", "https://libretemplates.com/en/presentations/business/financial-theme/102", "category_page", "Libretemplates license", "allowed", "allowed", "manual", "medium", 80, "en", ("finance", "libreoffice", "pptx"), "PPTX and ODP theme for business/financial presentations."),
        ("EaTemp Finance Template", "https://eatemp.com/finance-ppt-presentation-template/", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 83, "en", ("finance", "google-slides", "figma"), "Free finance template with multi-format support."),
        ("AiPPT Elegant Business Template", "https://www.aippt.com/templates/elegant-corporate-business-powerpoint-template", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 85, "en", ("business", "minimal", "red"), "Free business template on AiPPT."),
        ("SlideTeam GitHub", "https://www.slideteam.net/top-10-github-powerpoint-presentation-templates", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 68, "en", ("github", "business", "templates"), "Curated article on GitHub-themed templates."),
        ("SlideTeam Repository", "https://www.slidegeeks.com/ppt/repository", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 66, "en", ("repository", "project", "meeting"), "Repository-themed slide collection."),
        ("SlideTeam Git and GitHub", "https://www.slidegeeks.com/ppt/git", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 66, "en", ("git", "github", "tech"), "Git-themed slide collection."),
        ("优品PPT", "https://www.ypppt.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 88, "zh", ("free", "business", "domestic"), "Large Chinese free PPT template site."),
        ("第一PPT", "https://www.1ppt.com/moban/ppt_moban_1.html", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 87, "zh", ("free", "business", "education"), "Popular Chinese free PPT template site."),
        ("51PPT模板网", "https://www.51pptmoban.com/ppt/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 86, "zh", ("free", "business", "education"), "Chinese PPT template library with frequent updates."),
        ("LFPPT", "https://m.lfppt.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 85, "zh", ("free", "education", "business"), "Chinese PPT template and courseware site."),
        ("Canva PPT (zh-cn)", "https://www.canva.cn/create/ppt-slides/", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 86, "zh", ("business", "design", "presentation"), "Chinese Canva presentation landing page."),
        ("Slidesgo", "https://slidesgo.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 88, "en", ("business", "education", "simple"), "Global presentation template library."),
        ("SlidesMania", "https://slidesmania.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 87, "en", ("google-slides", "pptx", "free"), "100% free Google Slides and PPTX templates."),
        ("Visme Presentations", "https://www.visme.co/templates/presentations/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 84, "en", ("presentation", "business", "education"), "Presentation template gallery with business and education layouts."),
        ("Figma Presentations", "https://www.figma.com/community/presentations", "collection", "community_terms", "uncertain", "allowed", "manual", "medium", 82, "en", ("presentation", "pitch-deck", "design"), "Figma community presentation templates."),
        ("SlideEgg Free PowerPoint Templates", "https://www.slideegg.com/free-powerpoint-templates", "collection", "site_terms", "allowed", "allowed", "manual", "medium", 88, "en", ("business", "education", "data"), "Large free PowerPoint template catalog with editable vectors."),
        ("Genially Presentation Templates", "https://genially.com/templates/presentations/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 84, "en", ("interactive", "presentation", "animation"), "Interactive presentation template gallery."),
        ("Canva PPT (ja)", "https://www.canva.com/ja_jp/presentations/templates/", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 85, "ja", ("business", "education", "presentation"), "Japanese Canva presentation template page."),
        ("Canva PPT (ko)", "https://www.canva.com/ko_kr/presentations/templates/", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 85, "ko", ("business", "education", "presentation"), "Korean Canva presentation template page."),
        ("SlidesCarnival Japan", "https://www.slidescarnival.com/ja/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 83, "ja", ("free", "education", "business"), "Japanese SlidesCarnival collection."),
        ("Slidesgo Japan", "https://slidesgo.com/ja/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 83, "ja", ("business", "education", "presentation"), "Japanese Slidesgo page."),
        ("Slidesgo Korea", "https://slidesgo.com/ko/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 83, "ko", ("business", "education", "presentation"), "Korean Slidesgo page."),
        ("Adobe Express KR Presentations", "https://www.adobe.com/kr/express/templates/presentation", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 82, "ko", ("presentation", "business", "creative"), "Korean Adobe Express presentation templates."),
        ("MiriCanvas Presentations", "https://www.miricanvas.com/ko/template/presentation", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 84, "ko", ("presentation", "business", "education"), "Korean presentation template marketplace."),
        ("Design AC Presentation", "https://www.design-ac.net/free-template/presentation", "category_page", "site_terms", "uncertain", "allowed", "manual", "medium", 81, "ja", ("presentation", "business", "education"), "Japanese free presentation template page."),
        ("Raku-Pre", "https://raku-pre.com/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 80, "ja", ("presentation", "proposal", "business"), "Japanese PowerPoint template distribution site."),
        ("Slidesdocs Korea", "https://slidesdocs.com/ko/", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 82, "ko", ("business", "report", "presentation"), "Korean template library with editable slides."),
        ("Magnific PPT Templates", "https://www.magnific.com/kr/free-photos-vectors/ppt-%ED%85%9C%ED%94%8C%EB%A6%BF", "collection", "site_terms", "uncertain", "allowed", "manual", "medium", 79, "ko", ("free", "vector", "presentation"), "Korean free PPT vector/template search page."),
    ]

    topic_slugs = [
        "pptx",
        "ppt",
        "powerpoint",
        "presentation",
        "presentation-template",
        "slides",
        "slide-deck",
        "pitch-deck",
        "template",
        "theme",
        "business-plan",
        "annual-report",
        "quarterly-report",
        "financial-report",
        "sales-report",
        "marketing-plan",
        "marketing",
        "company-profile",
        "consulting",
        "project-proposal",
        "project-management",
        "roadmap",
        "timeline",
        "infographic",
        "dashboard",
        "data-visualization",
        "report",
        "case-study",
        "startup",
        "investor-pitch",
        "education",
        "thesis",
        "defense",
        "lecture",
        "resume",
        "portfolio",
        "medical",
        "healthcare",
        "science",
        "engineering",
        "technology",
        "art",
        "photography",
        "fashion",
        "travel",
        "tourism",
        "sports",
        "government",
        "nonprofit",
        "environment",
        "agriculture",
        "logistics",
        "manufacturing",
        "real-estate",
        "construction",
        "banking",
        "insurance",
        "finance",
        "ai",
        "cloud",
        "cybersecurity",
        "ecommerce",
        "retail",
        "media",
        "film",
        "music",
        "game",
        "workshop",
        "conference",
        "training",
        "strategy",
        "operations",
        "hr",
        "sales",
        "branding",
        "proposal",
        "meeting",
        "research",
        "poster",
        "newsletter",
        "kpi",
        "swot",
        "gantt-chart",
        "diagram",
        "workflow",
        "chart",
        "table",
        "icons",
        "corporate-design",
        "powerpoint-free",
        "powerpoint-presentations",
        "google-slides",
        "google-slides-templates",
        "presentation-templates",
        "slidesgo",
        "slideegg",
        "slidesmania",
        "figma-presentations",
        "genially",
    ]

    for slug in topic_slugs:
        url = f"https://github.com/topics/{slug}"
        entries.append(
            _make_source(
                name=f"GitHub topic: {slug}",
                url=url,
                source_type="topic_index",
                license_hint="unknown",
                commercial_use_hint="uncertain",
                modification_allowed_hint="n/a",
                download_method="topic_discovery",
                programmatic_access="high",
                quality_hint=58,
                language="en",
                tags=tuple(filter(None, slug.replace("-", " ").split())),
                notes="GitHub topic discovery page for open repositories and template projects.",
            )
        )

    for record in repo_sources + site_sources:
        entries.append(_make_source(*record))

    unique: dict[str, SourceEntry] = {}
    for entry in entries:
        unique[entry.source_url] = entry
    ordered = sorted(unique.values(), key=lambda item: (-item.score, item.name.lower()))
    return ordered


def build_template_catalog() -> list[dict[str, str]]:
    catalog: list[dict[str, str]] = []
    for idx, entry in enumerate(build_seed_sources(), start=1):
        catalog.append(
            {
                "template_id": f"seed-{idx:04d}",
                "template_name": entry.name,
                "source_url": entry.source_url,
                "source_website": entry.source_website,
                "license": entry.license_hint,
                "commercial_use": entry.commercial_use_hint,
                "modification_allowed": entry.modification_allowed_hint,
                "download_method": entry.download_method,
                "score": str(entry.score),
                "tags": ",".join(entry.tags),
            }
        )
    return catalog


