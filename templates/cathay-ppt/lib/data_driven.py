"""
Data-Driven Slide Generation for Cathay PPT Template.

Separates content from layout:
  - DataRegistry: single source of truth for all data points
  - Slide specs: declarative dict format for slide content
  - build_deck_from_specs(): specs list -> complete deck
  - render_spec(): render a single spec into a presentation
"""

import os
import json
import copy
import io

from pptx import Presentation
from pptx.util import Mm, Pt
from pptx.oxml.ns import qn

# ---------------------------------------------------------------------------
# Import shared engine from sibling modules
# ---------------------------------------------------------------------------
_LIB_DIR = os.path.dirname(os.path.abspath(__file__))
import sys
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

from text_engine import (
    validate_and_fix,
    save_with_validation,
    merge_slides,
    setup_text_frame,
    set_run_font,
    add_bullet_content,
    add_source_footer,
    add_kpi_row,
    add_table,
    safe_textbox,
    safe_chart_insert,
    set_title_with_conclusion,
    add_mixed_text,
    format_paragraph,
    TEMPLATE,
    CL, CT, CW, CH, GAP_H, GAP_V,
    HALF, THIRD, ONE_THIRD, TWO_THIRDS, ONE_QUARTER, THREE_QUARTER,
    X1, X2_HALF, X2_T23, X2_Q34, X2_MID, X3_RIGHT,
    CATHAY_RED, CATHAY_GOLD, CATHAY_WHITE, CATHAY_BLACK, CATHAY_GREY,
    CONTENT_BOTTOM_MM,
)

from slide_templates import (
    template_kpi_dashboard,
    template_value_chain_flow,
    template_chart_plus_analysis,
    template_comparison_matrix,
    template_two_column_analysis,
    template_sidebar_case_study,
    template_three_column_compare,
    template_stacked_cases,
    template_risk_cards,
    template_action_plan,
    template_donut_chart,
    template_before_after,
    template_funnel,
    template_swot,
    template_waterfall,
    template_stakeholder_map,
)

from qc_automation import full_qc_pipeline


# ═══════════════════════════════════════════════════════════════════════════
# 1. Slide Spec Format
# ═══════════════════════════════════════════════════════════════════════════

# Example slide spec — each slide in a deck is described by one of these dicts:
#
# SLIDE_SPEC = {
#     'template': 'kpi_dashboard',           # which template to use
#     'title': '核心观点',                      # slide title (topic)
#     'subtitle': 'AIDC是未来7年最确定的基建超级周期',  # conclusion for title bar
#     'data': {
#         'kpis': [('$430B', '全球DC收入(2026E)'), ...],
#         'bullets': [('AI算力需求驱动', 0), ('具体分析...', 1), ...],
#     },
#     'source': 'IEA, Gartner, Cathay Analysis',
#     'page_num': 2,
# }


# ═══════════════════════════════════════════════════════════════════════════
# 2. Data Registry — Single Source of Truth
# ═══════════════════════════════════════════════════════════════════════════

class DataRegistry:
    """Single source of truth for all data points used in a deck.

    Stores (value, source, year) triples.  Load from / save to JSON.
    Provides get() with source tracking for audit trails.
    """

    def __init__(self, data_file=None):
        self.data = {}      # key -> value
        self.sources = {}   # key -> source string
        self.years = {}     # key -> year (optional)
        if data_file:
            self.load(data_file)

    # --- Core accessors ---

    def set(self, key, value, source, year=None):
        """Register a data point with its source."""
        self.data[key] = value
        self.sources[key] = source
        if year is not None:
            self.years[key] = year

    def get(self, key, default=None):
        """Get a data point value."""
        return self.data.get(key, default)

    def get_with_source(self, key):
        """Get (value, source) tuple.  Returns (None, None) if key missing."""
        if key not in self.data:
            return (None, None)
        return (self.data[key], self.sources.get(key, "unknown"))

    def get_source(self, key):
        """Get just the source string for a data point."""
        return self.sources.get(key)

    def keys(self):
        """Return all registered keys."""
        return list(self.data.keys())

    def __contains__(self, key):
        return key in self.data

    def __len__(self):
        return len(self.data)

    # --- Persistence ---

    def save(self, path):
        """Save registry to JSON file."""
        payload = {
            "data": self.data,
            "sources": self.sources,
            "years": self.years,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f"DataRegistry saved: {len(self.data)} entries -> {path}")

    def load(self, path):
        """Load registry from JSON file."""
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.data = payload.get("data", {})
        self.sources = payload.get("sources", {})
        self.years = payload.get("years", {})
        print(f"DataRegistry loaded: {len(self.data)} entries from {path}")

    # --- Validation ---

    def validate(self):
        """Check all data points have sources.

        Returns:
            list of keys missing sources (empty = all good)
        """
        missing = []
        for key in self.data:
            if key not in self.sources or not self.sources[key]:
                missing.append(key)
        if missing:
            print(f"DataRegistry validation: {len(missing)} keys missing sources:")
            for k in missing[:10]:
                print(f"  - {k}")
            if len(missing) > 10:
                print(f"  ... and {len(missing) - 10} more")
        else:
            print(f"DataRegistry validation: all {len(self.data)} keys have sources")
        return missing

    # --- Bulk helpers ---

    def set_many(self, entries):
        """Bulk-register data points.

        Args:
            entries: list of (key, value, source) or (key, value, source, year) tuples
        """
        for entry in entries:
            if len(entry) == 3:
                self.set(entry[0], entry[1], entry[2])
            elif len(entry) >= 4:
                self.set(entry[0], entry[1], entry[2], entry[3])

    def collect_sources(self, keys):
        """Gather unique source strings for a set of keys (for slide footer)."""
        sources = set()
        for k in keys:
            s = self.sources.get(k)
            if s:
                sources.add(s)
        return ", ".join(sorted(sources))

    def __repr__(self):
        return f"<DataRegistry: {len(self.data)} entries>"


# ═══════════════════════════════════════════════════════════════════════════
# 3. Template Router
# ═══════════════════════════════════════════════════════════════════════════

TEMPLATE_ROUTER = {
    "kpi_dashboard": template_kpi_dashboard,
    "value_chain_flow": template_value_chain_flow,
    "chart_plus_analysis": template_chart_plus_analysis,
    "comparison_matrix": template_comparison_matrix,
    "two_column_analysis": template_two_column_analysis,
    "sidebar_case_study": template_sidebar_case_study,
    "three_column_compare": template_three_column_compare,
    "stacked_cases": template_stacked_cases,
    "risk_cards": template_risk_cards,
    "action_plan": template_action_plan,
    "donut_chart": template_donut_chart,
    "before_after": template_before_after,
    "funnel": template_funnel,
    "swot": template_swot,
    "waterfall": template_waterfall,
    "stakeholder_map": template_stakeholder_map,
}


def render_spec(prs, spec, data_registry=None):
    """Render a single slide spec into a presentation.

    Args:
        prs: python-pptx Presentation object
        spec: slide spec dict with keys: template, title, subtitle, data, source, page_num
        data_registry: optional DataRegistry for source lookups

    Returns:
        slide object
    """
    template_name = spec.get("template", "kpi_dashboard")
    renderer = TEMPLATE_ROUTER.get(template_name)

    if renderer is None:
        raise ValueError(
            f"Unknown template '{template_name}'. "
            f"Available: {list(TEMPLATE_ROUTER.keys())}"
        )

    # Build kwargs — unpack spec['data'] as template-specific kwargs
    title = spec.get("title", "")
    subtitle = spec.get("subtitle", "")
    source = spec.get("source", "")
    data = spec.get("data", {})

    # If data_registry provided, resolve source references
    if data_registry is not None and not source and "data_keys" in spec:
        source = data_registry.collect_sources(spec["data_keys"])

    try:
        slide = renderer(prs, title, subtitle, **data, source=source)
    except TypeError as e:
        print(f"WARNING: Template '{template_name}' call failed: {e}")
        print(f"  Hint: check that spec['data'] keys match template function params")
        slide = None

    return slide


# ═══════════════════════════════════════════════════════════════════════════
# 4. Deck Builder
# ═══════════════════════════════════════════════════════════════════════════

def build_deck_from_specs(specs, output_path, template_path=None, data_registry=None,
                          run_qc=True):
    """Build a complete deck from a list of slide specs.

    Each spec maps to a slide_templates function via TEMPLATE_ROUTER.
    Handles all merging, image handling, and QC.

    Args:
        specs: list of slide spec dicts
        output_path: path for the final .pptx file
        template_path: path to Cathay template (default: standard template)
        data_registry: optional DataRegistry instance
        run_qc: whether to run full_qc_pipeline after building

    Returns:
        dict with keys:
            'path': output file path
            'slides': number of slides
            'qc_report': QC report dict (if run_qc=True)
            'fixes': list of auto-fix descriptions
    """
    template_path = template_path or TEMPLATE
    prs = Presentation(template_path)

    # Clear template slides
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    # Render each spec
    for i, spec in enumerate(specs):
        if "page_num" not in spec:
            spec["page_num"] = i + 1
        try:
            render_spec(prs, spec, data_registry=data_registry)
        except Exception as e:
            print(f"Error rendering slide {i + 1} ('{spec.get('template', '?')}'): {e}")
            # Add a placeholder error slide
            slide = prs.slides.add_slide(prs.slide_layouts[4])
            for shape in slide.shapes:
                if hasattr(shape, "placeholder_format") and shape.placeholder_format is not None:
                    if shape.placeholder_format.type == 1:
                        shape.text = f"ERROR: Slide {i + 1}"
            txBox = slide.shapes.add_textbox(Mm(CL), Mm(CT), Mm(CW), Mm(30))
            tf = txBox.text_frame
            setup_text_frame(tf)
            run = tf.paragraphs[0].add_run()
            set_run_font(run, f"Failed to render template '{spec.get('template', '?')}': {e}",
                         size_pt=10, color_rgb=CATHAY_RED)

    # Validate and save
    fixes = validate_and_fix(prs)
    prs.save(output_path)
    print(f"Deck saved: {len(prs.slides)} slides -> {output_path}")

    if fixes:
        print(f"  Auto-fixed {len(fixes)} issues")

    # QC
    qc_report = {}
    if run_qc:
        try:
            qc_report = full_qc_pipeline(output_path)
        except Exception as e:
            print(f"  QC pipeline error: {e}")

    return {
        "path": output_path,
        "slides": len(prs.slides),
        "qc_report": qc_report,
        "fixes": fixes,
    }


# ═══════════════════════════════════════════════════════════════════════════
# Module-level exports
# ═══════════════════════════════════════════════════════════════════════════

__all__ = [
    "DataRegistry",
    "TEMPLATE_ROUTER",
    "render_spec",
    "build_deck_from_specs",
]
