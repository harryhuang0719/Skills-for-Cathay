"""
Cathay PPT Template v2 — Convenience Re-exports
=================================================

Usage:
    import sys, os
    sys.path.insert(0, os.path.expanduser("~/.claude/skills/cathay-ppt-template/lib"))

    from cathay_ppt import *        # everything
    from cathay_ppt import cards     # only Card building blocks
    from cathay_ppt import fonts     # set_run_font, add_mixed_text
    from cathay_ppt import slides    # create_content_slide, set_title_with_conclusion
"""

from constants import *

from fonts import (
    CJK_CHAR_WIDTH, LATIN_CHAR_WIDTH,
    get_char_width, set_run_font, add_mixed_text,
)

from text_layout import (
    setup_text_frame, format_paragraph,
    set_square_bullet, add_bullet_content, add_multi_text,
    calc_text_height, calc_textframe_height,
    smart_textbox,
)

from tables import add_table, smart_table

from charts import (
    setup_chart_style, safe_chart_insert, insert_chart_image,
    cathay_bar_chart, cathay_line_chart, cathay_waterfall_chart,
)

from slides import (
    create_cover_slide, create_content_slide,
    set_title, set_slide_title, set_title_with_conclusion,
    add_subtitle, add_source_footer, add_page_number,
)

from safe_layout import safe_textbox

from validation import (
    validate_and_fix, save_with_validation,
    validate_no_overlap, validate_text_fit,
    qc_presentation, export_to_pdf,
)

from merge import (
    _clean_shape, full_cleanup,
    merge_slides, reorder_slides, clear_slide,
)

from elements import (
    HeaderBar, ContentPanel, KpiStrip, Card, MetricRow,
    add_callout_box, add_flow_box, add_color_block, add_kpi_row,
    add_arrow, add_down_arrow, add_progress_bar,
    add_section_marker, auto_assign_icons,
)

from slide_templates import *
from qc_automation import *
from data_driven import *
from svg_embed import *
