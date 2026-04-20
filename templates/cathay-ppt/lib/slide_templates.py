"""
Cathay PPT Template — Pre-built Slide Templates
=================================================
10 ready-to-use slide template functions. Each creates a complete slide
layout that's validated for text fit.

Usage:
    import sys
    sys.path.insert(0, os.path.expanduser('~/.claude/skills/cathay-ppt-template/lib'))
    from text_engine import *
    from slide_templates import *
"""

from text_engine import (
    # Constants
    TEMPLATE, CATHAY_RED, CATHAY_GOLD, CATHAY_LTGOLD, CATHAY_ACCENT,
    CATHAY_BLACK, CATHAY_WHITE, CATHAY_GREY, CATHAY_LTGREY,
    CATHAY_COLORS,
    CL, CT, CB, CW, CH, GAP_H, GAP_V,
    HALF, THIRD, QUARTER, ONE_THIRD, TWO_THIRDS, ONE_QUARTER, THREE_QUARTER,
    X1, X2_HALF, X2_Q34, X2_T23, X2_MID, X3_RIGHT,
    Y1, Y2_HALF, Y2_MID, Y3_BOT, ROW_HALF, ROW_THIRD,
    SOURCE_Y_MM, SOURCE_BOX_HEIGHT_MM, SOURCE_FONT_PT, DEFAULT_FONT_SIZE,
    CONTENT_BOTTOM_MM,
    ICON_FINANCIAL, ICON_INSIGHT, ICON_RISK, ICON_CATALYST, ICON_ACTION,
    # Functions
    create_content_slide, set_title_with_conclusion, set_title,
    add_subtitle, add_source_footer, add_page_number,
    setup_text_frame, format_paragraph, set_run_font, add_mixed_text,
    set_square_bullet, add_bullet_content,
    smart_textbox, smart_table, add_table,
    add_callout_box, add_flow_box, add_arrow, add_down_arrow,
    add_color_block, add_kpi_row,
    safe_textbox, safe_chart_insert, setup_chart_style,
    calc_text_height, get_char_width,
    validate_and_fix,
    Presentation, Mm, Pt, PP_ALIGN, MSO_SHAPE, RGBColor,
    MARGIN_ALL, MSO_AUTO_SIZE,
)

from pptx.oxml.ns import qn
from lxml import etree


# ============================================================================
# TEMPLATE 1: KPI Dashboard
# ============================================================================

def template_kpi_dashboard(prs, title, subtitle, kpis, bullets, source):
    """KPI row (3-6 boxes) + bullet content below.

    Args:
        prs: Presentation object
        title: slide title string
        subtitle: gold subtitle string (or conclusion for title bar)
        kpis: list of (value, label) tuples, e.g. [("$1.2B", "Revenue"), ...]
        bullets: list of (text, level) tuples for add_bullet_content
        source: source text for footer

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    # KPI row at top of content zone
    body_y = add_kpi_row(slide, kpis, y_mm=CT)

    # Bullet content below KPIs
    smart_textbox(slide, X1, body_y, CW, bullets,
                  max_bottom_mm=CB, start_font=10, min_font=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 2: Value Chain Flow
# ============================================================================

def template_value_chain_flow(prs, title, subtitle, chain_items, table_data, source):
    """Horizontal flow chart (N boxes + arrows) + comparison table below.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        chain_items: list of str for flow boxes
        table_data: 2D list for table (first row = header)
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    n = len(chain_items)
    if n == 0:
        n = 1

    # Flow region: top 30mm of content zone
    flow_h = 20
    flow_y = CT + 2
    arrow_w = 8
    total_arrow_w = arrow_w * (n - 1) if n > 1 else 0
    total_gap = 3 * (n - 1) if n > 1 else 0
    box_w = (CW - total_arrow_w - total_gap) / n

    cur_x = X1
    for i, item_text in enumerate(chain_items):
        add_flow_box(slide, cur_x, flow_y, box_w, flow_h, item_text,
                     bg_color=CATHAY_RED, text_color=CATHAY_WHITE, font_size=9)
        cur_x += box_w

        if i < n - 1:
            add_arrow(slide, cur_x + 1, flow_y + flow_h / 2 - 3,
                      w_mm=arrow_w - 2, h_mm=6, color=CATHAY_GOLD)
            cur_x += arrow_w + 3 - box_w + box_w  # simplified: next box position
            cur_x = X1 + (i + 1) * (box_w + arrow_w + 3) - (arrow_w + 3) + (arrow_w + 3)

    # Recalculate x positions cleanly
    cur_x = X1
    for i, item_text in enumerate(chain_items):
        # Already drawn above; this block is for accurate arrow positioning
        pass

    # Table below flow chart
    table_top = flow_y + flow_h + GAP_V + 5
    if table_data and len(table_data) > 0:
        smart_table(slide, table_data, left_mm=CL, top_mm=table_top,
                    width_mm=CW, max_bottom_mm=CB, font_size=9)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 3: Chart + Analysis
# ============================================================================

def template_chart_plus_analysis(prs, title, subtitle, chart_path, analysis_items,
                                 source, chart_side='left'):
    """Chart on one side + text analysis on the other.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        chart_path: path to chart PNG file
        analysis_items: list of (text, level) tuples
        source: source text
        chart_side: 'left' or 'right'

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    chart_w = HALF
    text_w = HALF

    if chart_side == 'left':
        chart_x = X1
        text_x = X2_HALF
    else:
        chart_x = X2_HALF
        text_x = X1

    # Insert chart (width-only, preserve aspect)
    import os
    if chart_path and os.path.exists(chart_path):
        safe_chart_insert(slide, chart_path, x_mm=chart_x, y_mm=CT, w_mm=chart_w)

    # Analysis text on the other side
    smart_textbox(slide, text_x, CT, text_w, analysis_items,
                  max_bottom_mm=CB, start_font=10, min_font=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 4: Comparison Matrix
# ============================================================================

def template_comparison_matrix(prs, title, subtitle, callouts, table_data,
                               conclusion, source):
    """Callout boxes at top + full comparison table + conclusion.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        callouts: list of (value, label) for top callout boxes
        table_data: 2D list for comparison table
        conclusion: conclusion text (str) or list of (text, level) tuples
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    cur_y = CT

    # Callout boxes at top (if provided)
    if callouts:
        n = len(callouts)
        bw = (CW - GAP_H * (n - 1)) / n
        for i, (val, lbl) in enumerate(callouts):
            x = X1 + i * (bw + GAP_H)
            add_callout_box(slide, x, cur_y, bw, 18, val, lbl,
                            bg_color=CATHAY_RED, text_color=CATHAY_WHITE)
        cur_y += 18 + GAP_V

    # Comparison table
    table_bottom = CB
    if conclusion:
        table_bottom = CB - 25  # reserve space for conclusion

    if table_data and len(table_data) > 0:
        _, tbl_bottom = smart_table(slide, table_data, left_mm=CL, top_mm=cur_y,
                                    width_mm=CW, max_bottom_mm=table_bottom, font_size=9)
        cur_y = tbl_bottom + GAP_V

    # Conclusion
    if conclusion:
        if isinstance(conclusion, str):
            conclusion_items = [(conclusion, 1)]
        else:
            conclusion_items = conclusion
        smart_textbox(slide, X1, cur_y, CW, conclusion_items,
                      max_bottom_mm=CB, start_font=10, min_font=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 5: Two-Column Analysis
# ============================================================================

def template_two_column_analysis(prs, title, subtitle, left_items, right_items,
                                 bottom_kpis, source):
    """Two-column analysis with optional bottom KPI row.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        left_items: list of (text, level) tuples for left column
        right_items: list of (text, level) tuples for right column
        bottom_kpis: list of (value, label) for bottom KPI row, or None
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    # Calculate content bottom based on whether we have bottom KPIs
    content_bottom = CB
    if bottom_kpis:
        content_bottom = CB - 30  # reserve 30mm for KPI row

    # Left column
    smart_textbox(slide, X1, CT, HALF, left_items,
                  max_bottom_mm=content_bottom, start_font=10, min_font=8)

    # Right column
    smart_textbox(slide, X2_HALF, CT, HALF, right_items,
                  max_bottom_mm=content_bottom, start_font=10, min_font=8)

    # Bottom KPI row
    if bottom_kpis:
        kpi_y = content_bottom + GAP_V
        add_kpi_row(slide, bottom_kpis, y_mm=kpi_y)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 6: Sidebar Case Study
# ============================================================================

def template_sidebar_case_study(prs, title, subtitle, sidebar_metrics,
                                main_items, bottom_table, source):
    """1/4 dark sidebar + 3/4 main content + optional bottom table.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        sidebar_metrics: list of (value, label) for sidebar, or list of (text, level)
        main_items: list of (text, level) tuples for main content
        bottom_table: 2D list for bottom table, or None
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    main_bottom = CB
    if bottom_table:
        main_bottom = CB - 45  # reserve space for table

    # Dark sidebar (1/4)
    sidebar = add_color_block(slide, X1, CT, ONE_QUARTER, main_bottom - CT,
                              color=CATHAY_RED)

    # Sidebar content
    sidebar_box, sidebar_tf = safe_textbox(slide, X1 + 2, CT + 3,
                                           ONE_QUARTER - 4,
                                           max_bottom_mm=main_bottom - 3)
    if sidebar_metrics:
        # Check if metrics are (value, label) tuples or (text, level) tuples
        if sidebar_metrics and len(sidebar_metrics[0]) == 2:
            first_item = sidebar_metrics[0]
            if isinstance(first_item[1], int):
                # (text, level) format
                add_bullet_content(sidebar_tf, sidebar_metrics,
                                   size_pt=10, color_rgb=CATHAY_WHITE)
            else:
                # (value, label) format
                items = []
                for val, lbl in sidebar_metrics:
                    items.append((val, 0))
                    items.append((lbl, 1))
                add_bullet_content(sidebar_tf, items,
                                   size_pt=10, color_rgb=CATHAY_WHITE)

    # Main content (3/4)
    smart_textbox(slide, X2_Q34, CT, THREE_QUARTER, main_items,
                  max_bottom_mm=main_bottom, start_font=10, min_font=8)

    # Bottom table
    if bottom_table and len(bottom_table) > 0:
        table_top = main_bottom + GAP_V
        smart_table(slide, bottom_table, left_mm=CL, top_mm=table_top,
                    width_mm=CW, max_bottom_mm=CB, font_size=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 7: Three-Column Compare
# ============================================================================

def template_three_column_compare(prs, title, subtitle, col1, col2, col3, source):
    """Three equal columns with color-accented headers.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        col1: dict with 'header' (str) and 'items' (list of (text, level))
        col2: dict with 'header' (str) and 'items' (list of (text, level))
        col3: dict with 'header' (str) and 'items' (list of (text, level))
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    columns = [col1, col2, col3]
    col_x_positions = [X1, X2_MID, X3_RIGHT]
    col_colors = [CATHAY_RED, CATHAY_GOLD, CATHAY_GREY]
    header_h = 12

    for i, (col, x_pos, col_color) in enumerate(
            zip(columns, col_x_positions, col_colors)):
        # Color header bar
        add_color_block(slide, x_pos, CT, THIRD, header_h, color=col_color)

        # Header text
        hdr_box, hdr_tf = safe_textbox(slide, x_pos + 1, CT + 1,
                                       THIRD - 2, h_mm=header_h - 2)
        p = hdr_tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        header_text = col.get('header', f'Column {i+1}') if isinstance(col, dict) else f'Column {i+1}'
        set_run_font(run, header_text, size_pt=11, bold=True,
                     color_rgb=CATHAY_WHITE)

        # Column body
        body_y = CT + header_h + GAP_V
        items = col.get('items', []) if isinstance(col, dict) else col
        if items:
            smart_textbox(slide, x_pos, body_y, THIRD, items,
                          max_bottom_mm=CB, start_font=10, min_font=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 8: Stacked Cases
# ============================================================================

def template_stacked_cases(prs, title, subtitle, case1, case2, source):
    """Two cases stacked vertically with separator.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        case1: dict with 'header' (str), 'items' (list of (text, level)),
               optional 'color' (RGBColor)
        case2: dict with 'header' (str), 'items' (list of (text, level)),
               optional 'color' (RGBColor)
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    half_h = ROW_HALF
    separator_h = 1

    # Case 1 (top half)
    case1_header = case1.get('header', 'Case 1') if isinstance(case1, dict) else 'Case 1'
    case1_items = case1.get('items', []) if isinstance(case1, dict) else case1
    case1_color = case1.get('color', CATHAY_RED) if isinstance(case1, dict) else CATHAY_RED

    # Header bar for case 1
    add_color_block(slide, X1, CT, CW, 10, color=case1_color)
    hdr1_box, hdr1_tf = safe_textbox(slide, X1 + 2, CT + 1, CW - 4, h_mm=8)
    p = hdr1_tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    set_run_font(run, case1_header, size_pt=11, bold=True, color_rgb=CATHAY_WHITE)

    # Case 1 content
    case1_body_y = CT + 10 + 2
    case1_bottom = CT + half_h - separator_h
    if case1_items:
        smart_textbox(slide, X1, case1_body_y, CW, case1_items,
                      max_bottom_mm=case1_bottom, start_font=10, min_font=8)

    # Separator line
    sep_y = CT + half_h
    add_color_block(slide, X1, sep_y, CW, separator_h, color=CATHAY_LTGREY)

    # Case 2 (bottom half)
    case2_header = case2.get('header', 'Case 2') if isinstance(case2, dict) else 'Case 2'
    case2_items = case2.get('items', []) if isinstance(case2, dict) else case2
    case2_color = case2.get('color', CATHAY_GOLD) if isinstance(case2, dict) else CATHAY_GOLD

    case2_top = sep_y + separator_h + GAP_V
    add_color_block(slide, X1, case2_top, CW, 10, color=case2_color)
    hdr2_box, hdr2_tf = safe_textbox(slide, X1 + 2, case2_top + 1, CW - 4, h_mm=8)
    p2 = hdr2_tf.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    run2 = p2.add_run()
    set_run_font(run2, case2_header, size_pt=11, bold=True, color_rgb=CATHAY_WHITE)

    case2_body_y = case2_top + 10 + 2
    if case2_items:
        smart_textbox(slide, X1, case2_body_y, CW, case2_items,
                      max_bottom_mm=CB, start_font=10, min_font=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 9: Risk Cards
# ============================================================================

def template_risk_cards(prs, title, subtitle, risks, source):
    """5 risk cards with gradient colors.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        risks: list of dicts, each with 'title' (str), 'description' (str),
               optional 'severity' ('high'/'medium'/'low')
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    n = len(risks)
    if n == 0:
        add_source_footer(slide, source)
        return slide

    # Arrange cards: up to 5 across, or 2 rows if more
    if n <= 5:
        card_w = (CW - GAP_H * (n - 1)) / n
        card_h = CB - CT - 5
        cards_per_row = n
        rows = 1
    else:
        cards_per_row = min(n, 5)
        rows = 2
        card_w = (CW - GAP_H * (cards_per_row - 1)) / cards_per_row
        card_h = (CB - CT - GAP_V) / 2 - 2

    # Color gradient for severity
    severity_colors = {
        'high':   RGBColor(0xE6, 0x00, 0x00),  # red
        'medium': RGBColor(0xC8, 0xA4, 0x15),  # gold
        'low':    RGBColor(0x80, 0x80, 0x80),  # grey
    }
    default_gradient = [CATHAY_RED, CATHAY_ACCENT, CATHAY_GOLD, CATHAY_GREY,
                        CATHAY_LTGOLD]

    for idx, risk in enumerate(risks):
        row = idx // cards_per_row
        col = idx % cards_per_row
        x = X1 + col * (card_w + GAP_H)
        y = CT + row * (card_h + GAP_V)

        risk_title = risk.get('title', f'Risk {idx+1}') if isinstance(risk, dict) else str(risk)
        risk_desc = risk.get('description', '') if isinstance(risk, dict) else ''
        severity = risk.get('severity', None) if isinstance(risk, dict) else None

        if severity and severity in severity_colors:
            card_color = severity_colors[severity]
        else:
            card_color = default_gradient[idx % len(default_gradient)]

        # Card header
        header_h = 12
        add_color_block(slide, x, y, card_w, header_h, color=card_color)
        hdr_box, hdr_tf = safe_textbox(slide, x + 1, y + 1,
                                       card_w - 2, h_mm=header_h - 2)
        p = hdr_tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_mixed_text(p, risk_title, size_pt=9, bold=True, color_rgb=CATHAY_WHITE)

        # Card body (light background)
        body_y = y + header_h
        body_h = card_h - header_h
        add_color_block(slide, x, body_y, card_w, body_h,
                        color=RGBColor(0xF5, 0xF5, 0xF5))

        if risk_desc:
            desc_items = [(risk_desc, 1)]
            smart_textbox(slide, x + 1, body_y + 2, card_w - 2, desc_items,
                          max_bottom_mm=y + card_h - 1, start_font=9, min_font=7)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 10: Action Plan
# ============================================================================

def template_action_plan(prs, title, subtitle, tiers, timeline, conclusion, source):
    """Priority funnel + timeline + conclusion box.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        tiers: list of dicts with 'label' (str), 'items' (list of str),
               optional 'color' (RGBColor). Arranged as funnel rows.
        timeline: list of (phase, description) tuples for timeline bar
        conclusion: str or list of (text, level) tuples
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    cur_y = CT

    # Priority tiers (funnel)
    tier_colors = [CATHAY_RED, CATHAY_GOLD, CATHAY_GREY, CATHAY_LTGREY]
    if tiers:
        n_tiers = len(tiers)
        tier_h = min(18, (CB - CT - 50) / max(n_tiers, 1))

        for i, tier in enumerate(tiers):
            tier_label = tier.get('label', f'Tier {i+1}') if isinstance(tier, dict) else str(tier)
            tier_items = tier.get('items', []) if isinstance(tier, dict) else []
            tier_color = tier.get('color', tier_colors[i % len(tier_colors)]) if isinstance(tier, dict) else tier_colors[i % len(tier_colors)]

            # Funnel narrowing effect
            inset = i * 8
            box_w = CW - inset * 2
            box_x = X1 + inset

            add_color_block(slide, box_x, cur_y, box_w, tier_h, color=tier_color)

            # Tier text
            tier_box, tier_tf = safe_textbox(slide, box_x + 2, cur_y + 1,
                                             box_w - 4, h_mm=tier_h - 2)
            p = tier_tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            text_color = CATHAY_WHITE if i < 2 else CATHAY_BLACK
            run = p.add_run()
            set_run_font(run, tier_label, size_pt=10, bold=True, color_rgb=text_color)

            if tier_items:
                items_text = " | ".join(tier_items) if isinstance(tier_items, list) else str(tier_items)
                run2 = p.add_run()
                set_run_font(run2, f"  {items_text}", size_pt=9, color_rgb=text_color)

            cur_y += tier_h + 2

    cur_y += GAP_V

    # Timeline bar
    if timeline:
        n_phases = len(timeline)
        phase_w = (CW - GAP_H * (n_phases - 1)) / n_phases
        timeline_h = 22

        for i, (phase, desc) in enumerate(timeline):
            px = X1 + i * (phase_w + GAP_H)

            # Phase box
            add_flow_box(slide, px, cur_y, phase_w, 10, phase,
                         bg_color=CATHAY_GOLD, text_color=CATHAY_WHITE, font_size=8)

            # Description below
            desc_box, desc_tf = safe_textbox(slide, px, cur_y + 11,
                                             phase_w, h_mm=timeline_h - 11)
            p = desc_tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            add_mixed_text(p, desc, size_pt=8, color_rgb=CATHAY_BLACK)

            # Arrow between phases
            if i < n_phases - 1:
                arrow_x = px + phase_w + 1
                add_arrow(slide, arrow_x, cur_y + 2, w_mm=GAP_H - 2, h_mm=6,
                          color=CATHAY_GOLD)

        cur_y += timeline_h + GAP_V

    # Conclusion box
    if conclusion:
        if isinstance(conclusion, str):
            conclusion_items = [(conclusion, 1)]
        else:
            conclusion_items = conclusion

        # Add a subtle background for the conclusion
        remaining_h = CB - cur_y
        if remaining_h > 8:
            add_color_block(slide, X1, cur_y, CW, remaining_h,
                            color=RGBColor(0xF8, 0xF0, 0xE0))  # warm light bg
            smart_textbox(slide, X1 + 2, cur_y + 2, CW - 4, conclusion_items,
                          max_bottom_mm=CB - 2, start_font=10, min_font=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 11: Donut Chart (Matplotlib PNG + native shapes)
# ============================================================================

def template_donut_chart(prs, title, subtitle, segments, insight_bullets, source):
    """Native donut chart (matplotlib PNG) + insight panel.

    McKinsey pattern: clean donut visual on the left, key insight bullets on
    the right. Legend below the donut uses native colored squares.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        segments: list of (label, value, color_hex) — values are percentages
            that sum to 100.
            e.g. [("AI/ML", 45, "800000"), ("Cloud", 30, "C8A415"),
                  ("Other", 25, "808080")]
        insight_bullets: list of (text, level) tuples for right panel
        source: source text

    Returns:
        slide object
    """
    import tempfile
    import os

    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    default_hex = ['800000', 'C8A415', '808080', 'E60000']

    # --- Left side: donut chart image (55% of CW) ---
    chart_zone_w = CW * 0.55
    chart_img_w = chart_zone_w - 10  # padding

    # Build color list and values
    values = [s[1] for s in segments]
    colors_hex = [s[2] if len(s) > 2 and s[2] else default_hex[i % len(default_hex)]
                  for i, s in enumerate(segments)]
    colors_plt = [f'#{c}' for c in colors_hex]

    # Render donut via matplotlib
    tmp_path = None
    try:
        setup_chart_style()
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt

        fig, ax = plt.subplots(figsize=(3.2, 3.2), dpi=150)
        wedges, _ = ax.pie(
            values, colors=colors_plt, startangle=90,
            wedgeprops={'width': 0.35, 'edgecolor': 'white', 'linewidth': 2})
        ax.set_aspect('equal')
        plt.tight_layout(pad=0)

        tmp_path = tempfile.mktemp(suffix='.png')
        fig.savefig(tmp_path, transparent=True, bbox_inches='tight',
                    pad_inches=0.05)
        plt.close(fig)

        # Insert donut image
        donut_x = X1 + (chart_zone_w - chart_img_w) / 2
        chart_bottom = safe_chart_insert(slide, tmp_path,
                                         x_mm=donut_x, y_mm=CT + 2,
                                         w_mm=chart_img_w)
    except ImportError:
        # matplotlib not available — add placeholder text
        chart_bottom = CT + 80
        placeholder_box, placeholder_tf = safe_textbox(
            slide, X1, CT + 20, chart_zone_w, h_mm=30)
        p = placeholder_tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run_font(p.add_run(), "[Donut Chart]",
                     size_pt=14, bold=True, color_rgb=CATHAY_GREY)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    # --- Legend below donut: colored squares + label + pct ---
    legend_y = chart_bottom + 3
    legend_x = X1 + 5
    sq_size = 5
    legend_gap = 3

    for i, seg in enumerate(segments):
        label = seg[0]
        val = seg[1]
        color_hex_str = seg[2] if len(seg) > 2 and seg[2] else default_hex[i % len(default_hex)]
        r = int(color_hex_str[0:2], 16)
        g = int(color_hex_str[2:4], 16)
        b = int(color_hex_str[4:6], 16)

        # Colored square
        add_color_block(slide, legend_x, legend_y, sq_size, sq_size,
                        color=RGBColor(r, g, b))

        # Label + percentage text
        lbl_box, lbl_tf = safe_textbox(slide, legend_x + sq_size + 2,
                                       legend_y, chart_zone_w - sq_size - 15,
                                       h_mm=sq_size + 1)
        lp = lbl_tf.paragraphs[0]
        lp.alignment = PP_ALIGN.LEFT
        set_run_font(lp.add_run(), f"{label}  {val}%",
                     size_pt=8, color_rgb=CATHAY_BLACK)

        legend_y += sq_size + legend_gap

    # --- Right side: insight bullets (40% of CW) ---
    insight_x = X1 + CW * 0.60
    insight_w = CW * 0.40
    smart_textbox(slide, insight_x, CT, insight_w, insight_bullets,
                  max_bottom_mm=CB, start_font=10, min_font=8)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 12: Before / After Comparison
# ============================================================================

def template_before_after(prs, title, subtitle, before_items, after_items, source):
    """Side-by-side before/after comparison with visual distinction.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        before_items: list of (text, level) tuples for "Before" column
        after_items: list of (text, level) tuples for "After" column
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    col_w = HALF
    header_h = 14

    # --- Left column: Before / 现状 ---
    left_x = X1

    # Grey header bar
    add_color_block(slide, left_x, CT, col_w, header_h,
                    color=RGBColor(0xD9, 0xD9, 0xD9))
    hdr_left, hdr_left_tf = safe_textbox(slide, left_x + 2, CT + 1,
                                         col_w - 4, h_mm=header_h - 2)
    p_left = hdr_left_tf.paragraphs[0]
    p_left.alignment = PP_ALIGN.CENTER
    add_mixed_text(p_left, "Before / \u73b0\u72b6", size_pt=11, bold=True,
                   color_rgb=CATHAY_BLACK)

    # Before bullets
    before_body_y = CT + header_h + GAP_V
    smart_textbox(slide, left_x, before_body_y, col_w, before_items,
                  max_bottom_mm=CB, start_font=10, min_font=8)

    # --- Right column: After / 优化后 ---
    right_x = X2_HALF

    # Light gold header bar
    add_color_block(slide, right_x, CT, col_w, header_h,
                    color=RGBColor(0xFF, 0xF8, 0xE1))
    hdr_right, hdr_right_tf = safe_textbox(slide, right_x + 2, CT + 1,
                                           col_w - 4, h_mm=header_h - 2)
    p_right = hdr_right_tf.paragraphs[0]
    p_right.alignment = PP_ALIGN.CENTER
    add_mixed_text(p_right, "After / \u4f18\u5316\u540e", size_pt=11, bold=True,
                   color_rgb=CATHAY_RED)

    # After bullets
    after_body_y = CT + header_h + GAP_V
    smart_textbox(slide, right_x, after_body_y, col_w, after_items,
                  max_bottom_mm=CB, start_font=10, min_font=8)

    # --- Thin red arrow between columns at mid-height ---
    arrow_y = CT + (CB - CT) / 2 - 4
    arrow_x = left_x + col_w + 0.5
    arrow_w = GAP_H - 1
    add_arrow(slide, arrow_x, arrow_y, w_mm=arrow_w, h_mm=8,
              color=CATHAY_RED)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 13: Funnel Chart
# ============================================================================

def template_funnel(prs, title, subtitle, stages, source):
    """Top-to-bottom funnel chart with progressive narrowing.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        stages: list of (label, value_str, description) tuples, widest first
            e.g. [("Total Market", "$430B", "Global data center market"), ...]
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    n = len(stages)
    if n == 0:
        add_source_footer(slide, source)
        return slide

    funnel_colors = [CATHAY_RED, CATHAY_GOLD, CATHAY_GREY, CATHAY_ACCENT,
                     CATHAY_LTGOLD]

    # Layout: each stage is a rounded rect, progressively narrower
    # Left ~65% for the funnel bars, right ~30% for descriptions
    funnel_zone_w = CW * 0.60
    desc_zone_x = X1 + CW * 0.65
    desc_zone_w = CW * 0.35

    stage_gap = 4
    available_h = CB - CT - 5
    stage_h = min(22, (available_h - stage_gap * (n - 1)) / n)

    for i, (label, value_str, description) in enumerate(stages):
        # Progressive narrowing: first stage full width, last ~40%
        shrink_ratio = 1.0 - (i * 0.6 / max(n - 1, 1))
        bar_w = funnel_zone_w * shrink_ratio
        inset = (funnel_zone_w - bar_w) / 2
        bar_x = X1 + inset
        bar_y = CT + i * (stage_h + stage_gap)

        bar_color = funnel_colors[i % len(funnel_colors)]

        # Funnel bar (empty text — we overlay text separately for control)
        add_flow_box(slide, bar_x, bar_y, bar_w, stage_h, "",
                     bg_color=bar_color, text_color=CATHAY_WHITE,
                     font_size=9)

        # Value + label inside bar
        inner_box, inner_tf = safe_textbox(slide, bar_x + 2, bar_y + 1,
                                           bar_w - 4, h_mm=stage_h - 2)
        pi = inner_tf.paragraphs[0]
        pi.alignment = PP_ALIGN.CENTER
        text_color = CATHAY_WHITE if i < 3 else CATHAY_BLACK
        set_run_font(pi.add_run(), value_str, size_pt=12, bold=True,
                     color_rgb=text_color)
        run_label = pi.add_run()
        set_run_font(run_label, f"  {label}", size_pt=9,
                     color_rgb=text_color)

        # Description to the right
        desc_box, desc_tf = safe_textbox(slide, desc_zone_x, bar_y + 1,
                                         desc_zone_w, h_mm=stage_h - 2)
        pd = desc_tf.paragraphs[0]
        pd.alignment = PP_ALIGN.LEFT
        add_mixed_text(pd, description, size_pt=9, color_rgb=CATHAY_BLACK)

        # Down arrow between stages (not after last)
        if i < n - 1:
            arrow_cx = X1 + funnel_zone_w / 2 - 3
            arrow_y_pos = bar_y + stage_h
            add_down_arrow(slide, arrow_cx, arrow_y_pos,
                           w_mm=6, h_mm=stage_gap, color=CATHAY_GOLD)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 14: Enhanced SWOT Matrix
# ============================================================================

def template_swot(prs, title, subtitle, strengths, weaknesses,
                  opportunities, threats, source):
    """Color-coded SWOT 2x2 matrix with Cathay branding.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        strengths: list of (text, level) tuples
        weaknesses: list of (text, level) tuples
        opportunities: list of (text, level) tuples
        threats: list of (text, level) tuples
        source: source text

    Returns:
        slide object
    """
    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    quad_gap = 3
    header_h = 16
    quad_w = (CW - quad_gap) / 2
    quad_h = (CB - CT - quad_gap) / 2

    # Quadrant definitions: (label, bg_color, header_color, items, col, row)
    quadrants = [
        ("Strengths",      RGBColor(0xFF, 0xF8, 0xE1), CATHAY_RED,
         strengths, 0, 0),
        ("Weaknesses",     RGBColor(0xF5, 0xF5, 0xF5), CATHAY_GREY,
         weaknesses, 1, 0),
        ("Opportunities",  RGBColor(0xFF, 0xF8, 0xE1), CATHAY_GOLD,
         opportunities, 0, 1),
        ("Threats",        RGBColor(0xFF, 0xE0, 0xE0), CATHAY_ACCENT,
         threats, 1, 1),
    ]

    for (label, bg_color, hdr_color, items, col, row) in quadrants:
        qx = X1 + col * (quad_w + quad_gap)
        qy = CT + row * (quad_h + quad_gap)

        # Background fill for entire quadrant
        add_color_block(slide, qx, qy, quad_w, quad_h, color=bg_color)

        # Header bar
        add_color_block(slide, qx, qy, quad_w, header_h, color=hdr_color)
        hdr_box, hdr_tf = safe_textbox(slide, qx + 2, qy + 1,
                                       quad_w - 4, h_mm=header_h - 2)
        hp = hdr_tf.paragraphs[0]
        hp.alignment = PP_ALIGN.LEFT
        add_mixed_text(hp, label, size_pt=11, bold=True,
                       color_rgb=CATHAY_WHITE)

        # Bullet content below header
        body_y = qy + header_h + 2
        body_bottom = qy + quad_h - 2
        if items:
            smart_textbox(slide, qx + 2, body_y, quad_w - 4, items,
                          max_bottom_mm=body_bottom, start_font=9,
                          min_font=7)

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 15: Waterfall Chart
# ============================================================================

def template_waterfall(prs, title, subtitle, items, source):
    """Waterfall chart showing incremental changes via matplotlib.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        items: list of (label, value, is_total) tuples
            e.g. [("Base Rev", 100, False), ("AI Growth", 30, False),
                  ("Total", 130, True)]
            Positive values in gold, negative in accent red, totals in dark red
        source: source text

    Returns:
        slide object
    """
    import tempfile
    import os

    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    labels = [it[0] for it in items]
    values = [it[1] for it in items]
    is_totals = [it[2] if len(it) > 2 else False for it in items]

    tmp_path = None
    try:
        setup_chart_style()
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import numpy as np

        n = len(items)
        # Calculate running cumulative for waterfall positioning
        cumulative = 0
        bottoms = []
        bar_vals = []
        bar_colors = []

        for i in range(n):
            label_i = labels[i]
            val = values[i]
            is_total = is_totals[i]
            if is_total:
                bottoms.append(0)
                bar_vals.append(cumulative)
                bar_colors.append('#800000')  # dark red for totals
            else:
                if val >= 0:
                    bottoms.append(cumulative)
                    bar_vals.append(val)
                    bar_colors.append('#C8A415')  # gold for positive
                    cumulative += val
                else:
                    cumulative += val
                    bottoms.append(cumulative)
                    bar_vals.append(abs(val))
                    bar_colors.append('#E60000')  # accent red for negative

        fig, ax = plt.subplots(figsize=(7, 3.5), dpi=150)

        x_pos = np.arange(n)
        bars = ax.bar(x_pos, bar_vals, bottom=bottoms, color=bar_colors,
                      edgecolor='white', linewidth=0.8, width=0.6)

        # Add value labels on bars
        max_val = max(abs(v) for v in values) if values else 1
        for i, (bar, bv, bt) in enumerate(zip(bars, bar_vals, bottoms)):
            val = values[i]
            label_y = bt + bv + (max_val * 0.02)
            display_text = f'{val:g}' if is_totals[i] else f'{val:+g}'
            ax.text(bar.get_x() + bar.get_width() / 2, label_y,
                    display_text,
                    ha='center', va='bottom', fontsize=8, fontweight='bold')

        # Connector lines between bars (thin grey dashes)
        for i in range(n - 1):
            if not is_totals[i]:
                top_of_current = bottoms[i] + bar_vals[i]
                ax.plot([x_pos[i] + 0.3, x_pos[i + 1] - 0.3],
                        [top_of_current, top_of_current],
                        color='#808080', linewidth=0.5, linestyle='--')

        ax.set_xticks(x_pos)
        ax.set_xticklabels(labels, fontsize=8, rotation=30, ha='right')
        ax.set_ylabel('')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_linewidth(0.5)
        ax.spines['bottom'].set_linewidth(0.5)
        ax.tick_params(axis='y', labelsize=8)
        ax.grid(axis='y', alpha=0.3)
        plt.tight_layout()

        tmp_path = tempfile.mktemp(suffix='.png')
        fig.savefig(tmp_path, transparent=False, bbox_inches='tight',
                    pad_inches=0.1, facecolor='white')
        plt.close(fig)

        safe_chart_insert(slide, tmp_path, x_mm=X1, y_mm=CT, w_mm=CW)

    except ImportError:
        # matplotlib not available — fallback to text summary
        fallback_items = []
        for i in range(len(items)):
            lbl = labels[i]
            val = values[i]
            is_total = is_totals[i]
            prefix = "TOTAL: " if is_total else ("+" if val >= 0 else "")
            fallback_items.append((f"{lbl}: {prefix}{val}", 1))
        smart_textbox(slide, X1, CT, CW, fallback_items,
                      max_bottom_mm=CB, start_font=10, min_font=8)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    add_source_footer(slide, source)
    return slide


# ============================================================================
# TEMPLATE 16: Stakeholder Map
# ============================================================================

def template_stakeholder_map(prs, title, subtitle, stakeholders, source):
    """Visual stakeholder / relationship map with center + surrounding nodes.

    Args:
        prs: Presentation object
        title: slide title
        subtitle: conclusion text
        stakeholders: list of (name, role, position) tuples
            position: 'center', 'top', 'bottom', 'left', 'right',
                      'top-left', 'top-right', 'bottom-left', 'bottom-right'
        source: source text

    Returns:
        slide object
    """
    import math

    slide = create_content_slide(prs, topic=title, conclusion=subtitle)

    # Center of the content zone
    map_cx = CL + CW / 2
    map_cy = CT + CH / 2

    # Position coordinate map (absolute mm)
    pos_map = {
        'center':       (map_cx, map_cy),
        'top':          (map_cx, CT + 20),
        'bottom':       (map_cx, CB - 20),
        'left':         (CL + 30, map_cy),
        'right':        (CL + CW - 30, map_cy),
        'top-left':     (CL + 40, CT + 25),
        'top-right':    (CL + CW - 40, CT + 25),
        'bottom-left':  (CL + 40, CB - 25),
        'bottom-right': (CL + CW - 40, CB - 25),
    }

    # Separate center stakeholder from others
    center_sh = None
    outer_shs = []
    for (name, role, position) in stakeholders:
        if position == 'center':
            center_sh = (name, role, position)
        else:
            outer_shs.append((name, role, position))

    # If no explicit center, use first stakeholder
    if center_sh is None and stakeholders:
        center_sh = (stakeholders[0][0], stakeholders[0][1], 'center')
        outer_shs = list(stakeholders[1:])

    center_d = 25  # diameter mm for center node
    outer_d = 18   # diameter mm for outer nodes

    # Draw connecting lines first (behind circles)
    if center_sh:
        c_x, c_y = pos_map['center']
        for (name, role, position) in outer_shs:
            pos_key = position if position in pos_map else 'right'
            o_x, o_y = pos_map[pos_key]

            dx = o_x - c_x
            dy = o_y - c_y

            # Line as thin rectangle (IRON RULE: no connectors)
            if abs(dx) >= abs(dy):
                # Mostly horizontal — draw horizontal line at midpoint Y
                mid_y = (c_y + o_y) / 2
                line_x = min(c_x, o_x)
                add_color_block(slide, line_x, mid_y - 0.5,
                                abs(dx), 1, color=CATHAY_LTGREY)
            else:
                # Mostly vertical — draw vertical line at midpoint X
                mid_x = (c_x + o_x) / 2
                line_y = min(c_y, o_y)
                add_color_block(slide, mid_x - 0.5, line_y,
                                1, abs(dy), color=CATHAY_LTGREY)

    # Draw center node (dark red circle)
    if center_sh:
        c_name, c_role, _ = center_sh
        c_x, c_y = pos_map['center']

        sh = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Mm(c_x - center_d / 2), Mm(c_y - center_d / 2),
            Mm(center_d), Mm(center_d))
        sh.fill.solid()
        sh.fill.fore_color.rgb = CATHAY_RED
        sh.line.fill.background()

        tf = sh.text_frame
        setup_text_frame(tf)
        p_name = tf.paragraphs[0]
        p_name.alignment = PP_ALIGN.CENTER
        set_run_font(p_name.add_run(), c_name, size_pt=9, bold=True,
                     color_rgb=CATHAY_WHITE)
        if c_role:
            p_role = tf.add_paragraph()
            p_role.alignment = PP_ALIGN.CENTER
            set_run_font(p_role.add_run(), c_role, size_pt=7,
                         color_rgb=CATHAY_WHITE)

    # Draw outer nodes (gold circles)
    for (name, role, position) in outer_shs:
        pos_key = position if position in pos_map else 'right'
        o_x, o_y = pos_map[pos_key]

        sh = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Mm(o_x - outer_d / 2), Mm(o_y - outer_d / 2),
            Mm(outer_d), Mm(outer_d))
        sh.fill.solid()
        sh.fill.fore_color.rgb = CATHAY_GOLD
        sh.line.fill.background()

        tf = sh.text_frame
        setup_text_frame(tf)
        p_name = tf.paragraphs[0]
        p_name.alignment = PP_ALIGN.CENTER
        set_run_font(p_name.add_run(), name, size_pt=8, bold=True,
                     color_rgb=CATHAY_WHITE)
        if role:
            p_role = tf.add_paragraph()
            p_role.alignment = PP_ALIGN.CENTER
            set_run_font(p_role.add_run(), role, size_pt=6,
                         color_rgb=CATHAY_WHITE)

    add_source_footer(slide, source)
    return slide
