"""
Backward-compatibility shim — re-exports all symbols from split modules.

Old code that does `from text_engine import *` continues to work.
New code should use the focused modules directly.

Usage (v1 compat):
    from text_engine import *

Usage (v2 preferred):
    from fonts import set_run_font
    from slides import create_content_slide, add_source_footer
    from elements import Card, KpiStrip
"""

# Import everything the old text_engine.py used to export
from constants import *
from fonts import *
from text_layout import *
from tables import *
from charts import *
from slides import *
from safe_layout import *
from validation import *
from merge import *
from elements import *
