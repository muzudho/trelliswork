# from .features.canvas import render_canvas
# from .features.cards import render_all_cards
# と書くのがめんどくさいので、
# from .features import render_canvas, render_all_cards と書けばよいようにする仕組み
from .canvas import render_canvas
from .cards import render_all_cards
from .line_tapes import render_all_line_tapes
from .pillars import render_all_pillar_rugs
from .rectangles import render_all_rectangles
from .ruler import render_ruler
from .shadow_of_cards import render_shadow_of_all_cards
from .shadow_of_line_tapes import render_shadow_of_all_line_tapes
from .shadow_of_terminals import render_shadow_of_all_terminals
from .terminals import render_all_terminals
from .xl_text import render_all_xl_texts
