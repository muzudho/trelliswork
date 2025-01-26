# from .shared_modules.rectangle import VarRectangle
# from .shared_modules.share import Share
# と書くのがめんどくさいので、
# from .shared_modules import VarRectangle, Share と書けばよいようにする仕組み

from .depth140.canvas import Canvas
from .depth150.card import Card
from .depth110.color_system import ColorSystem
from .depth110.file_path import FilePath
from .depth120.innings_pitched import InningsPitched
from .depth150.pillar import Pillar
from .depth110.point import Point
from .depth130.var_rectangle import VarRectangle
from .depth150.terminal import Terminal
from .depth110.xl_alignment import XlAlignment
from .depth130.xl_font import XlFont
from .depth120.var_color import VarColor
from .depth110.share import Share
from .depth110.web_safe_color import WebSafeColor
