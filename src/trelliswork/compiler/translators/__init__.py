# from .compiler.translators.auto_shadow import AutoShadow
# from .compiler.translators.auto_split_pillar import AutoSplitSegmentByPillar
# と書くのがめんどくさいので、
# from .compiler.translators import AutoShadow, AutoSplitSegmentByPillar と書けばよいようにする仕組み
from .auto_shadow import AutoShadow
from .auto_split_pillar import AutoSplitSegmentByPillar
from .imports import Imports
from .resolve_alias_of_color import ResolveAliasOfColor
from .resolve_var_bounds import ResolveVarBounds
