"""Internal logging helpers."""

from __future__ import annotations

import builtins
import logging
import os


_LOGGER = logging.getLogger("pptchartengine")


def debug_print(*args, **kwargs) -> None:
    """Route legacy diagnostic prints to debug logging by default.

    Set ``PPTCHARTENGINE_DEBUG_STDOUT=1`` to restore the old stdout diagnostics
    while investigating OOXML output locally.
    """

    if os.environ.get("PPTCHARTENGINE_DEBUG_STDOUT"):
        builtins.print(*args, **kwargs)
        return

    file = kwargs.get("file")
    if file is not None:
        builtins.print(*args, **kwargs)
        return

    sep = kwargs.get("sep", " ")
    end = kwargs.get("end", "\n")
    message = sep.join(str(arg) for arg in args)
    if end not in ("", "\n"):
        message += end
    _LOGGER.debug(message)
