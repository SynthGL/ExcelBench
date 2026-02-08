"""Performance benchmark track.

This package measures speed (wall/CPU time) and best-effort memory for each
library across each feature+operation in the manifest.
"""

from excelbench.perf.renderer import render_perf_results
from excelbench.perf.runner import run_perf

__all__ = ["run_perf", "render_perf_results"]
