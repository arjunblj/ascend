from __future__ import annotations

import resource
import sys


def peak_rss_bytes() -> int:
    value = int(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    if sys.platform == "darwin":
        return value
    return value * 1024


def sample_with_memory(duration_ms: float) -> dict[str, float | int]:
    peak = peak_rss_bytes()
    return {"durationMs": duration_ms, "peakRssBytes": peak, "rssAfterBytes": peak}
