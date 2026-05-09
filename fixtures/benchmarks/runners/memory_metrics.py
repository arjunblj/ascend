from __future__ import annotations

import resource
import sys

import psutil


def peak_rss_bytes() -> int:
    value = int(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    if sys.platform == "darwin":
        return value
    return value * 1024


def current_rss_bytes() -> int:
    return int(psutil.Process().memory_info().rss)


def memory_baseline() -> dict[str, int]:
    rss = current_rss_bytes()
    peak = peak_rss_bytes()
    return {"rss": rss, "peak": peak}


def sample_with_memory(
    duration_ms: float, before: dict[str, int] | None = None
) -> dict[str, float | int]:
    rss = current_rss_bytes()
    peak = peak_rss_bytes()
    sample = {"durationMs": duration_ms, "peakRssBytes": peak, "rssAfterBytes": rss}
    if before is not None:
        sample["rssDeltaBytes"] = max(0, rss - before["rss"])
        sample["retainedRssDeltaBytes"] = sample["rssDeltaBytes"]
    return sample
