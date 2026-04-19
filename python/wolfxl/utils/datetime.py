"""Excel serial ↔ Python datetime conversions, openpyxl-compatible.

Reproduces openpyxl's handling of the **1900 leap-year bug**: Excel believes
1900-02-29 exists, so the 1900 epoch is offset to ``datetime(1899, 12, 30)``
and serials in ``(0, 60)`` get a +1 day correction. Pinned by
``tests/parity/test_utils_parity.py``.
"""

from __future__ import annotations

import datetime as _dt

WINDOWS_EPOCH = _dt.datetime(1899, 12, 30)
MAC_EPOCH = _dt.datetime(1904, 1, 1)

# openpyxl exports CALENDAR_WINDOWS_1900 as the WINDOWS_EPOCH datetime itself
# (re-bound at module load — see openpyxl/utils/datetime.py:17).
CALENDAR_WINDOWS_1900 = WINDOWS_EPOCH
CALENDAR_MAC_1904 = MAC_EPOCH

_SECS_PER_DAY = 86400


def from_excel(
    value: float | int | None,
    epoch: _dt.datetime = WINDOWS_EPOCH,
    timedelta: bool = False,
) -> _dt.datetime | _dt.time | _dt.timedelta | None:
    """Excel serial → ``datetime`` (or ``time`` for fractional-only values).

    Returns ``None`` when ``value`` is ``None``. Bug-for-bug compatible with
    openpyxl 3.1.x — including the 1900 leap-year adjustment.
    """
    if value is None:
        return None

    if timedelta:
        td = _dt.timedelta(days=value)
        if td.microseconds:
            td = _dt.timedelta(
                seconds=td.total_seconds() // 1,
                microseconds=round(td.microseconds, -3),
            )
        return td

    day, fraction = divmod(value, 1)
    diff = _dt.timedelta(milliseconds=round(fraction * _SECS_PER_DAY * 1000))
    if 0 <= value < 1 and diff.days == 0:
        return _days_to_time(diff)
    if 0 < value < 60 and epoch == WINDOWS_EPOCH:
        day += 1
    return epoch + _dt.timedelta(days=day) + diff


def _days_to_time(value: _dt.timedelta) -> _dt.time:
    mins, seconds = divmod(value.seconds, 60)
    hours, mins = divmod(mins, 60)
    return _dt.time(hours, mins, seconds, value.microseconds)
