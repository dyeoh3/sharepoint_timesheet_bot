"""
Australian public-holiday helper.

Uses the ``holidays`` library to check whether a given date is a public
holiday in a configured Australian state (default: NSW).
"""

from __future__ import annotations

from datetime import date

import holidays as _holidays


def get_au_holidays(state: str = "NSW", year: int | None = None) -> _holidays.HolidayBase:
    """
    Return an Australian holiday calendar for *state* and *year*.

    Args:
        state: Two-letter Australian state code — ``NSW``, ``VIC``, ``QLD``,
               ``SA``, ``WA``, ``TAS``, ``NT``, ``ACT``.
        year:  Calendar year.  Defaults to the current year.
    """
    if year is None:
        year = date.today().year
    return _holidays.Australia(state=state.upper(), years=year)


def is_public_holiday(d: date, state: str = "NSW") -> bool:
    """Return ``True`` if *d* is a public holiday in *state*."""
    cal = get_au_holidays(state, d.year)
    return d in cal


def holiday_name(d: date, state: str = "NSW") -> str | None:
    """Return the holiday name if *d* is a public holiday, else ``None``."""
    cal = get_au_holidays(state, d.year)
    return cal.get(d)


def get_holidays_in_range(
    start: date, end: date, state: str = "NSW"
) -> dict[date, str]:
    """
    Return ``{date: name}`` for every public holiday between *start* and
    *end* (inclusive).
    """
    years = set(range(start.year, end.year + 1))
    cal = _holidays.Australia(state=state.upper(), years=years)
    return {d: name for d, name in sorted(cal.items()) if start <= d <= end}
