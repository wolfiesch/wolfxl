"""openpyxl-compatible ``dataframe_to_rows`` helper.

Accepts a pandas DataFrame and yields rows suitable for ``ws.append()``.
Pandas is imported lazily so that ``import wolfxl.utils.dataframe`` works
without pandas installed - only calling ``dataframe_to_rows()`` triggers
the import.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    import pandas as pd  # pragma: no cover - type-only


def dataframe_to_rows(
    df: pd.DataFrame,
    index: bool = True,
    header: bool = True,
) -> Iterator[list[Any]]:
    """Yield rows from a pandas DataFrame, matching openpyxl's helper.

    - ``header=True`` yields a header row first (column labels); when
      the DataFrame has a MultiIndex on columns, each level is yielded
      as a separate header row.
    - ``index=True`` prepends the row index to each data row; with a
      MultiIndex index, each index level becomes its own leading column.
    - An empty separator row is yielded between header and data when
      ``index=True`` to match openpyxl's layout convention.
    """
    import pandas as pd

    if header:
        if isinstance(df.columns, pd.MultiIndex):
            for level in range(df.columns.nlevels):
                row: list[Any] = []
                if index:
                    row.extend([None] * df.index.nlevels)
                row.extend(df.columns.get_level_values(level).tolist())
                yield row
        else:
            row = []
            if index:
                row.extend([None] * df.index.nlevels)
            row.extend(df.columns.tolist())
            yield row

    if index and header:
        # Blank separator row between header and data (openpyxl convention).
        yield []

    for idx, values in zip(df.index, df.itertuples(index=False, name=None)):
        if index:
            if isinstance(df.index, pd.MultiIndex):
                yield list(idx) + list(values)
            else:
                yield [idx] + list(values)
        else:
            yield list(values)


__all__ = ["dataframe_to_rows"]
