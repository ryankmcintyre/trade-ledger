from collections import Counter
from datetime import date
from pathlib import Path

import pytest

from trade_ingestion.adapters.fidelity import parse_fidelity_csv
from trade_ingestion.matcher import match_trades_with_summary
from trade_ingestion.models import CanonicalTrade

FIXTURE_DIR = Path(__file__).parent / "fixtures"
FIDELITY_HISTORY_FIXTURE = FIXTURE_DIR / "History_for_Account_Zxxxxxxx.csv"


def _row(trade: CanonicalTrade) -> tuple[object, ...]:
    return (
        trade.stock,
        trade.open_date,
        trade.exp_date,
        trade.call_or_put,
        trade.side,
        trade.strike,
        trade.premium,
        trade.quantity,
        trade.exit_price,
        trade.close_date,
        trade.account,
    )


@pytest.mark.integration
def test_fidelity_fixture_matches_expected_canonical_trades() -> None:
    content = FIDELITY_HISTORY_FIXTURE.read_text(encoding="utf-8-sig")

    events = parse_fidelity_csv(content)
    result = match_trades_with_summary(events)

    expected = [
        (
            "P",
            date(2026, 5, 27),
            None,
            None,
            "C",
            None,
            76.0,
            0.5,
            79.322,
            date(2026, 5, 27),
            "Fidelity",
        ),
        ("P", date(2026, 5, 27), None, None, "C", None, 76.0, 0.5, None, None, "Fidelity"),
        ("SOLS", None, None, None, "C", None, None, 1.0, 86.3, date(2026, 5, 27), "Fidelity"),
        (
            "NOK",
            date(2026, 5, 27),
            date(2026, 10, 16),
            "C",
            "B",
            16.0,
            3.29,
            6.0,
            None,
            None,
            "Fidelity",
        ),
        (
            "S&P 500 INDEX",
            date(2026, 5, 27),
            date(2026, 6, 18),
            "P",
            "S",
            7410.0,
            56.47,
            2.0,
            None,
            None,
            "Fidelity",
        ),
        (
            "S&P 500 INDEX",
            date(2026, 5, 27),
            date(2026, 6, 18),
            "P",
            "B",
            7400.0,
            54.07,
            2.0,
            None,
            None,
            "Fidelity",
        ),
    ]

    assert len(result.trades) == 6
    assert result.skipped_duplicates == 0
    assert result.open_positions == 4
    assert Counter(_row(trade) for trade in result.trades) == Counter(expected)
