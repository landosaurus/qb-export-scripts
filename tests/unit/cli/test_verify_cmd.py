from __future__ import annotations

from unittest.mock import MagicMock

from click.testing import CliRunner
from pytest_mock import MockerFixture

from qb_cli.ops.verify_op import VerifyResult


def _fake_ctx() -> MagicMock:
    return MagicMock(name="FakeContext")


def test_verify_all_found_exits_zero(mocker: MockerFixture) -> None:
    from qb_cli.cli.root import qb

    mocker.patch(
        "qb_cli.cli.verify_cmd._verify",
        return_value=VerifyResult(
            found={"customer": {"Acme"}, "vendor": set(), "item": set(), "terms": set()},
            missing={"customer": set(), "vendor": set(), "item": set(), "terms": set()},
        ),
    )

    runner = CliRunner()
    result = runner.invoke(
        qb, ["verify", "--customer", "Acme"], obj=_fake_ctx()
    )

    assert result.exit_code == 0, result.output
    assert "found customer: Acme" in result.output


def test_verify_missing_exits_one_with_stderr(mocker: MockerFixture) -> None:
    from qb_cli.cli.root import qb

    mocker.patch(
        "qb_cli.cli.verify_cmd._verify",
        return_value=VerifyResult(
            found={"customer": {"A"}},
            missing={"customer": {"B"}},
        ),
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        ["verify", "--customer", "A", "--customer", "B"],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 1
    assert "MISSING customer: B" in result.stderr
    assert "found customer: A" in result.output


def test_verify_forwards_all_buckets(mocker: MockerFixture) -> None:
    from qb_cli.cli.root import qb

    mock_verify = mocker.patch(
        "qb_cli.cli.verify_cmd._verify",
        return_value=VerifyResult(found={}, missing={}),
    )

    runner = CliRunner()
    result = runner.invoke(
        qb,
        [
            "verify",
            "--customer", "C1",
            "--vendor", "V1",
            "--item", "I1",
            "--terms", "T1",
        ],
        obj=_fake_ctx(),
    )

    assert result.exit_code == 0, result.output
    kwargs = mock_verify.call_args.kwargs
    assert kwargs["customers"] == ["C1"]
    assert kwargs["vendors"] == ["V1"]
    assert kwargs["items"] == ["I1"]
    assert kwargs["terms"] == ["T1"]
