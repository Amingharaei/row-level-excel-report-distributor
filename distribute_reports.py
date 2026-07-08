"""Row-Level Report Distributor

It reads config.toml and the recipients list, prints the resolved scope for every
recipient (so we can catch a wrong or missing access key BEFORE anything is
generated), then for each enabled recipient it generates a correctly-scoped
workbook and -- if delivery is turned on -- emails and/or copies it. 
It writes the event log and exits 0 if every recipient succeeded or 1 if any failed.
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import replace
from datetime import date
from pathlib import Path

from config import Config, ConfigError, load_config
from delivery import deliver
from eventlog import EventLog, setup_text_logger
from excel_engine import RecipientResult, generate_for_recipient
from recipients import Recipient, RecipientsError, load_recipients


def _console(message: str = "") -> None:
    """Best-effort console feedback; safe even if there is no console (pythonw)."""
    try:
        print(message, flush=True)
    except Exception:
        pass


def _print_preview(recipients: list[Recipient], cfg: Config, deliver_on: bool) -> None:
    _console("Recipients and their resolved scope:")
    _console(f"  {'Recipient':<24}{'Sees IDs':<22}{'Delivery'}")
    _console("  " + "-" * 66)
    for r in recipients:
        if not r.enabled:
            _console(f"  {r.name:<24}{'(disabled)':<22}-")
            continue
        channels = []
        if deliver_on and cfg.email.enabled and r.email:
            channels.append(f"email:{cfg.email.mode}")
        if deliver_on and cfg.onedrive.enabled and r.onedrive_folder:
            channels.append("onedrive")
        plan = ", ".join(channels) if channels else ("generate only" if not deliver_on else "no channel")
        _console(f"  {r.name:<24}{r.scope_label():<22}{plan}")
    _console()
    if not deliver_on:
        _console("Delivery is OFF: copies will be generated but nothing will be sent.")
    _console()


def _deliver_and_log(recipient: Recipient, result: RecipientResult, cfg: Config,
                     events: EventLog, logger, run_date: str) -> bool:
    """Deliver one generated report; return True if every enabled channel succeeded."""
    all_ok = True
    for outcome in deliver(recipient, Path(result.output_path), cfg, run_date):
        events.emit(scope="delivery", target=f"{recipient.name}/{outcome.channel}",
                    event="DELIVERY", status=outcome.status, detail=outcome.detail,
                    error_type=outcome.error_type, error_message=outcome.error_message)
        if outcome.status == "FAILED":
            all_ok = False
            logger.error("[%s] %s delivery failed: %s", recipient.name, outcome.channel, outcome.error_message)
        else:
            logger.info("[%s] %s: %s", recipient.name, outcome.channel, outcome.detail)
        _console(f"        {outcome.channel:<9} {outcome.status:<8} {outcome.detail}")
    return all_ok


def run(config: Config, deliver_on: bool, recipients: list[Recipient] | None = None) -> int:
    _console("Row-Level Report Distributor")
    _console("Generating one scoped report per recipient. This can take a while")
    _console("depending on how large the report is and how many recipients there are.")
    _console()

    logger = setup_text_logger(config.log_dir)
    events = EventLog(config.log_dir)
    run_date = date.today().isoformat()      # shared by every recipient's file name and email

    if recipients is None:
        try:
            recipients = load_recipients(config.recipients_path, config.recipients_sheet)
        except RecipientsError as exc:
            _console(f"Recipients error: {exc}")
            logger.error("Recipients error: %s", exc)
            events.close()
            return 2

    logger.info("Run %s start. Master: %s  Recipients: %d",
                events.run_id, config.master_workbook, len(recipients))
    events.emit(scope="run", target=str(config.master_workbook), event="RUN", status="STARTED",
                detail=f"{len(recipients)} recipient(s); delivery={'on' if deliver_on else 'off'}; date={run_date}")

    _print_preview(recipients, config, deliver_on)

    active = [r for r in recipients if r.enabled]
    results: list[tuple[RecipientResult, bool]] = []      # (generation result, delivery_ok)
    try:
        for i, recipient in enumerate(active, start=1):
            _console(f"[{i}/{len(active)}] {recipient.name}  (sees {recipient.scope_label()}) ...")
            events.emit(scope="recipient", target=recipient.name, event="RECIPIENT", status="STARTED",
                        detail=f"scope={recipient.scope_label()}")

            result = generate_for_recipient(
                config.master_workbook, config.output_dir, recipient, config, events, logger,
                run_date=run_date)

            delivery_ok = True
            if result.status == "SUCCESS":
                _console(f"    generated {result.duration:.1f}s -> {Path(result.output_path).name}")
                if deliver_on:
                    delivery_ok = _deliver_and_log(recipient, result, config, events, logger, run_date)
            else:
                _console(f"    FAILED ({result.duration:.1f}s): {result.error_message or 'query failure(s)'}")

            overall = "SUCCESS" if (result.status == "SUCCESS" and delivery_ok) else "FAILED"
            detail = "" if delivery_ok else "delivery failure"
            events.emit(scope="recipient", target=recipient.name, event="RECIPIENT", status=overall,
                        duration_seconds=result.duration, error_type=result.error_type,
                        error_message=result.error_message, detail=detail)
            results.append((result, delivery_ok))
    finally:
        failures = sum(1 for r, d in results if r.status != "SUCCESS" or not d)
        events.emit(scope="run", target=str(config.master_workbook), event="RUN",
                    status="FAILED" if failures else "SUCCESS",
                    detail=f"{len(results)} recipient(s), {failures} failure(s)")
        events.close()

    ok = len(results) - failures
    _console()
    _console(f"Done. {ok}/{len(results)} recipient(s) completed successfully.")
    if failures:
        _console("Check distributor.log and distribution-events.csv in the log folder for details.")
    _console(f"Reports are in: {config.output_dir}")
    logger.info("Run done. %d recipient(s), failures=%d", len(results), failures)
    return 1 if failures else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate and distribute row-level-scoped Excel reports.")
    parser.add_argument("--config", type=Path, default=Path("config.toml"),
                        help="Path to the config file (default: ./config.toml).")
    parser.add_argument("--dry-run", action="store_true",
                        help="Generate copies but deliver nothing, regardless of config.")
    parser.add_argument("--only", type=str, default=None,
                        help="Generate for a single recipient by exact name (for testing one report).")
    args = parser.parse_args(argv)

    if os.name != "nt":
        _console("This tool drives the Excel desktop app (and its Power Pivot data model) over COM, "
                 "which is Windows only.")
        return 3

    try:
        config = load_config(args.config)
    except ConfigError as exc:
        _console(f"Configuration error: {exc}")
        return 2

    deliver_on = config.delivery_enabled and not args.dry_run

    subset: list[Recipient] | None = None
    if args.only is not None:
        try:
            everyone = load_recipients(config.recipients_path, config.recipients_sheet)
        except RecipientsError as exc:
            _console(f"Recipients error: {exc}")
            return 2
        subset = [r for r in everyone if r.name.lower() == args.only.lower()]
        if not subset:
            _console(f"No recipient named '{args.only}' in {config.recipients_path}.")
            return 2

    return run(config, deliver_on, recipients=subset)


if __name__ == "__main__":
    sys.exit(main())