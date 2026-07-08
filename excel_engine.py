"""The Excel side of the job: 
generate one correctly-scoped workbook per recipient, in that recipient's own isolated Excel instance, guarded by a watchdog.

Per recipient, the sequence is:

  1. Open the master workbook (never saved -- we only ever write COPIES).
  2. Write the recipient's access key into the driver cell (a named range).
  3. Refresh each Power Query (OLEDB/ODBC) connection individually, with
     BackgroundQuery = False so the call blocks until done.
  4. Refresh the PivotTables so they reflect the reloaded (filtered) model.
  5. Recalculate, then SaveCopyAs a per-recipient, date-stamped file. SaveCopyAs
     writes a copy without touching the in-memory master, which is what lets one
     open workbook produce many scoped copies.

If any query fails, the copy is NOT written.
"""

from __future__ import annotations

import os
import re
import signal
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from time import perf_counter

import xlwings as xw

_XL_OLEDB = 1        # Power Query connections are OLEDB.
_XL_ODBC = 2


@dataclass
class ConnectionResult:
    name: str
    status: str                     # SUCCESS | FAILED | SKIPPED
    duration: float = 0.0
    error_type: str = ""
    error_message: str = ""


@dataclass
class RecipientResult:
    name: str
    status: str = "SUCCESS"         # SUCCESS | FAILED
    duration: float = 0.0
    output_path: str = ""
    connections: list[ConnectionResult] = field(default_factory=list)
    error_type: str = ""
    error_message: str = ""

    @property
    def failed_connections(self) -> list[ConnectionResult]:
        return [c for c in self.connections if c.status == "FAILED"]


def safe_filename(name: str) -> str:
    """A file-system-safe stem from a recipient name (keeps it recognisable)."""
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", name).strip().rstrip(".")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "recipient"


@contextmanager
def _excel_app(timeout_seconds: int, logger):
    """A fresh, invisible Excel instance with a watchdog that kills it on timeout.

    One instance per recipient, so a hang or crash while generating one report
    cannot corrupt the others. (Identical approach to the refresh orchestrator.)
    """
    app = xw.App(visible=False, add_book=False)
    pid = app.pid
    done = threading.Event()

    def watchdog() -> None:
        if not done.wait(timeout_seconds):          # False => timed out
            logger.error("Timeout after %ss; killing Excel PID %s.", timeout_seconds, pid)
            try:
                os.kill(pid, signal.SIGTERM)        # TerminateProcess on Windows
            except OSError as exc:
                logger.error("Could not kill Excel PID %s: %s", pid, exc)

    watcher = threading.Thread(target=watchdog, daemon=True)
    watcher.start()
    try:
        app.display_alerts = False                  # also suppresses the "save changes?" prompt
        app.screen_updating = False
        try:
            app.api.AskToUpdateLinks = False
        except Exception:
            pass
        yield app
    finally:
        done.set()                                  # stop the watchdog
        try:
            app.quit()
        except Exception:                           # may already be killed
            pass
        watcher.join(timeout=5)


def _set_access_key(wb, cell_name: str, value: str) -> None:
    """Write the access key into the driver named range before refreshing.

    Written as text; the AccessKey query coerces with Text.From either way, so a
    single numeric id and a comma list are handled the same.
    """
    try:
        rng = wb.names[cell_name].refers_to_range
    except Exception as exc:
        raise RuntimeError(
            f"Named cell '{cell_name}' was not found in the master workbook. "
            f"Create it on the Control sheet (see MASTER-WORKBOOK-SETUP.md)."
        ) from exc
    rng.value = value if value != "" else None      # blank cell -> "" in Power Query -> unrestricted


def _background_off(conn) -> bool:
    """Turn off background refresh so the refresh call blocks until done.
    Returns True if this is a refreshable data connection (OLEDB/ODBC)."""
    ctype = conn.Type
    if ctype == _XL_OLEDB:
        conn.OLEDBConnection.BackgroundQuery = False
        return True
    if ctype == _XL_ODBC:
        conn.ODBCConnection.BackgroundQuery = False
        return True
    return False


def _refresh_with_retry(conn, logger) -> None:
    """Refresh one connection, retrying once (the first code-driven refresh after
    opening a file sometimes throws a spurious init error; a real failure repeats)."""
    try:
        conn.Refresh()
    except Exception as first:
        logger.info("Retrying connection '%s' after: %s", conn.Name, first)
        conn.Refresh()


def _refresh_connections(wb, events, logger, recipient_name: str) -> list[ConnectionResult]:
    results: list[ConnectionResult] = []
    for conn in wb.api.Connections:
        name = conn.Name
        try:
            refreshable = _background_off(conn)
        except Exception:
            refreshable = False
        if not refreshable:
            results.append(ConnectionResult(name, "SKIPPED"))
            events.emit(scope="connection", target=name, event="REFRESH",
                        status="SKIPPED", detail="not an OLEDB/ODBC connection")
            continue

        c_start = perf_counter()
        try:
            _refresh_with_retry(conn, logger)
        except Exception as exc:
            dur = perf_counter() - c_start
            cr = ConnectionResult(name, "FAILED", dur, type(exc).__name__, str(exc)[:500])
            results.append(cr)
            events.emit(scope="connection", target=name, event="REFRESH", status="FAILED",
                        duration_seconds=dur, error_type=cr.error_type, error_message=cr.error_message)
            logger.error("[%s] query '%s' failed: %s", recipient_name, name, exc)
        else:
            dur = perf_counter() - c_start
            results.append(ConnectionResult(name, "SUCCESS", dur))
            events.emit(scope="connection", target=name, event="REFRESH",
                        status="SUCCESS", duration_seconds=dur)
    return results


def _refresh_pivots(wb, logger, recipient_name: str) -> None:
    """Refresh every PivotTable to make sure it reflects the freshly reloaded model.
    """
    for sheet in wb.sheets:
        try:
            pivots = sheet.api.PivotTables()
            count = pivots.Count
        except Exception:
            continue
        for i in range(1, count + 1):
            try:
                pivots.Item(i).RefreshTable()
            except Exception as exc:
                logger.info("[%s] pivot refresh on sheet '%s' #%d: %s",
                            recipient_name, sheet.name, i, exc)


def generate_for_recipient(
    master_path: Path, output_dir: Path, recipient, cfg, events, logger, *, run_date: str
) -> RecipientResult:
    """Produce one scoped workbook for a recipient, named with run_date so a
    later run never overwrites an earlier day's file.
    """
    result = RecipientResult(name=recipient.name)
    start = perf_counter()
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"{safe_filename(recipient.name)} - {run_date}.xlsx"

    try:
        with _excel_app(cfg.timeout_seconds, logger) as app:
            wb = app.books.open(str(master_path), update_links=False)
            try:
                _set_access_key(wb, cfg.access_key_cell, recipient.access_key)
                result.connections = _refresh_connections(wb, events, logger, recipient.name)

                if result.failed_connections:
                    # Do NOT write a copy on a failed refresh: it could be wrongly scoped.
                    result.status = "FAILED"
                    logger.error("[%s] %d query failure(s); no copy written.",
                                 recipient.name, len(result.failed_connections))
                    return result

                _refresh_pivots(wb, logger, recipient.name)
                app.calculate()
                wb.api.SaveCopyAs(str(out_path))
                result.output_path = str(out_path)
            finally:
                try:
                    wb.close()                       # never save the master
                except Exception:
                    pass
    except Exception as exc:
        result.status = "FAILED"
        result.error_type = type(exc).__name__
        result.error_message = str(exc)[:500]
        logger.error("[%s] generation failed: %s", recipient.name, exc)

    result.duration = perf_counter() - start
    return result
