"""Deliver a generated report to its recipient, by either or both channels.

  Email    -- through the Outlook desktop app we are already signed in to
              (Outlook.Application COM), so no password and no SMTP setup.
              Requires CLASSIC Outlook -- the "new Outlook" has no COM automation.

  OneDrive -- a plain file copy into the recipient's folder (which OneDrive or
              SharePoint sync then picks up). No Outlook involved, so this
              channel works even on machines that only have the new Outlook.

A delivery failure is reported (it means the report did not reach someone, which
we must know) but it is raised as a per-recipient delivery failure, not a crash:
one bad address or missing folder never stops the rest of the batch.

run_date is passed in by the caller (computed once per run in distribute_reports.py)
rather than read from the clock here,
so the {date} in an email and the date in the generated file name always agree.
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path


@dataclass
class DeliveryOutcome:
    channel: str            # "email" | "onedrive"
    status: str             # SUCCESS | FAILED | SKIPPED
    detail: str = ""
    error_type: str = ""
    error_message: str = ""


def _format(template: str, recipient_name: str, run_date: str) -> str:
    return template.replace("{name}", recipient_name).replace("{date}", run_date)


def deliver_email(recipient, report_path: Path, email_cfg, run_date: str) -> DeliveryOutcome:
    if not recipient.email:
        return DeliveryOutcome("email", "SKIPPED", detail="no email address on this recipient")
    try:
        import win32com.client            # provided by pywin32 (installed with xlwings)

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)      # 0 = olMailItem
        mail.To = recipient.email
        mail.Subject = _format(email_cfg.subject, recipient.name, run_date)
        mail.Body = _format(email_cfg.body, recipient.name, run_date)
        mail.Attachments.Add(str(report_path.resolve()))
        if email_cfg.mode == "send":
            mail.Send()
            return DeliveryOutcome("email", "SUCCESS", detail=f"sent to {recipient.email}")
        mail.Save()                       # draft -> Outlook Drafts folder
        return DeliveryOutcome("email", "SUCCESS", detail=f"draft saved for {recipient.email}")
    except Exception as exc:
        return DeliveryOutcome("email", "FAILED", error_type=type(exc).__name__, error_message=str(exc)[:500])


def deliver_onedrive(recipient, report_path: Path) -> DeliveryOutcome:
    folder_str = recipient.onedrive_folder
    if not folder_str:
        return DeliveryOutcome("onedrive", "SKIPPED", detail="no OneDrive folder on this recipient")
    try:
        folder = Path(folder_str)
        folder.mkdir(parents=True, exist_ok=True)
        dest = folder / report_path.name
        shutil.copy2(report_path, dest)
        return DeliveryOutcome("onedrive", "SUCCESS", detail=f"copied to {dest}")
    except Exception as exc:
        return DeliveryOutcome("onedrive", "FAILED", error_type=type(exc).__name__, error_message=str(exc)[:500])


def deliver(recipient, report_path: Path, cfg, run_date: str) -> list[DeliveryOutcome]:
    """Run every enabled channel for one recipient and return their outcomes."""
    outcomes: list[DeliveryOutcome] = []
    if cfg.email.enabled:
        outcomes.append(deliver_email(recipient, report_path, cfg.email, run_date))
    if cfg.onedrive.enabled:
        outcomes.append(deliver_onedrive(recipient, report_path))
    if not outcomes:
        outcomes.append(DeliveryOutcome("none", "SKIPPED", detail="no delivery channel enabled"))
    return outcomes
