"""Load and validate config.toml.

There are no secrets in the config: 
The Outlook delivery uses the signed-in desktop Outlook, 
and OneDrive delivery is a plain file copy to a folder we already sync.
"""

from __future__ import annotations

import tomllib
from dataclasses import dataclass
from pathlib import Path

EMAIL_MODES = ("draft", "send")


class ConfigError(Exception):
    """Raised for a missing or invalid setting."""


@dataclass(frozen=True)
class EmailConfig:
    enabled: bool
    mode: str                 # "draft" | "send"
    subject: str
    body: str


@dataclass(frozen=True)
class OneDriveConfig:
    enabled: bool


@dataclass(frozen=True)
class Config:
    master_workbook: Path
    output_dir: Path
    log_dir: Path
    recipients_path: Path
    recipients_sheet: str
    access_key_cell: str
    timeout_seconds: int
    delivery_enabled: bool
    email: EmailConfig
    onedrive: OneDriveConfig


def _req(table: dict, key: str, where: str) -> object:
    if key not in table:
        raise ConfigError(f"Missing '{key}' under [{where}] in config.toml.")
    return table[key]


def load_config(path: Path) -> Config:
    if not path.is_file():
        raise ConfigError(f"Config file not found: {path}")
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{path} is not valid TOML: {exc}") from exc

    paths = data.get("paths", {})
    master = Path(str(_req(paths, "master_workbook", "paths")))
    output_dir = Path(str(_req(paths, "output_dir", "paths")))
    log_dir = Path(str(paths.get("log_dir", output_dir.parent / "rls-distributor-logs")))

    recips = data.get("recipients", {})
    recipients_path = Path(str(_req(recips, "path", "recipients")))
    recipients_sheet = str(recips.get("sheet", "Recipients"))

    model = data.get("model", {})
    access_key_cell = str(model.get("access_key_cell", "pAccessKey"))
    timeout = int(model.get("timeout_seconds", 600))
    if timeout <= 0:
        raise ConfigError("[model].timeout_seconds must be a positive number of seconds.")

    delivery = data.get("delivery", {})
    delivery_enabled = bool(delivery.get("enabled", False))

    email_table = delivery.get("email", {})
    email_mode = str(email_table.get("mode", "send")).lower()
    if email_mode not in EMAIL_MODES:
        raise ConfigError(f"[delivery.email].mode must be 'draft' or 'send' (got '{email_mode}').")
    email = EmailConfig(
        enabled=bool(email_table.get("enabled", False)),
        mode=email_mode,
        subject=str(email_table.get("subject", "Your report - {date}")),
        body=str(email_table.get("body", "Hi {name},\n\nAttached is your report as of {date}.")),
    )

    od_table = delivery.get("onedrive", {})
    onedrive = OneDriveConfig(enabled=bool(od_table.get("enabled", False)))

    return Config(
        master_workbook=master,
        output_dir=output_dir,
        log_dir=log_dir,
        recipients_path=recipients_path,
        recipients_sheet=recipients_sheet,
        access_key_cell=access_key_cell,
        timeout_seconds=timeout,
        delivery_enabled=delivery_enabled,
        email=email,
        onedrive=onedrive,
    )
