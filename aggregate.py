#!/usr/bin/env python3
"""Aggregate parking durations from licence plate Excel exports."""
from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
from openpyxl.utils import get_column_letter


@dataclass
class ColumnMapping:
    plate: str
    event: str
    timestamp: str


@dataclass
class AppConfig:
    source_folder: Path
    output_file: Path
    columns: ColumnMapping
    entry_marker: str
    exit_marker: str
    timestamp_format: Optional[str]
    recursive: bool


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Aggregate parking durations by licence plate based on ENTRY/EXIT "
            "events stored in Excel workbooks."
        )
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Path to the JSON configuration file (default: config.json)",
    )
    parser.add_argument(
        "--source-folder",
        help="Override source folder that contains XLS/XLSX files",
    )
    parser.add_argument(
        "--output-file",
        help="Override destination Excel file for aggregated results",
    )
    parser.add_argument(
        "--timestamp-format",
        help="Override timestamp format understood by pandas.to_datetime",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Recursively search for Excel files inside the source folder",
    )
    return parser.parse_args()


def load_config(config_path: Path) -> Dict:
    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file {config_path} was not found. "
            "Create it based on config.example.json."
        )
    with config_path.open("r", encoding="utf-8") as cfg_file:
        return json.load(cfg_file)


def normalize_config(raw_cfg: Dict, overrides: argparse.Namespace, config_base: Path) -> AppConfig:
    def _resolve_path(value: Optional[str], default: Optional[Path] = None) -> Path:
        if value:
            candidate = Path(value)
            if not candidate.is_absolute():
                candidate = (config_base / candidate).resolve()
            return candidate
        if default is None:
            raise ValueError("Path configuration missing and no default provided")
        return default

    source_folder = _resolve_path(
        overrides.source_folder or raw_cfg.get("source_folder")
    )
    output_file = _resolve_path(
        overrides.output_file or raw_cfg.get("output_file", "output/parking_durations.xlsx")
    )
    columns_cfg = raw_cfg.get("columns", {})
    column_map = ColumnMapping(
        plate=columns_cfg.get("plate", "Plate"),
        event=columns_cfg.get("event", "Event"),
        timestamp=columns_cfg.get("timestamp", "Timestamp"),
    )

    entry_marker = (raw_cfg.get("entry_marker") or "01 ENTRY").upper()
    exit_marker = (raw_cfg.get("exit_marker") or "02 EXIT").upper()
    timestamp_format = overrides.timestamp_format or raw_cfg.get("timestamp_format")
    recursive = overrides.recursive or raw_cfg.get("recursive", False)

    return AppConfig(
        source_folder=source_folder,
        output_file=output_file,
        columns=column_map,
        entry_marker=entry_marker,
        exit_marker=exit_marker,
        timestamp_format=timestamp_format,
        recursive=recursive,
    )


def discover_excel_files(folder: Path, recursive: bool) -> List[Path]:
    pattern = "**/*.xls*" if recursive else "*.xls*"
    files = [
        file_path
        for file_path in folder.glob(pattern)
        if file_path.is_file() and not file_path.name.startswith("~$")
    ]
    files.sort()
    return files


def load_events(
    files: Iterable[Path],
    columns: ColumnMapping,
    timestamp_format: Optional[str],
) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for file_path in files:
        try:
            frame = pd.read_excel(file_path)
        except Exception as exc:  # pragma: no cover - user feedback
            print(f"[SKIP] Failed to read {file_path}: {exc}", file=sys.stderr)
            continue

        missing = {
            columns.plate,
            columns.event,
            columns.timestamp,
        } - set(frame.columns)
        if missing:
            print(
                f"[SKIP] {file_path} missing required columns: {sorted(missing)}",
                file=sys.stderr,
            )
            continue

        trimmed = frame[[columns.plate, columns.event, columns.timestamp]].copy()
        trimmed = trimmed.dropna(subset=[columns.plate])
        trimmed.columns = ["plate", "event", "timestamp"]
        frames.append(trimmed)
        print(f"[LOAD] {file_path}: {len(trimmed)} rows")

    if not frames:
        return pd.DataFrame(columns=["plate", "event", "timestamp"])

    events = pd.concat(frames, ignore_index=True)
    events["timestamp"] = pd.to_datetime(
        events["timestamp"],
        format=timestamp_format,
        errors="coerce",
    )
    events = events.dropna(subset=["timestamp"]).copy()
    events["plate"] = events["plate"].astype(str).str.strip().str.upper()
    events["event"] = events["event"].astype(str).str.strip().str.upper()
    events = events[(events["plate"] != "") & (events["plate"] != "NAN")]
    return events


def build_intervals(
    events: pd.DataFrame,
    entry_marker: str,
    exit_marker: str,
) -> Tuple[pd.DataFrame, List[Dict[str, str]]]:
    intervals: List[Dict[str, object]] = []
    issues: List[Dict[str, str]] = []

    if events.empty:
        return pd.DataFrame(columns=["plate", "entry_time", "exit_time", "duration_minutes"]), issues

    for plate, group in events.groupby("plate"):
        group = group.sort_values("timestamp")
        if plate.isdigit():
            for _, hazard_row in group.iterrows():
                issues.append(
                    {
                        "plate": plate,
                        "issue": "Hazard plate number",
                        "timestamp": hazard_row["timestamp"].isoformat(),
                    }
                )
            continue
        open_entry: Optional[pd.Timestamp] = None
        for _, row in group.iterrows():
            event = row["event"]
            timestamp = row["timestamp"]
            if event == entry_marker:
                if open_entry is not None:
                    issues.append(
                        {
                            "plate": plate,
                            "issue": "Consecutive ENTRY without EXIT",
                            "timestamp": open_entry.isoformat(),
                        }
                    )
                open_entry = timestamp
            elif event == exit_marker:
                if open_entry is None:
                    issues.append(
                        {
                            "plate": plate,
                            "issue": "EXIT without matching ENTRY",
                            "timestamp": timestamp.isoformat(),
                        }
                    )
                    continue
                duration = (timestamp - open_entry).total_seconds() / 60.0
                if duration < 0:
                    issues.append(
                        {
                            "plate": plate,
                            "issue": "EXIT earlier than ENTRY",
                            "timestamp": timestamp.isoformat(),
                        }
                    )
                    open_entry = None
                    continue
                intervals.append(
                    {
                        "plate": plate,
                        "entry_time": open_entry,
                        "exit_time": timestamp,
                        "duration_minutes": round(duration, 2),
                    }
                )
                open_entry = None
        if open_entry is not None:
            issues.append(
                {
                    "plate": plate,
                    "issue": "ENTRY without matching EXIT",
                    "timestamp": open_entry.isoformat(),
                }
            )

    intervals_df = pd.DataFrame(intervals)
    return intervals_df, issues


def summarize_monthly(intervals: pd.DataFrame) -> pd.DataFrame:
    if intervals.empty:
        return pd.DataFrame(
            columns=["plate", "month", "visits", "total_minutes", "total_hours"]
        )
    monthly = intervals.copy()
    monthly["month"] = monthly["entry_time"].dt.to_period("M").astype(str)
    agg = (
        monthly.groupby(["plate", "month"])["duration_minutes"]
        .agg(["count", "sum"])
        .rename(columns={"count": "visits", "sum": "total_minutes"})
        .reset_index()
    )
    agg["total_hours"] = (agg["total_minutes"] / 60.0).round(2)
    agg["total_minutes"] = agg["total_minutes"].round(2)
    return agg


def write_output(
    output_file: Path,
    intervals: pd.DataFrame,
    monthly: pd.DataFrame,
    issues: List[Dict[str, str]],
) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file) as writer:
        monthly.to_excel(writer, sheet_name="monthly_totals", index=False)
        _autosize_sheet(writer, "monthly_totals", monthly)

        intervals.to_excel(writer, sheet_name="intervals", index=False)
        _autosize_sheet(writer, "intervals", intervals)

        if issues:
            issues_df = pd.DataFrame(issues)
            issues_df.to_excel(writer, sheet_name="issues", index=False)
            _autosize_sheet(writer, "issues", issues_df)
    print(f"Saved aggregated results to {output_file}")


def _autosize_sheet(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    if dataframe.empty and not list(dataframe.columns):
        return
    for idx, column in enumerate(dataframe.columns, start=1):
        header_length = len(str(column))
        if dataframe.empty:
            max_length = header_length
        else:
            series = dataframe[column].astype(str)
            cell_lengths = series.map(len).tolist()
            max_length = max([header_length] + cell_lengths)
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 2, 60)


def main() -> None:
    args = parse_args()
    config_path = Path(args.config)
    raw_cfg = load_config(config_path)
    app_config = normalize_config(raw_cfg, args, config_path.parent)

    excel_files = discover_excel_files(app_config.source_folder, app_config.recursive)
    if not excel_files:
        print(
            f"No Excel files were found in {app_config.source_folder}. "
            "Ensure the folder path in the configuration is correct.",
            file=sys.stderr,
        )
        sys.exit(1)

    events = load_events(excel_files, app_config.columns, app_config.timestamp_format)
    if events.empty:
        print("No valid events could be read from the Excel files.", file=sys.stderr)
        sys.exit(1)

    intervals, issues = build_intervals(events, app_config.entry_marker, app_config.exit_marker)
    monthly = summarize_monthly(intervals)
    write_output(app_config.output_file, intervals, monthly, issues)


if __name__ == "__main__":
    main()
