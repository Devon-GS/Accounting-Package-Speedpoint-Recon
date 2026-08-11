from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from uuid import uuid4
import sys

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def get_app_root() -> Path:
	if getattr(sys, "frozen", False):
		return Path(sys.executable).resolve().parent
	return Path(__file__).resolve().parent


BASE_DIR = get_app_root()
SUPPORTED_EXTENSIONS = {".xlsx"}
SPEEDPOINT_HIGHLIGHT = PatternFill(fill_type="solid", fgColor="FF86B9")
BANK_HIGHLIGHT = PatternFill(fill_type="solid", fgColor="FFFF00")


def normalize_text(value: object) -> str:
	text = str(value or "").strip().upper()
	return re.sub(r"[^A-Z0-9]+", " ", text).strip()


def normalize_tokens(value: object) -> list[str]:
	text = normalize_text(value)
	if not text:
		return []
	return text.split()


def compact_text(value: object) -> str:
	return re.sub(r"\s+", " ", str(value or "").strip())


def parse_money(value: object) -> float | None:
	if value in (None, ""):
		return None
	if isinstance(value, str):
		cleaned = value.replace(" ", "").replace(",", "")
		if not cleaned:
			return None
		try:
			number = float(cleaned)
		except ValueError:
			return None
	else:
		try:
			number = float(value)
		except (TypeError, ValueError):
			return None
	if abs(number) < 1e-9:
		return None
	return round(number, 2)


def parse_amount(value: object) -> float | None:
	number = parse_money(value)
	if number is None:
		return None
	return round(number, 2)


def format_download_name(original_name: str, strip_leading_number: bool = False) -> str:
	path = Path(original_name)
	stem = path.stem.replace("_", " ")
	if strip_leading_number:
		stem = re.sub(r"^\s*\d+[\s._-]*", "", stem)
	stem = re.sub(r"\s+", " ", stem).strip()
	if not stem:
		stem = "download"
	suffix = path.suffix or ".xlsx"
	return f"{stem}{suffix}"


def sanitize_filename(filename: str, fallback: str = "upload.xlsx") -> str:
	name = Path(filename).name.strip()
	if not name:
		return fallback
	name = re.sub(r'[<>:"/\\|?*]+', "_", name)
	name = re.sub(r"\s+", " ", name).strip()
	return name or fallback


def sha256_file(path: Path) -> str:
	digest = hashlib.sha256()
	with path.open("rb") as handle:
		for chunk in iter(lambda: handle.read(8192), b""):
			digest.update(chunk)
	return digest.hexdigest()


def pick_bank_statement_sheet(workbook):
	for worksheet in workbook.worksheets:
		for row_number in range(1, min(worksheet.max_row, 5) + 1):
			values = [normalize_text(worksheet.cell(row_number, column).value) for column in (1, 2, 4)]
			if values == ["DATE", "TRANSACTIONS", "CREDIT"]:
				return worksheet
	return workbook.worksheets[0]


def pick_speedpoint_sheet(workbook):
	for worksheet in workbook.worksheets:
		if normalize_text(worksheet.title) in {"SPEED POINTS", "SPEEDPOINT", "SPEEDPOINTS"}:
			return worksheet
	return workbook.worksheets[0]


def parse_speedpoint_bank_transaction(transaction: object) -> dict[str, object] | None:
	text = compact_text(transaction).upper()
	if not text:
		return None

	speedpoint_match = re.match(r"^SPEEDPOINT\s*(?P<terminal>\d+)\s+(?P<batch>\d+)$", text, flags=re.IGNORECASE)
	if speedpoint_match:
		terminal = int(speedpoint_match.group("terminal"))
		terminal_map = {
			946389: (2, 3),
			946390: (4, 5),
		}
		speedpoint_column, amount_column = terminal_map.get(terminal, (2, 3))
		return {
			"transaction_type": "SPEEDPOINT",
			"bank_terminal": speedpoint_match.group("terminal"),
			"batch_number": int(speedpoint_match.group("batch")),
			"speedpoint_terminal": terminal,
			"speedpoint_column": speedpoint_column,
			"amount_column": amount_column,
		}

	nedlnk_match = re.search(r"NEDLNK\s+DP(?:\s+(?P<terminal>\d+))?(?:\s+(?P<batch>\d+))?", text, flags=re.IGNORECASE)
	if nedlnk_match:
		return {
			"transaction_type": "NEDLNK DP",
			"bank_terminal": nedlnk_match.group("terminal"),
			"batch_number": int(nedlnk_match.group("batch")) if nedlnk_match.group("batch") else None,
		}

	return None


def find_speedpoint_match(
	worksheet,
	terminal_column: int,
	batch_number: int,
	expected_amount: float,
) -> tuple[int | None, float | None, bool]:
	candidate_row: int | None = None
	candidate_amount: float | None = None
	for row_number in range(1, worksheet.max_row + 1):
		batch_value = worksheet.cell(row=row_number, column=terminal_column).value
		if batch_value in (None, ""):
			continue
		try:
			current_batch = int(float(str(batch_value).strip()))
		except (TypeError, ValueError):
			continue
		if current_batch != batch_number:
			continue

		amount_value = parse_amount(worksheet.cell(row=row_number, column=terminal_column + 1).value)
		if amount_value == expected_amount:
			return row_number, amount_value, True
		if candidate_row is None:
			candidate_row = row_number
			candidate_amount = amount_value
	return candidate_row, candidate_amount, False


def find_speedpoint_total_match(
	worksheet,
	expected_amount: float,
	start_column: int = 7,
	end_column: int = 15,
) -> tuple[int | None, float | None, bool]:
	candidate_row: int | None = None
	candidate_total: float | None = None
	for row_number in range(1, worksheet.max_row + 1):
		row_total = 0.0
		has_amount = False
		for column_number in range(start_column, end_column + 1):
			amount = parse_amount(worksheet.cell(row=row_number, column=column_number).value)
			if amount is None:
				continue
			row_total += amount
			has_amount = True
		if not has_amount:
			continue
		row_total = round(row_total, 2)
		if row_total == expected_amount:
			return row_number, row_total, True
		if candidate_row is None:
			candidate_row = row_number
			candidate_total = row_total
	return candidate_row, candidate_total, False


def highlight_speedpoint_match(worksheet, row_number: int, terminal_column: int) -> None:
	worksheet.cell(row=row_number, column=terminal_column + 1).fill = SPEEDPOINT_HIGHLIGHT


def highlight_bank_match(worksheet, row_number: int, credit_column: int) -> None:
	worksheet.cell(row=row_number, column=credit_column).fill = BANK_HIGHLIGHT


def highlight_speedpoint_range(worksheet, row_number: int, start_column: int, end_column: int) -> None:
	for column_number in range(start_column, end_column + 1):
		worksheet.cell(row=row_number, column=column_number).fill = SPEEDPOINT_HIGHLIGHT


def process_speedpoint_reconciliation(
	speedpoint_path: Path,
	bank_paths: list[Path],
) -> dict[str, object]:
	if not bank_paths:
		raise ValueError("Please choose at least one Bank Statements .xlsx file.")
	if speedpoint_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
		raise ValueError("The Speedpoints file must be a .xlsx workbook.")
	for bank_path in bank_paths:
		if bank_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
			raise ValueError("All Bank Statements files must be .xlsx workbooks.")

	run_id = uuid4().hex
	speedpoint_original_name = speedpoint_path.name
	speedpoint_workbook = load_workbook(speedpoint_path)
	speedpoint_sheet = pick_speedpoint_sheet(speedpoint_workbook)

	unresolved_rows: list[dict[str, object]] = []
	total_matches = 0
	edited_files: list[dict[str, str]] = []

	for index, bank_path in enumerate(bank_paths, start=1):
		bank_original_name = bank_path.name
		bank_workbook = load_workbook(bank_path)
		bank_sheet = pick_bank_statement_sheet(bank_workbook)
		header_row = 2

		for row_number in range(1, min(bank_sheet.max_row, 10) + 1):
			row_values = [normalize_text(bank_sheet.cell(row_number, column).value) for column in range(1, 6)]
			if "DATE" in row_values and "TRANSACTIONS" in row_values and "CREDIT" in row_values:
				header_row = row_number
				break

		for row_number in range(header_row + 1, bank_sheet.max_row + 1):
			transaction_cell = bank_sheet.cell(row=row_number, column=2)
			credit_cell = bank_sheet.cell(row=row_number, column=4)
			transaction_info = parse_speedpoint_bank_transaction(transaction_cell.value)
			credit_amount = parse_amount(credit_cell.value)
			if not transaction_info or credit_amount is None:
				continue

			expected_amount = float(credit_amount)
			match_row: int | None = None
			match_amount: float | None = None
			exact_match = False
			reason = "No match found"

			if transaction_info["transaction_type"] == "NEDLNK DP":
				match_row, match_amount, exact_match = find_speedpoint_total_match(
					speedpoint_sheet,
					expected_amount,
					start_column=7,
					end_column=15,
				)
				if exact_match and match_row is not None:
					highlight_speedpoint_range(speedpoint_sheet, match_row, 7, 15)
					highlight_bank_match(bank_sheet, row_number, 4)
					total_matches += 1
					continue
				if match_row is not None:
					reason = "Amount mismatch"
				else:
					reason = "No matching G:O total"
			else:
				terminal_column = int(transaction_info["speedpoint_column"])
				match_row, match_amount, exact_match = find_speedpoint_match(
					speedpoint_sheet,
					terminal_column,
					int(transaction_info["batch_number"]),
					expected_amount,
				)
				if exact_match and match_row is not None:
					highlight_speedpoint_match(speedpoint_sheet, match_row, terminal_column)
					highlight_bank_match(bank_sheet, row_number, 4)
					total_matches += 1
					continue
				reason = "Batch not found"
				if match_row is not None:
					reason = "Amount mismatch"

			unresolved_rows.append(
				{
					"file_name": bank_original_name,
					"sheet_name": bank_sheet.title,
					"row_number": row_number,
					"transaction": compact_text(transaction_cell.value),
					"credit": expected_amount,
					"terminal": transaction_info.get("speedpoint_terminal") or transaction_info.get("bank_terminal") or "",
					"bank_terminal": transaction_info["bank_terminal"],
					"batch_number": transaction_info["batch_number"],
					"transaction_type": transaction_info["transaction_type"],
					"reason": reason,
					"matched_row": match_row,
					"matched_amount": match_amount,
				}
			)

		bank_workbook.save(bank_path)
		edited_files.append(
			{
				"download_name": bank_original_name,
				"stored_name": str(bank_path),
				"original_path": str(bank_path),
				"kind": "bank",
			}
		)

	speedpoint_workbook.save(speedpoint_path)
	edited_files.insert(
		0,
		{
			"download_name": speedpoint_original_name,
			"stored_name": str(speedpoint_path),
			"original_path": str(speedpoint_path),
			"kind": "speedpoint",
		},
	)

	manifest = {
		"run_id": run_id,
		"speedpoint_original_name": speedpoint_original_name,
		"speedpoint_download_name": speedpoint_original_name,
		"bank_count": len(bank_paths),
		"matched_count": total_matches,
		"unresolved_count": len(unresolved_rows),
		"unresolved_rows": unresolved_rows,
		"edited_files": edited_files,
	}
	manifest["source_folder"] = str(speedpoint_path.parent)
	manifest["output_files"] = [item["stored_name"] for item in edited_files]
	return manifest


def open_result_folder_path(manifest: dict[str, object]) -> Path:
	source_folder = manifest.get("source_folder")
	if source_folder:
		return Path(source_folder)
	raise ValueError("Result folder is not available.")
