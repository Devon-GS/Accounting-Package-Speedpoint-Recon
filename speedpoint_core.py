from __future__ import annotations

import os
import hashlib
import json
import logging
from logging.handlers import RotatingFileHandler
import re
from collections import Counter
from pathlib import Path
from datetime import date, datetime
from uuid import uuid4
import sys

import keyring
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

try:
	import psycopg
	from psycopg.rows import dict_row
except ImportError:  # pragma: no cover - optional dependency for database matching
	psycopg = None
	dict_row = None

def get_app_root() -> Path:
	if getattr(sys, "frozen", False):
		return Path(sys.executable).resolve().parent
	return Path(__file__).resolve().parent


BASE_DIR = get_app_root()
SUPPORTED_EXTENSIONS = {".xlsx"}
SPEEDPOINT_HIGHLIGHT = PatternFill(fill_type="solid", fgColor="FF86B9")
BANK_HIGHLIGHT = PatternFill(fill_type="solid", fgColor="FFFF00")
OFX_HIGHLIGHT = PatternFill(fill_type="solid", fgColor="FF00B0F0")
IGNORED_SUPPLIER_NAMES = {
	"SASOL OIL 0000015733",
	"FIGMENT SASOLE2607 SL32",
}
APP_NAME = "SpeedpointRecon"
load_dotenv()


class SingleFileRotatingHandler(RotatingFileHandler):
	"""Rotating handler that truncates the log file when it hits maxBytes, keeping strictly one file."""

	def doRollover(self):
		if self.stream:
			self.stream.close()
			self.stream = None
		# Clear the current file so no backup files are created and size stays <= maxBytes
		with open(self.baseFilename, "w", encoding=self.encoding):
			pass
		if not self.delay:
			self.stream = self._open()


def setup_logger() -> logging.Logger:
	app_logger = logging.getLogger("speedpoint")
	if not app_logger.handlers:
		app_logger.setLevel(logging.DEBUG)

		formatter = logging.Formatter(
			"%(asctime)s [%(levelname)s] %(message)s",
			datefmt="%Y-%m-%d %H:%M:%S",
		)

		# Single log file capped at 10MB
		log_path = BASE_DIR / "speedpoint.log"
		file_handler = SingleFileRotatingHandler(
			log_path,
			maxBytes=10 * 1024 * 1024,
			backupCount=0,
			encoding="utf-8",
		)
		file_handler.setLevel(logging.DEBUG)
		file_handler.setFormatter(formatter)
		app_logger.addHandler(file_handler)

		# # Also output to console
		# console_handler = logging.StreamHandler(sys.stdout)
		# console_handler.setLevel(logging.DEBUG)
		# console_handler.setFormatter(formatter)
		# app_logger.addHandler(console_handler)

	return app_logger


logger = setup_logger()


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


def normalize_date_key(value: object) -> str | None:
	if value in (None, ""):
		return None
	if isinstance(value, datetime):
		return value.date().isoformat()
	if isinstance(value, date):
		return value.isoformat()
	if isinstance(value, str):
		text = value.strip()
		for fmt in (
			"%Y-%m-%d",
			"%Y/%m/%d",
			"%d/%m/%Y",
			"%d-%m-%Y",
			"%d %b %Y",
			"%d %B %Y",
		):
			try:
				return datetime.strptime(text, fmt).date().isoformat()
			except ValueError:
				pass
		try:
			return date.fromisoformat(text).isoformat()
		except ValueError:
			pass
		try:
			return datetime.fromisoformat(text).date().isoformat()
		except ValueError:
			pass
	return None


def parse_money(value: object) -> float | None:
	if value in (None, ""):
		return None
	if isinstance(value, str):
		text = value.strip()
		negative = text.startswith("(") and text.endswith(")")
		if negative:
			text = text[1:-1]
		cleaned = re.sub(r"[^0-9.\-]", "", text)
		if not cleaned:
			return None
		try:
			number = float(cleaned)
		except ValueError:
			return None
		if negative and number > 0:
			number = -number
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


def normalize_amount_key(value: object) -> float | None:
	number = parse_money(value)
	if number is None:
		return None
	return round(abs(number), 2)


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


def build_database_config() -> dict[str, object]:
	database_url = keyring.get_password(APP_NAME, "DATABASE_URL")

	if not database_url:
		database_url = os.environ.get("DATABASE_URL", "").strip()

	if database_url:
		return {"dsn": database_url}

	return {
		"host": os.environ.get("PGHOST") or os.environ.get("DB_HOST") or "localhost",
		"port": int(os.environ.get("PGPORT") or os.environ.get("DB_PORT_BINDING") or "5432"),
		"dbname": os.environ.get("PGDATABASE") or os.environ.get("DB_NAME") or "accounting",
		"user": os.environ.get("PGUSER") or os.environ.get("DB_USER") or "accounting",
		"password": os.environ.get("PGPASSWORD") or os.environ.get("DB_PASSWORD") or "",
	}


def connect_database():
	if psycopg is None or dict_row is None:
		raise RuntimeError(
			"Database support requires psycopg. Install the project dependencies first."
		)

	config = build_database_config()
	if "dsn" in config:
		return psycopg.connect(config["dsn"], row_factory=dict_row)
	return psycopg.connect(row_factory=dict_row, **config)


def check_database_connection() -> tuple[bool, str]:
	try:
		connection = connect_database()
	except Exception as exc:  # noqa: BLE001 - surfaced in the UI as a status message
		return False, str(exc)

	try:
		with connection:
			pass
	except Exception as exc:  # noqa: BLE001 - surfaced in the UI as a status message
		return False, str(exc)
	return True, "Connected"


def fetch_ofx_payments() -> list[dict[str, object]]:
	query = """
		SELECT payment_date, supplier_name, amount_paid, source_sheet
		FROM payments
		WHERE UPPER(TRIM(source_sheet)) = 'OFX'
		ORDER BY payment_date, supplier_name, amount_paid
	"""
	with connect_database() as connection:
		with connection.cursor() as cursor:
			cursor.execute(query)
			rows = cursor.fetchall()
	return [dict(row) for row in rows]


def build_payment_signature(payment_date: object, supplier_name: object, amount_paid: object) -> tuple[str, str, float] | None:
	supplier_key = normalize_text(supplier_name)
	if not supplier_key or supplier_key in IGNORED_SUPPLIER_NAMES:
		return None

	date_key = normalize_date_key(payment_date)
	amount_key = normalize_amount_key(amount_paid)
	if not date_key or amount_key is None:
		return None

	return date_key, supplier_key, amount_key


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


def highlight_ofx_payment_match(worksheet, row_number: int) -> None:
	for column_number in range(3, 4):
		worksheet.cell(row=row_number, column=column_number).fill = OFX_HIGHLIGHT


def process_speedpoint_reconciliation(
	speedpoint_paths: list[Path] | Path,
	bank_paths: list[Path],
) -> dict[str, object]:
	if not bank_paths:
		raise ValueError("Please choose at least one Bank Statements .xlsx file.")
	if isinstance(speedpoint_paths, Path):
		speedpoint_paths = [speedpoint_paths]
	if not speedpoint_paths:
		raise ValueError("Please choose at least one Speedpoints .xlsx file.")
	for speedpoint_path in speedpoint_paths:
		if speedpoint_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
			raise ValueError("All Speedpoints files must be .xlsx workbooks.")
	for bank_path in bank_paths:
		if bank_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
			raise ValueError("All Bank Statements files must be .xlsx workbooks.")

	run_id = uuid4().hex
	speedpoint_workbooks = []
	for speedpoint_path in speedpoint_paths:
		speedpoint_workbook = load_workbook(speedpoint_path)
		speedpoint_workbooks.append(
			(speedpoint_path, speedpoint_workbook, pick_speedpoint_sheet(speedpoint_workbook))
		)

	logger.info("=" * 70)
	logger.info("Starting reconciliation run %s", run_id)
	logger.debug("Fetching OFX payments from database...")
	ofx_payments = fetch_ofx_payments()
	logger.debug("Total OFX rows returned from database: %d", len(ofx_payments))

	ofx_payment_counts: Counter[tuple[str, str, float]] = Counter()
	db_signatures_list: list[tuple[str, str, float]] = []

	for payment in ofx_payments:
		raw_d = payment.get("payment_date")
		raw_s = payment.get("supplier_name")
		raw_a = payment.get("amount_paid")
		signature = build_payment_signature(raw_d, raw_s, raw_a)
		if signature is None:
			logger.debug(
				"Could not build signature for DB payment: date=%r, supplier=%r, amount=%r",
				raw_d,
				raw_s,
				raw_a,
			)
			continue
		ofx_payment_counts[signature] += 1
		db_signatures_list.append(signature)

	logger.debug(
		"Active unique DB OFX signatures: %d (total count: %d)",
		len(ofx_payment_counts),
		sum(ofx_payment_counts.values()),
	)
	if ofx_payment_counts:
		sample_lines = "\n".join(
			f"  Date: {sig[0]} | Supplier: {sig[1]!r} | Amount: {sig[2]}"
			for sig in list(ofx_payment_counts.keys())[:5]
		)
		logger.debug("Sample DB signatures:\n%s", sample_lines)

	unresolved_rows: list[dict[str, object]] = []
	total_matches = 0
	ofx_payment_matches = 0
	edited_files: list[dict[str, str]] = []

	for index, bank_path in enumerate(bank_paths, start=1):
		bank_original_name = bank_path.name
		bank_workbook = load_workbook(bank_path)
		bank_sheet = pick_bank_statement_sheet(bank_workbook)
		header_row = 2

		logger.info("Processing bank statement: %s (Sheet: %s)", bank_original_name, bank_sheet.title)
		for row_number in range(1, min(bank_sheet.max_row, 10) + 1):
			row_values = [normalize_text(bank_sheet.cell(row_number, column).value) for column in range(1, 6)]
			if "DATE" in row_values and "TRANSACTIONS" in row_values and "CREDIT" in row_values:
				header_row = row_number
				break

		logger.debug("Header row identified at line: %d", header_row)

		for row_number in range(header_row + 1, bank_sheet.max_row + 1):
			c1_val = bank_sheet.cell(row=row_number, column=1).value
			c2_val = bank_sheet.cell(row=row_number, column=2).value
			c3_val = bank_sheet.cell(row=row_number, column=3).value
			c4_val = bank_sheet.cell(row=row_number, column=4).value

			# Skip rows completely blank in the first 4 columns
			if not any([c1_val, c2_val, c3_val, c4_val]):
				continue

			payment_signature = build_payment_signature(c1_val, c2_val, c3_val)

			if payment_signature is not None and ofx_payment_counts.get(payment_signature, 0) > 0:
				logger.info("Row %d: MATCH FOUND for %s", row_number, payment_signature)
				highlight_ofx_payment_match(bank_sheet, row_number)
				ofx_payment_counts[payment_signature] -= 1
				if ofx_payment_counts[payment_signature] <= 0:
					del ofx_payment_counts[payment_signature]
				ofx_payment_matches += 1
			else:
				amt_c3 = normalize_amount_key(c3_val)
				amt_c4 = normalize_amount_key(c4_val)
				date_norm = normalize_date_key(c1_val)
				desc_norm = normalize_text(c2_val)

				target_amount = amt_c3 or amt_c4
				near_misses = [
					sig for sig in db_signatures_list
					if target_amount is not None and sig[2] == target_amount
				]

				if near_misses:
					candidate_details = "\n".join(
						f"  DB Record: Date={cand[0]} [{'MATCH' if cand[0] == date_norm else 'MISMATCH'}], "
						f"Supplier={cand[1]!r} [{'MATCH' if cand[1] == desc_norm else 'MISMATCH'}], Amount={cand[2]}"
						for cand in near_misses[:3]
					)
					logger.warning(
						"Row %d failed to match, but DB has record(s) with same amount (%s):\n"
						"  Bank Row : Date=%r (norm: %s) | Desc=%r (norm: %s) | Col3=%r | Col4=%r\n"
						"  Bank Sig : %s\n%s",
						row_number,
						target_amount,
						c1_val,
						date_norm,
						c2_val,
						desc_norm,
						c3_val,
						c4_val,
						payment_signature,
						candidate_details,
					)
				elif payment_signature is not None:
					date_matches = [sig for sig in db_signatures_list if sig[0] == date_norm]
					if len(date_matches) <= 2 and date_matches:
						logger.debug(
							"Row %d: Bank Sig=%s not in DB. DB entries on this date: %s",
							row_number,
							payment_signature,
							date_matches,
						)

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

			for speedpoint_path, speedpoint_workbook, speedpoint_sheet in speedpoint_workbooks:
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
						break
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
						break
					reason = "Batch not found"
					if match_row is not None:
						reason = "Amount mismatch"
			if exact_match and match_row is not None:
				continue

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

	for speedpoint_path, speedpoint_workbook, _ in speedpoint_workbooks:
		speedpoint_workbook.save(speedpoint_path)
		edited_files.insert(
			0,
			{
				"download_name": speedpoint_path.name,
				"stored_name": str(speedpoint_path),
				"original_path": str(speedpoint_path),
				"kind": "speedpoint",
			},
		)

	logger.info("=" * 70)
	logger.info("OFX Matches: %d", ofx_payment_matches)
	logger.info("Unmatched OFX payments remaining in DB: %d", sum(ofx_payment_counts.values()))
	if ofx_payment_counts:
		remaining_lines = "\n".join(
			f"  Remaining (count={count}): Date={sig[0]} | Supplier={sig[1]!r} | Amount={sig[2]}"
			for sig, count in ofx_payment_counts.items()
		)
		logger.info("Remaining unmatched DB records:\n%s", remaining_lines)
	logger.info("=" * 70)

	manifest = {
		"run_id": run_id,
		"speedpoint_original_name": speedpoint_paths[0].name,
		"speedpoint_download_name": speedpoint_paths[0].name,
		"bank_count": len(bank_paths),
		"matched_count": total_matches,
		"ofx_payment_matched_count": ofx_payment_matches,
		"ofx_payment_unmatched_count": sum(ofx_payment_counts.values()),
		"ofx_payment_ignored_count": sum(
			1
			for payment in ofx_payments
			if normalize_text(payment.get("supplier_name")) in IGNORED_SUPPLIER_NAMES
		),
		"unresolved_count": len(unresolved_rows),
		"unresolved_rows": unresolved_rows,
		"edited_files": edited_files,
	}
	manifest["output_files"] = [item["stored_name"] for item in edited_files]
	return manifest
