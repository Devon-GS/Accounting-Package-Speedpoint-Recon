from __future__ import annotations

import os
import queue
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from speedpoint_core import (
	open_result_folder_path,
	process_speedpoint_reconciliation,
)


class SpeedpointApp(ctk.CTk):
	def __init__(self) -> None:
		super().__init__()
		ctk.set_appearance_mode("System")
		ctk.set_default_color_theme("blue")

		self.title("Speedpoint Reconciliation")
		self.geometry("1320x860")
		self.minsize(1120, 760)

		self.speedpoint_file: Path | None = None
		self.bank_files: list[Path] = []
		self.results: dict[str, object] | None = None
		self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()

		self._build_ui()
		self.after(100, self._poll_worker_queue)

	def _build_ui(self) -> None:
		self.grid_columnconfigure(0, weight=1)
		self.grid_rowconfigure(1, weight=1)
		self._configure_tree_style()

		controls = ctk.CTkFrame(self, corner_radius=18)
		controls.grid(row=0, column=0, padx=24, pady=(24, 16), sticky="ew")
		controls.grid_columnconfigure((0, 1), weight=1, uniform="controls")

		self._build_speedpoint_card(controls).grid(row=0, column=0, padx=(20, 10), pady=20, sticky="nsew")
		self._build_bank_card(controls).grid(row=0, column=1, padx=(10, 20), pady=20, sticky="nsew")

		action_bar = ctk.CTkFrame(self, corner_radius=18)
		action_bar.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="nsew")
		action_bar.grid_columnconfigure(0, weight=1)
		action_bar.grid_rowconfigure(1, weight=1)

		button_row = ctk.CTkFrame(action_bar, fg_color="transparent")
		button_row.grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 10))
		button_row.grid_columnconfigure(3, weight=1)

		self.process_button = ctk.CTkButton(
			button_row,
			text="Process Files",
			height=42,
			command=self.process_files,
		)
		self.process_button.grid(row=0, column=0, sticky="w")

		self.open_output_button = ctk.CTkButton(
			button_row,
			text="Open Source Folder",
			height=42,
			state="disabled",
			command=self.open_source_folder,
		)
		self.open_output_button.grid(row=0, column=1, padx=(12, 0), sticky="w")

		self.status_label = ctk.CTkLabel(button_row, text="Selected files will be edited in place.")
		self.status_label.grid(row=0, column=3, sticky="e")

		results_frame = ctk.CTkFrame(action_bar, corner_radius=14)
		results_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
		results_frame.grid_columnconfigure(0, weight=1)
		results_frame.grid_rowconfigure(1, weight=1)

		self.summary_label = ctk.CTkLabel(
			results_frame,
			text="No reconciliation has been run yet.",
			font=ctk.CTkFont(size=16, weight="bold"),
		)
		self.summary_label.grid(row=0, column=0, padx=18, pady=(18, 8), sticky="w")

		self.results_body = ctk.CTkFrame(results_frame, fg_color="transparent")
		self.results_body.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
		self.results_body.grid_columnconfigure(0, weight=1)
		self.results_body.grid_rowconfigure(0, weight=1)

		self.empty_state = ctk.CTkLabel(
			self.results_body,
			text="No unresolved transactions were found.",
			text_color=("gray35", "gray75"),
		)
		self.empty_state.grid(row=0, column=0, sticky="nsew")

		self._build_results_table(self.results_body)

	def _build_speedpoint_card(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
		card = ctk.CTkFrame(parent, corner_radius=16)
		card.grid_columnconfigure(0, weight=1)

		heading = ctk.CTkLabel(card, text="Speedpoints", font=ctk.CTkFont(size=20, weight="bold"))
		heading.grid(row=0, column=0, padx=18, pady=(18, 4), sticky="w")

		badge = ctk.CTkLabel(card, text="Single .xlsx", corner_radius=12, fg_color="#1f1f1f")
		badge.grid(row=0, column=1, padx=18, pady=(18, 4), sticky="e")

		description = ctk.CTkLabel(
			card,
			text="Upload the Speedpoints workbook that contains the terminal columns for 946389 and 946390.",
			wraplength=480,
			justify="left",
		)
		description.grid(row=1, column=0, columnspan=2, padx=18, pady=(0, 14), sticky="w")

		label = ctk.CTkLabel(card, text="Speedpoints workbook")
		label.grid(row=2, column=0, padx=18, pady=(0, 6), sticky="w")

		row = ctk.CTkFrame(card, fg_color="transparent")
		row.grid(row=3, column=0, columnspan=2, padx=18, pady=(0, 6), sticky="ew")
		row.grid_columnconfigure(0, weight=1)

		self.speedpoint_entry = ctk.CTkEntry(row, placeholder_text="No file selected")
		self.speedpoint_entry.grid(row=0, column=0, sticky="ew")
		self.speedpoint_entry.configure(state="disabled")

		browse = ctk.CTkButton(row, text="Browse", width=110, command=self.choose_speedpoint_file)
		browse.grid(row=0, column=1, padx=(12, 0))

		helper = ctk.CTkLabel(card, text="Choose one `.xlsx` file only.", text_color=("gray35", "gray75"))
		helper.grid(row=4, column=0, columnspan=2, padx=18, pady=(0, 18), sticky="w")
		return card

	def _build_bank_card(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
		card = ctk.CTkFrame(parent, corner_radius=16)
		card.grid_columnconfigure(0, weight=1)
		card.grid_rowconfigure(3, weight=1)

		heading = ctk.CTkLabel(card, text="Bank Statements", font=ctk.CTkFont(size=20, weight="bold"))
		heading.grid(row=0, column=0, padx=18, pady=(18, 4), sticky="w")

		badge = ctk.CTkLabel(card, text="Multiple .xlsx", corner_radius=12, fg_color="#1f1f1f")
		badge.grid(row=0, column=1, padx=18, pady=(18, 4), sticky="e")

		description = ctk.CTkLabel(
			card,
			text="Only the Date, Transactions, and Credit columns are used from these files.",
			wraplength=480,
			justify="left",
		)
		description.grid(row=1, column=0, columnspan=2, padx=18, pady=(0, 14), sticky="w")

		buttons = ctk.CTkFrame(card, fg_color="transparent")
		buttons.grid(row=2, column=0, columnspan=2, padx=18, pady=(0, 10), sticky="ew")
		buttons.grid_columnconfigure(1, weight=1)

		add_button = ctk.CTkButton(buttons, text="Add Files", width=110, command=self.choose_bank_files)
		add_button.grid(row=0, column=0, sticky="w")

		clear_button = ctk.CTkButton(buttons, text="Clear", width=90, command=self.clear_bank_files)
		clear_button.grid(row=0, column=1, padx=(12, 0), sticky="w")

		self.bank_list = tk_list = ctk.CTkTextbox(card, height=180, wrap="none")
		tk_list.grid(row=3, column=0, columnspan=2, padx=18, pady=(0, 18), sticky="nsew")
		tk_list.insert("1.0", "No bank statement files selected.")
		tk_list.configure(state="disabled")
		return card

	def _build_results_table(self, parent: ctk.CTkFrame) -> None:
		table_wrap = ctk.CTkFrame(parent, corner_radius=12)
		table_wrap.grid(row=0, column=0, sticky="nsew")
		table_wrap.grid_rowconfigure(0, weight=1)
		table_wrap.grid_columnconfigure(0, weight=1)

		columns = ("file", "row", "transaction", "credit", "terminal", "batch", "reason")
		self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=10)
		self.tree.heading("file", text="File")
		self.tree.heading("row", text="Row")
		self.tree.heading("transaction", text="Transaction")
		self.tree.heading("credit", text="Credit")
		self.tree.heading("terminal", text="Terminal")
		self.tree.heading("batch", text="Batch")
		self.tree.heading("reason", text="Reason")

		self.tree.column("file", width=240, anchor="w")
		self.tree.column("row", width=70, anchor="center")
		self.tree.column("transaction", width=280, anchor="w")
		self.tree.column("credit", width=90, anchor="e")
		self.tree.column("terminal", width=100, anchor="center")
		self.tree.column("batch", width=100, anchor="center")
		self.tree.column("reason", width=180, anchor="w")

		ys = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
		xs = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
		self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)

		self.tree.grid(row=0, column=0, sticky="nsew")
		ys.grid(row=0, column=1, sticky="ns")
		xs.grid(row=1, column=0, sticky="ew")

	def _configure_tree_style(self) -> None:
		style = ttk.Style(self)
		try:
			style.theme_use("clam")
		except tk.TclError:
			pass
		style.configure(
			"Treeview",
			rowheight=28,
			font=("Segoe UI", 10),
			foreground="#1f1f1f",
			background="#ffffff",
			fieldbackground="#ffffff",
		)
		style.configure(
			"Treeview.Heading",
			font=("Segoe UI", 10, "bold"),
			padding=(8, 6),
		)

	def choose_speedpoint_file(self) -> None:
		path = filedialog.askopenfilename(
			title="Choose Speedpoints workbook",
			filetypes=[("Excel workbook", "*.xlsx")],
		)
		if not path:
			return
		selected = Path(path)
		if selected.suffix.lower() != ".xlsx":
			messagebox.showwarning("Invalid file", "The Speedpoints file must be a .xlsx workbook.")
			return
		self.speedpoint_file = selected
		self._set_entry_text(self.speedpoint_entry, str(selected))

	def choose_bank_files(self) -> None:
		paths = filedialog.askopenfilenames(
			title="Choose Bank Statements workbooks",
			filetypes=[("Excel workbooks", "*.xlsx")],
		)
		if not paths:
			return
		selected_paths = []
		for item in paths:
			selected = Path(item)
			if selected.suffix.lower() != ".xlsx":
				continue
			if selected not in selected_paths:
				selected_paths.append(selected)
		self.bank_files = selected_paths
		self._refresh_bank_list()

	def clear_bank_files(self) -> None:
		self.bank_files = []
		self._refresh_bank_list()

	def _set_entry_text(self, entry: ctk.CTkEntry, value: str) -> None:
		entry.configure(state="normal")
		entry.delete(0, "end")
		entry.insert(0, value)
		entry.configure(state="disabled")

	def _refresh_bank_list(self) -> None:
		self.bank_list.configure(state="normal")
		self.bank_list.delete("1.0", "end")
		if self.bank_files:
			for path in self.bank_files:
				self.bank_list.insert("end", f"{path}\n")
		else:
			self.bank_list.insert("1.0", "No bank statement files selected.")
		self.bank_list.configure(state="disabled")

	def process_files(self) -> None:
		if self.speedpoint_file is None:
			messagebox.showwarning("Missing file", "Please choose one Speedpoints .xlsx file.")
			return
		if not self.bank_files:
			messagebox.showwarning("Missing files", "Please choose at least one Bank Statements .xlsx file.")
			return

		self.process_button.configure(state="disabled")
		self.open_output_button.configure(state="disabled")
		self.summary_label.configure(text="Processing files...")
		self._clear_results_table()

		worker = threading.Thread(target=self._worker_process, daemon=True)
		worker.start()

	def _worker_process(self) -> None:
		try:
			manifest = process_speedpoint_reconciliation(self.speedpoint_file, self.bank_files, self.output_root)
		except Exception as exc:  # noqa: BLE001 - surfaced to the UI
			self.worker_queue.put(("error", exc))
			return
		self.worker_queue.put(("result", manifest))

	def _poll_worker_queue(self) -> None:
		try:
			while True:
				kind, payload = self.worker_queue.get_nowait()
				if kind == "error":
					self._handle_error(payload)
				else:
					self._handle_result(payload)
		except queue.Empty:
			pass
		finally:
			self.after(120, self._poll_worker_queue)

	def _handle_error(self, exc: Exception) -> None:
		self.process_button.configure(state="normal")
		messagebox.showerror("Processing failed", str(exc))
		self.summary_label.configure(text="Processing failed.")

	def _handle_result(self, manifest: dict[str, object]) -> None:
		self.results = manifest
		self.process_button.configure(state="normal")
		self.open_output_button.configure(state="normal")

		matched = int(manifest.get("matched_count", 0))
		unresolved = int(manifest.get("unresolved_count", 0))
		output_files = manifest.get("output_files", [])
		self.summary_label.configure(
			text=f"Matched {matched} transaction{'s' if matched != 1 else ''} and found {unresolved} unresolved item{'s' if unresolved != 1 else ''}. The selected files were saved back in place."
		)
		if output_files:
			self.status_label.configure(text=f"Saved {len(output_files)} file{'s' if len(output_files) != 1 else ''} in place.")
		else:
			self.status_label.configure(text="Selected files were edited in place.")

		rows = manifest.get("unresolved_rows", [])
		if rows:
			self.empty_state.grid_remove()
			for row in rows:
				self.tree.insert(
					"",
					"end",
					values=(
						row.get("file_name", ""),
						row.get("row_number", ""),
						row.get("transaction", ""),
						f'{float(row.get("credit", 0.0)):.2f}',
						row.get("terminal", ""),
						row.get("batch_number", ""),
						row.get("reason", ""),
					),
				)
		else:
			self.empty_state.grid()
		self._refresh_results_visibility(rows)

	def _refresh_results_visibility(self, rows: list[dict[str, object]]) -> None:
		if rows:
			self.tree.master.grid()
		else:
			self.tree.master.grid_remove()

	def _clear_results_table(self) -> None:
		for item in self.tree.get_children():
			self.tree.delete(item)
		self.empty_state.grid()
		self.tree.master.grid_remove()

	def open_source_folder(self) -> None:
		if not self.results:
			return
		try:
			folder = open_result_folder_path(self.results)
		except Exception as exc:  # noqa: BLE001 - surfaced to the UI
			messagebox.showerror("Cannot open folder", str(exc))
			return
		if os.name == "nt":
			os.startfile(folder)  # type: ignore[attr-defined]
		else:
			import subprocess

			opener = "open" if sys.platform == "darwin" else "xdg-open"
			subprocess.Popen([opener, str(folder)])  # noqa: S603,S607 - local GUI action


def main() -> None:
	app = SpeedpointApp()
	app.mainloop()


if __name__ == "__main__":
	main()
