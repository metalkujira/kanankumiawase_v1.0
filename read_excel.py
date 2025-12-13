from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl
import pandas as pd


def main() -> int:
	parser = argparse.ArgumentParser(description="Quickly inspect チームリスト.xlsx")
	parser.add_argument(
		"--file",
		type=str,
		default="チームリスト.xlsx",
		help="Path to チームリスト.xlsx (default: チームリスト.xlsx)",
	)
	args = parser.parse_args()

	file_path = Path(args.file)
	if not file_path.exists():
		raise FileNotFoundError(f"File not found: {file_path}")

	wb = openpyxl.load_workbook(str(file_path), read_only=True)
	sheet = wb.active
	data = list(sheet.values)
	df = pd.DataFrame(data[1:], columns=data[0])

	print(f"{file_path}:")
	print(df.head(20))
	print("Total teams:", len(df))
	if "レベル" in df.columns:
		print("Level counts:")
		print(df["レベル"].value_counts())
	return 0


if __name__ == "__main__":
	raise SystemExit(main())