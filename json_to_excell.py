# json_to_excell.py

import json
import itertools
import os
import pathlib
import re
import sys
import traceback
from typing import List, Tuple, Dict, Any

import pandas as pd


def load_json_safely(path: pathlib.Path) -> List[dict]:
    text = path.read_text(encoding="utf-8", errors="replace").strip()

    try:
        data = json.loads(text)
        if isinstance(data, dict):
            data = [data]
        elif not isinstance(data, list):
            data = [dict(value=data)]
        return data
    except json.JSONDecodeError:
        pass

    records = []
    for i, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line:
            continue
        try:
            obj = json.loads(line)
        except json.JSONDecodeError as e:
            raise ValueError(f"NDJSON parse error in {path.name} at line {i}: {e}") from e
        if isinstance(obj, dict):
            records.append(obj)
        else:
            records.append({"value": obj})
    return records


def safe_base_name(stem: str) -> str:
    name = re.sub(r'[\[\]:*?/\\]', '_', stem)
    name = name.strip() or "file"
    return name


def make_sheet_name(prefix: str, local: str, used: set) -> str:
    raw = f"{prefix}{local}" if local else prefix
    raw = re.sub(r'[\[\]:*?/\\]', '_', raw)
    if not raw:
        raw = "sheet"

    base = raw[:31]
    if base not in used:
        used.add(base)
        return base

    root = base[:28]
    for n in itertools.count(2):
        cand = f"{root}_{n}"[:31]
        if cand not in used:
            used.add(cand)
            return cand


def flatten_records(records: List[dict]) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    main_rows = []
    child_tables = {}

    def add_child(sheet_key: str, row: Dict[str, Any]):
        child_tables.setdefault(sheet_key or "list", []).append(row)

    for ridx, rec in enumerate(records):
        base = {}
        stack = [((), rec)]

        while stack:
            path, val = stack.pop()
            if isinstance(val, dict):
                for k, v in val.items():
                    stack.append((path + (str(k),), v))
            elif isinstance(val, list):
                sheet_key = ".".join(path) if path else "list"
                for i, item in enumerate(val):
                    row = {"_parent_id": ridx, "index": i, "_path": ".".join(path)}
                    if isinstance(item, dict):
                        row.update(item)
                    elif isinstance(item, list):
                        row["value"] = item
                    else:
                        row["value"] = item
                    add_child(sheet_key, row)
            else:
                col = ".".join(path) if path else "root"
                base[col] = val

        main_rows.append({"__id": ridx, **base})

    main_df = pd.json_normalize(main_rows, sep=".")

    child_dfs = {}
    for sheet_key, rows in child_tables.items():
        child_dfs[sheet_key] = pd.json_normalize(rows, sep=".")

    return main_df, child_dfs


def find_json_files(root: pathlib.Path) -> List[pathlib.Path]:
    exts = (".json", ".jsonl", ".ndjson")
    return sorted([p for p in root.iterdir() if p.is_file() and p.suffix.lower() in exts])


def choose_engine():
    try:
        import xlsxwriter
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl
            return "openpyxl"
        except Exception:
            raise RuntimeError(
                "No Excel writer engine found. Install one:\n"
                "  pip install xlsxwriter\n"
                "or\n"
                "  pip install openpyxl"
            )


def process_one_directory(dir_path: pathlib.Path, service_root: pathlib.Path) -> bool:
    """
    Write Excel files into service_root (the <service>-results directory)
    instead of inside dir_path (the score folder).
    """
    json_files = find_json_files(dir_path)
    if not json_files:
        return False

    engine = choose_engine()
    used_sheet_names = set()

    out_path = service_root / f"{dir_path.name}.xlsx"

    print(f"\nDirectory: {dir_path}")
    print(f"Found {len(json_files)} JSON file(s). Writing -> {out_path.name} (engine: {engine})")

    with pd.ExcelWriter(out_path, engine=engine) as xw:
        for jpath in json_files:
            try:
                prefix = safe_base_name(jpath.stem)[:12]
                records = load_json_safely(jpath)
                main_df, child_dfs = flatten_records(records)

                main_name = make_sheet_name(prefix, "main", used_sheet_names)
                main_df.to_excel(xw, sheet_name=main_name, index=False)

                for local_key, df in child_dfs.items():
                    local_short = local_key.replace(".", "_")[:16] or "data"
                    sheet = make_sheet_name(prefix, local_short, used_sheet_names)
                    df.to_excel(xw, sheet_name=sheet, index=False)

                print(f"  ✓ {jpath.name} -> sheets starting with '{prefix}*'")
            except Exception:
                print(f"  ✗ Failed on {jpath.name}")
                traceback.print_exc(file=sys.stdout)

    print(f"Done: {out_path}")
    return True


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", dest="input_dir", default=None,
                        help="Service results directory to process")
    args, _ = parser.parse_known_args()

    if args.input_dir:
        service_root = pathlib.Path(args.input_dir).resolve()
    else:
        service_root = pathlib.Path(__file__).resolve().parent

    wrote_any = process_one_directory(service_root, service_root)

    for dirpath, dirnames, filenames in os.walk(service_root):
        d = pathlib.Path(dirpath)
        if d == service_root:
            continue
        if d.name.startswith('.'):
            continue
        wrote = process_one_directory(d, service_root)
        wrote_any = wrote_any or wrote

    if not wrote_any:
        print("No .json/.jsonl/.ndjson files found in this tree.")


if __name__ == "__main__":
    main()

 
