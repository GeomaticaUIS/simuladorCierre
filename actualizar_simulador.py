"""Refresh the simulator's embedded data from the project Excel workbook.

Install the dependency with ``python -m pip install openpyxl``, then run:
    python actualizar_simulador.py

The workbook and HTML paths can be overridden with --excel and --html.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import tempfile
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


DEFAULT_WORKBOOK = "ProyeccionCierreProyecto_TendencialXSitio_13Nov2026.xlsx"
DEFAULT_JS = "app.js"
SHEET_NAME = "PLAN DE CIERRE FNTC 613-2023"
DATA_PATTERN = re.compile(
    r"(const\s+DATA\s*=\s*)\{.*?\}(;)(?=\s*(?:const\s+byId\b|const\s+byId\s*=))",
    re.DOTALL,
)

SITE_COLUMNS = {
    "instalacion": "INSTALACION",
    "codmun": "CODMUN",
    "cod_cotecmar": "COD COTECMAR",
    "item": "ITEM",
    "aa": "AUTORIDAD AMBIENTAL",
    "at": "AUTORIDAD DE TRANSPORTE",
    "municipio": "MUNICIPIO",
    "departamento": "DEPARTAMENTO",
    "convenio": "FECHA CONVENIO",
    "remision_aa": "REMISION CORREO AA",
    "auto_inicio": "AUTO DE INICIO",
    "visita": "VISITA TECNICA",
    "res_aa": "RESOLUCION EMITIDA AA",
    "remision_at": "REMISION CORREO AT",
    "res_at": "RESOLUCION EMITIDA AT",
    "estado": "ESTADO ACTUAL",
    "solicitud_poc": "SOLICITUD POC",
    "poc": "PROYECCION POC",
    "acor": "PROYECCION ACOR",
    "instalacion_proj": "PROYECCION INSTALACION",
    "buffer_inst": "DIAS INSTAL",
}


def normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(re.sub(r"[^a-z0-9]+", " ", text.lower()).split())


def excel_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def find_column(headers: dict[str, int], name: str) -> int:
    normalized = normalize(name)
    try:
        return headers[normalized]
    except KeyError as error:
        raise ValueError(f"Required Excel column not found: {name}") from error


def header_map(sheet: Any, row: int, start_column: int, end_column: int) -> dict[str, int]:
    return {
        normalize(sheet.cell(row, column).value): column
        for column in range(start_column, end_column + 1)
        if sheet.cell(row, column).value is not None
    }


def find_history_table(sheet: Any, required_headers: set[str]) -> tuple[int, dict[str, int]]:
    for row in range(1, sheet.max_row + 1):
        headers = header_map(sheet, row, 5, 8)
        if required_headers.issubset(headers):
            return row, headers
    expected = ", ".join(sorted(required_headers))
    raise ValueError(f"Could not find historical table with headers: {expected}")


def identity(site: dict[str, Any]) -> tuple[str, str, str, str]:
    return (
        normalize(site.get("instalacion")),
        normalize(site.get("codmun")),
        normalize(site.get("municipio")),
        normalize(site.get("departamento")),
    )


def load_existing_data(html: str) -> tuple[dict[str, Any], re.Match[str]]:
    match = DATA_PATTERN.search(html)
    if match is None:
        raise ValueError("Could not find the embedded `const DATA = {...};` block in the HTML")
    try:
        data_start = len(match.group(1))
        data_end = -len(match.group(2))
        return json.loads(match.group(0)[data_start:data_end]), match
    except json.JSONDecodeError as error:
        raise ValueError("The embedded DATA block is not valid JSON") from error


def build_historical_data(sheet: Any) -> tuple[dict[str, Any], dict[str, Any]]:
    aa_headers = {
        "corporacion",
        "plazo",
        "plazo medio",
        "plazo minimo",
    }
    aa_row, aa_columns = find_history_table(sheet, aa_headers)
    aa: dict[str, Any] = {}
    for row in range(aa_row + 1, sheet.max_row + 1):
        name = sheet.cell(row, aa_columns["corporacion"]).value
        if not name or normalize(name).startswith("plazos en"):
            break
        aa[str(name).strip()] = {
            "plazo": excel_value(sheet.cell(row, aa_columns["plazo"]).value),
            "medio": excel_value(sheet.cell(row, aa_columns["plazo medio"]).value),
            "minimo": excel_value(sheet.cell(row, aa_columns["plazo minimo"]).value),
        }

    at_headers = {
        "corporacion",
        "plazo promedio",
        "plazo maximo",
        "plazo minimo",
    }
    at_row, at_columns = find_history_table(sheet, at_headers)
    at: dict[str, Any] = {}
    for row in range(at_row + 1, sheet.max_row + 1):
        name = sheet.cell(row, at_columns["corporacion"]).value
        if not name or normalize(name).startswith("plazos en"):
            break
        at[str(name).strip()] = {
            "medio": excel_value(sheet.cell(row, at_columns["plazo promedio"]).value),
            "maximo": excel_value(sheet.cell(row, at_columns["plazo maximo"]).value),
            "minimo": excel_value(sheet.cell(row, at_columns["plazo minimo"]).value),
        }

    if not aa or not at:
        raise ValueError("One or both historical tables are empty")
    return aa, at


def build_sites(sheet: Any, existing_sites: list[dict[str, Any]]) -> list[dict[str, Any]]:
    headers = header_map(sheet, 2, 1, sheet.max_column)
    columns = {key: find_column(headers, name) for key, name in SITE_COLUMNS.items()}

    existing_coordinates = {
        identity(site): (site.get("lat"), site.get("lon"))
        for site in existing_sites
        if site.get("lat") is not None and site.get("lon") is not None
    }

    sites = []
    for row in range(3, sheet.max_row + 1):
        installation = sheet.cell(row, columns["instalacion"]).value
        if installation is None:
            continue

        site: dict[str, Any] = {"row": row}
        for key, column in columns.items():
            value = excel_value(sheet.cell(row, column).value)
            if key in {"aa", "at", "municipio", "departamento"}:
                value = str(value).strip() if value is not None else None
            elif key == "estado":
                value = str(value).strip() if value is not None else ""
            site[key] = value

        lat, lon = existing_coordinates.get(identity(site), (None, None))
        site["lat"] = lat
        site["lon"] = lon
        sites.append(site)

    if not sites:
        raise ValueError("No site records were found in the Excel sheet")
    if any(site["instalacion_proj"] is None for site in sites):
        raise ValueError("At least one site has no installation projection date")
    return sites


def simulation_cutoff(sheet: Any, sites: list[dict[str, Any]]) -> str:
    projection_floor = min(date.fromisoformat(site["instalacion_proj"]) for site in sites)
    report_header = str(sheet["I1"].value or "")
    report_date_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", report_header)
    if report_date_match is None:
        raise ValueError("Could not determine the report cutoff date from cell I1")
    day, month, year = map(int, report_date_match.groups())
    report_date = date(year, month, day)
    return max(report_date, projection_floor).isoformat()


def update_javascript(workbook_path: Path, javascript_path: Path) -> tuple[int, int, int]:
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel workbook does not exist: {workbook_path}")
    if not javascript_path.is_file():
        raise FileNotFoundError(f"JavaScript file does not exist: {javascript_path}")

    javascript = javascript_path.read_text(encoding="utf-8")
    existing_data, match = load_existing_data(javascript)
    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        if SHEET_NAME not in workbook.sheetnames:
            raise ValueError(f"Worksheet not found: {SHEET_NAME}")
        sheet = workbook[SHEET_NAME]
        sites = build_sites(sheet, existing_data.get("sites", []))
        aa, at = build_historical_data(sheet)
        data = {
            "cutoff": simulation_cutoff(sheet, sites),
            "sites": sites,
            "aa": aa,
            "at": at,
        }
    finally:
        workbook.close()

    replacement = match.group(1) + json.dumps(data, ensure_ascii=False, indent=2) + match.group(2)
    updated_javascript = javascript[: match.start()] + replacement + javascript[match.end() :]

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="",
            dir=javascript_path.parent,
            prefix=f".{javascript_path.name}.",
            suffix=".tmp",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
            temporary_file.write(updated_javascript)
        os.replace(temporary_path, javascript_path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()

    matched_coordinates = sum(site["lat"] is not None and site["lon"] is not None for site in sites)
    return len(sites), len(aa), matched_coordinates


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Read the project Excel workbook and refresh the embedded data in the simulator JavaScript."
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=Path(__file__).resolve().parent / DEFAULT_WORKBOOK,
        help=f"Source workbook (default: {DEFAULT_WORKBOOK})",
    )
    parser.add_argument(
        "--js",
        type=Path,
        default=Path(__file__).resolve().parent / DEFAULT_JS,
        help=f"Simulator JavaScript to update (default: {DEFAULT_JS})",
    )
    args = parser.parse_args()

    sites, aa_count, coordinate_count = update_javascript(args.excel, args.js)
    print(
        f"Updated {args.js}: {sites} sites, {aa_count} environmental authorities, "
        f"and {coordinate_count} retained coordinate pairs."
    )
    if coordinate_count < sites:
        print(
            "Some coordinates were unavailable in the workbook and did not match an existing "
            "HTML site by installation, municipality, and department; those map points remain unset."
        )


if __name__ == "__main__":
    main()
