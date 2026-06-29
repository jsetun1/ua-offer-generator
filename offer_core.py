"""Core logic for UA Offer Generator."""

from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Sequence

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# "Artikl" is the UA style + colour key, e.g. 6009833-001.
OUTPUT_COLUMNS = [
    "Artikl",
    "Size",
    "Název",
    "Local warehouse 101",
    "Local warehouse 501",
    "Central warehouse",
    "Celkem ks styl/barva",
    "Plný sizerun",
    "Top style",
    "Materiál",
    "ORDER",
    "MOC CZK",
    "MOC EUR",
    "EAN",
    "Gender",
    "Silhouette",
    "Fit",
    "End use",
    "Season",
    "C/O",
]

EXTRA_COLUMNS = [
    "Dostupnost ve vybraných skladech",
    "Total available",
    "Dostupné velikosti styl/barva",
    "Style code",
    "Color group",
    "Color name",
    "Color code",
    "Detail silhouette",
    "Composition",
]

SIZE_ORDER = {
    "XXS": 10,
    "XS": 20,
    "SM": 30,
    "S": 30,
    "MD": 40,
    "M": 40,
    "LG": 50,
    "L": 50,
    "XL": 60,
    "XXL": 70,
    "2XL": 70,
    "XXXL": 80,
    "3XL": 80,
    "4XL": 90,
    "5XL": 100,
    "OSFA": 900,
    "OSFM": 900,
}

TRUE_MARKERS = {"1", "true", "yes", "y", "ano", "x", "top", "top style"}
CORE_SIZE_RUNS = {
    "mens": {"S", "M", "L", "XL", "2XL"},
    "womens": {"XS", "S", "M", "L", "XL"},
}


def normalize_ean(value) -> str:
    """Return a safe text EAN without decimals or spaces."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text or text.casefold() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits if len(digits) >= 8 else text


def read_excel(path_or_buffer) -> pd.DataFrame:
    """Read the first worksheet into a DataFrame."""
    from openpyxl import load_workbook

    wb = load_workbook(path_or_buffer, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    try:
        header_row = next(rows)
    except StopIteration:
        return pd.DataFrame()

    headers: list[str] = []
    seen: dict[str, int] = {}
    for idx, value in enumerate(header_row, start=1):
        name = str(value).strip() if value is not None else f"Column {idx}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        headers.append(name)

    return pd.DataFrame(list(rows), columns=headers)


def find_col(df: pd.DataFrame, candidates: Sequence[str], required: bool = True) -> str | None:
    normalized = {str(c).strip().casefold(): c for c in df.columns}
    for candidate in candidates:
        key = candidate.strip().casefold()
        if key in normalized:
            return normalized[key]
    if required:
        raise KeyError(f"Missing required column. Tried: {', '.join(candidates)}")
    return None


def to_number(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
        .replace({"nan": np.nan, "None": np.nan, "": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def clean_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def is_top_style(value) -> bool:
    return str(value).strip().casefold() in TRUE_MARKERS


def material_group(value) -> str:
    """Classify textile composition for quick material filters.

    Cotton blends are intentionally classed as Bavlna. Technical therefore
    means a textile composition without cotton that contains common technical
    fibres such as polyester, elastane, nylon or polyamide.
    """
    text = str(value or "").casefold()
    if re.search(r"cotton|bavlna", text):
        return "Bavlna"
    if re.search(r"polyester|elastane|spandex|nylon|polyamid|polyamide|polypropylene", text):
        return "Technické"
    return "Ostatní / neurčeno"


def standard_size(size: str) -> str:
    """Normalize UA size labels used in master data to standard size-run keys."""
    value = str(size or "").strip().upper().replace(" ", "")
    mapping = {
        "XXS": "XXS",
        "XS": "XS",
        "SM": "S",
        "S": "S",
        "MD": "M",
        "M": "M",
        "LG": "L",
        "L": "L",
        "XL": "XL",
        "XXL": "2XL",
        "2XL": "2XL",
        "XXXL": "3XL",
        "3XL": "3XL",
        "4XL": "4XL",
        "5XL": "5XL",
    }
    return mapping.get(value, value)


def first_available_cols_central(central_df: pd.DataFrame) -> list[str]:
    week_cols = [
        c for c in central_df.columns
        if re.match(r"^\s*week\s*\d+", str(c), flags=re.I)
    ]
    if len(week_cols) < 2:
        raise KeyError(
            "Central stock file must contain at least two Week columns, "
            "e.g. Week 25 and Week 26."
        )
    return week_cols[:2]


def aggregate_stock(
    df: pd.DataFrame,
    ean_candidates: Sequence[str],
    qty_candidates: Sequence[str],
    output_col: str,
) -> pd.DataFrame:
    ean_col = find_col(df, ean_candidates)
    qty_col = find_col(df, qty_candidates)
    tmp = pd.DataFrame(
        {
            "EAN": df[ean_col].map(normalize_ean),
            output_col: to_number(df[qty_col]),
        }
    )
    tmp = tmp[tmp["EAN"] != ""]
    return tmp.groupby("EAN", as_index=False)[output_col].sum()


def aggregate_central(central_df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    ean_col = find_col(central_df, ["EAN", "EAN poslední"])
    week_cols = first_available_cols_central(central_df)

    tmp = pd.DataFrame({"EAN": central_df[ean_col].map(normalize_ean)})
    tmp["Central warehouse"] = sum(to_number(central_df[col]) for col in week_cols)
    tmp = tmp[tmp["EAN"] != ""]
    return tmp.groupby("EAN", as_index=False)["Central warehouse"].sum(), week_cols


def prepare_master(master_df: pd.DataFrame) -> pd.DataFrame:
    """Standardize master-data headers while accepting the current Czech export."""
    col_map = {
        "EAN": ["EAN"],
        "SKU": ["SKU"],
        "Artikl": ["Artikl", "Article", "ARTICLE_GENERIC"],
        "Size": ["Size US", "US Size", "Size", "Charakter.1", "Charakter"],
        "Název": ["Název", "Nazev", "Name", "ARTICLE_GENERIC_DESC"],
        "MOC CZK": ["MOC CZK", "MOC", "Prodejní cena"],
        "MOC EUR": ["MOC EUR", "EUR RRP", "RRP EUR"],
        "Division": ["Division"],
        "Gender": ["Gender"],
        "Silhouette": ["Silhouette"],
        "Fit": ["Fit"],
        "Segment": ["Segment"],
        "End use": ["End use", "End Use"],
        "Detail silhouette": ["Detail silhouette", "Detail Silhouette"],
        "Season": ["Season"],
        "C/O": ["C/O", "CO"],
        "Color group": ["Color group", "Colour group", "Color Group", "Colour Group"],
        "Color name": ["Color name", "Colour name", "Color Name", "Colour Name"],
        "Color code": ["Color code", "Colour code", "Color Code", "Colour Code"],
        "Composition": ["Composition", "Materiál", "Material"],
        "Top style raw": ["Top style", "Top styles", "Top Style", "TOP STYLE"],
    }

    out = pd.DataFrame(index=master_df.index)
    required = {"EAN", "Artikl", "Size", "Název"}
    for canonical, candidates in col_map.items():
        col = find_col(master_df, candidates, required=canonical in required)
        out[canonical] = "" if col is None else master_df[col]

    out["EAN"] = out["EAN"].map(normalize_ean)
    text_cols = [col for col in out.columns if col not in {"MOC CZK", "MOC EUR"}]
    for col in text_cols:
        out[col] = clean_text(out[col])

    out["MOC CZK"] = to_number(out["MOC CZK"])
    out["MOC EUR"] = to_number(out["MOC EUR"])
    out["Style code"] = out["Artikl"].str.split("-", n=1).str[0].str.strip()
    out["Materiál"] = out["Composition"].map(material_group)

    # A marker in any EAN row labels the entire base style as a top style.
    # This lets the master-data editor mark one row instead of every size and colour.
    out["_top_style_flag"] = out["Top style raw"].map(is_top_style)
    style_flag = out.groupby("Style code")["_top_style_flag"].transform("max")
    out["Top style"] = style_flag.fillna(False).astype(bool)

    out = out[out["EAN"] != ""].drop_duplicates(subset=["EAN"], keep="first")
    return out.drop(columns=["Top style raw", "_top_style_flag"])


def build_dataset(
    master_df: pd.DataFrame,
    local101_df: pd.DataFrame,
    local501_df: pd.DataFrame,
    central_df: pd.DataFrame,
) -> tuple[pd.DataFrame, list[str]]:
    master = prepare_master(master_df)
    qty_candidates = [
        "Qty", "Quantity", "Množství", "Mnozstvi", "Počet", "Pocet",
        "Počet ks", "Pocet ks", "Stock", "Available", "Available Qty",
    ]
    stock101 = aggregate_stock(
        local101_df, ["EAN poslední", "EAN"], qty_candidates, "Local warehouse 101"
    )
    stock501 = aggregate_stock(
        local501_df, ["EAN poslední", "EAN"], qty_candidates, "Local warehouse 501"
    )
    central, central_cols = aggregate_central(central_df)

    data = master.merge(stock101, on="EAN", how="left")
    data = data.merge(stock501, on="EAN", how="left")
    data = data.merge(central, on="EAN", how="left")

    warehouse_cols = ["Local warehouse 101", "Local warehouse 501", "Central warehouse"]
    for col in warehouse_cols:
        data[col] = data[col].fillna(0).round(0).astype(int)
    data["Total available"] = data[warehouse_cols].sum(axis=1)
    return data, central_cols


def product_type_mask(
    data: pd.DataFrame,
    product_types: Sequence[str],
    exclude_cotton: bool,
) -> pd.Series:
    if not product_types:
        return pd.Series(True, index=data.index)

    masks: list[pd.Series] = []
    division_apparel = data["Division"].str.casefold().eq("apparel")
    for product_type in product_types:
        if product_type == "Technical T-shirts":
            top_mask = data["Silhouette"].str.contains("tops", case=False, na=False)
            detail_mask = data["Detail silhouette"].str.contains(
                "sleeve|sleeveless|tee|t-shirt",
                case=False,
                na=False,
                regex=True,
            )
            mask = division_apparel & top_mask & detail_mask
            if exclude_cotton:
                mask &= ~data["Composition"].str.contains("cotton|bavlna", case=False, na=False, regex=True)
            masks.append(mask)
        elif product_type == "Shorts":
            detail_mask = data["Detail silhouette"].str.contains("shorts", case=False, na=False)
            silhouette_mask = data["Silhouette"].str.contains("shorts", case=False, na=False)
            masks.append(division_apparel & (detail_mask | silhouette_mask))
        elif product_type == "Tops":
            masks.append(division_apparel & data["Silhouette"].str.contains("tops", case=False, na=False))
        elif product_type == "Footwear":
            masks.append(data["Division"].str.contains("footwear", case=False, na=False))
        elif product_type == "Accessories":
            masks.append(data["Division"].str.contains("accessories", case=False, na=False))

    if not masks:
        return pd.Series(True, index=data.index)

    result = masks[0].copy()
    for current in masks[1:]:
        result |= current
    return result


def color_mask(data: pd.DataFrame, selected_colors: Sequence[str]) -> pd.Series:
    if not selected_colors:
        return pd.Series(True, index=data.index)

    color_group = clean_text(data["Color group"])
    color_name = clean_text(data["Color name"])
    masks: list[pd.Series] = []

    for color in selected_colors:
        if color == "Black":
            masks.append(
                color_group.str.contains("black", case=False, na=False)
                | color_name.str.contains("black", case=False, na=False)
            )
        elif color == "Dark Blue":
            masks.append(
                color_name.str.contains("navy|midnight|academy", case=False, na=False, regex=True)
            )
        else:
            masks.append(color_group.str.contains(re.escape(color), case=False, na=False, regex=True))

    result = masks[0].copy()
    for current in masks[1:]:
        result |= current
    return result


def size_sort_value(size: str) -> int:
    key = str(size).strip().upper()
    if key in SIZE_ORDER:
        return SIZE_ORDER[key]
    try:
        return int(float(key) * 10)
    except Exception:
        return 1000


def add_stock_metrics(data: pd.DataFrame, selected_warehouses: Sequence[str]) -> pd.DataFrame:
    """Calculate the style/colour aggregate and full-size-run status.

    Every metric is calculated before the final row filters. Therefore the
    displayed style/colour stock is not reduced merely because another size is
    hidden by the minimum-per-EAN setting.
    """
    out = data.copy()
    warehouse_cols = ["Local warehouse 101", "Local warehouse 501", "Central warehouse"]
    selected = [col for col in selected_warehouses if col in warehouse_cols]
    if not selected:
        selected = warehouse_cols

    out["Dostupnost ve vybraných skladech"] = out[selected].sum(axis=1).round(0).astype(int)
    out["_core_size"] = out["Size"].map(standard_size)
    out["Celkem ks styl/barva"] = (
        out.groupby("Artikl")["Dostupnost ve vybraných skladech"].transform("sum").round(0).astype(int)
    )

    available_rows = out[out["Dostupnost ve vybraných skladech"] > 0].copy()
    size_counts = available_rows.groupby("Artikl")["_core_size"].nunique()
    out["Dostupné velikosti styl/barva"] = (
        out["Artikl"].map(size_counts).fillna(0).astype(int)
    )

    full_flags: dict[str, bool | None] = {}
    for artikel, group in out.groupby("Artikl", sort=False):
        genders = clean_text(group["Gender"]).str.casefold().unique().tolist()
        gender = genders[0] if genders else ""
        required = CORE_SIZE_RUNS.get(gender)
        if not required:
            full_flags[artikel] = None
            continue
        available = set(
            group.loc[group["Dostupnost ve vybraných skladech"] > 0, "_core_size"].tolist()
        )
        full_flags[artikel] = required.issubset(available)

    out["_full_sizerun"] = out["Artikl"].map(full_flags)
    out["Plný sizerun"] = out["_full_sizerun"].map(
        lambda value: "Ano" if value is True else ("Ne" if value is False else "")
    )
    return out


def apply_filters(
    data: pd.DataFrame,
    product_types: Sequence[str] = ("Technical T-shirts",),
    genders: Sequence[str] = ("Mens",),
    colors: Sequence[str] = ("Black", "Dark Blue"),
    price_min: float = 0,
    price_max: float = 999,
    price_eur_min: float = 0,
    price_eur_max: float = 999,
    selected_warehouses: Sequence[str] = ("Local warehouse 101", "Local warehouse 501", "Central warehouse"),
    min_total_available: int = 1,
    min_style_color_qty: int = 1,
    min_article_sizes: int = 1,
    seasons: Sequence[str] = (),
    end_uses: Sequence[str] = (),
    detail_silhouettes: Sequence[str] = (),
    fits: Sequence[str] = (),
    co_values: Sequence[str] = (),
    exclude_cotton: bool = True,
    technical_material: bool = False,
    cotton_material: bool = False,
    only_top_styles: bool = False,
    only_full_sizerun: bool = False,
) -> pd.DataFrame:
    data = add_stock_metrics(data, selected_warehouses)
    mask = pd.Series(True, index=data.index)

    # Cotton checkbox intentionally overrides the legacy "Exclude cotton" setting.
    effective_exclude_cotton = exclude_cotton and not cotton_material
    mask &= product_type_mask(data, product_types, effective_exclude_cotton)

    if genders:
        mask &= data["Gender"].isin(genders)
    if colors:
        mask &= color_mask(data, colors)

    mask &= data["MOC CZK"].between(price_min, price_max, inclusive="both")
    mask &= data["MOC EUR"].between(price_eur_min, price_eur_max, inclusive="both")

    if seasons:
        mask &= data["Season"].isin(seasons)
    if end_uses:
        mask &= data["End use"].isin(end_uses)
    if detail_silhouettes:
        mask &= data["Detail silhouette"].isin(detail_silhouettes)
    if fits:
        mask &= data["Fit"].isin(fits)
    if co_values:
        mask &= data["C/O"].isin(co_values)

    if technical_material or cotton_material:
        material_options = set()
        if technical_material:
            material_options.add("Technické")
        if cotton_material:
            material_options.add("Bavlna")
        mask &= data["Materiál"].isin(material_options)

    if only_top_styles:
        mask &= data["Top style"]

    if only_full_sizerun:
        mask &= data["_full_sizerun"].eq(True)

    mask &= data["Dostupnost ve vybraných skladech"] >= int(min_total_available)
    mask &= data["Celkem ks styl/barva"] >= int(min_style_color_qty)
    mask &= data["Dostupné velikosti styl/barva"] >= int(min_article_sizes)

    filtered = data[mask].copy()
    if filtered.empty:
        return filtered

    filtered["_size_sort"] = filtered["Size"].map(size_sort_value)
    filtered = (
        filtered.sort_values(["Artikl", "_size_sort", "Size", "EAN"])
        .drop(columns=["_size_sort", "_core_size", "_full_sizerun"], errors="ignore")
    )
    return filtered


def display_qty(value: int | float) -> int | str:
    try:
        number = int(round(float(value)))
    except Exception:
        return ""
    return "100+" if number > 100 else number


def to_offer_table(filtered: pd.DataFrame, include_extra_columns: bool = False) -> pd.DataFrame:
    offer = pd.DataFrame(
        {
            "Artikl": filtered["Artikl"],
            "Size": filtered["Size"],
            "Název": filtered["Název"],
            "Local warehouse 101": filtered["Local warehouse 101"].map(display_qty),
            "Local warehouse 501": filtered["Local warehouse 501"].map(display_qty),
            "Central warehouse": filtered["Central warehouse"].map(display_qty),
            "Celkem ks styl/barva": filtered["Celkem ks styl/barva"],
            "Plný sizerun": filtered["Plný sizerun"],
            "Top style": filtered["Top style"].map(lambda value: "Ano" if bool(value) else ""),
            "Materiál": filtered["Materiál"],
            "ORDER": "",
            "MOC CZK": filtered["MOC CZK"],
            "MOC EUR": filtered["MOC EUR"],
            "EAN": filtered["EAN"].astype(str),
            "Gender": filtered["Gender"],
            "Silhouette": filtered["Silhouette"],
            "Fit": filtered["Fit"],
            "End use": filtered["End use"],
            "Season": filtered["Season"],
            "C/O": filtered["C/O"],
        }
    )

    if include_extra_columns:
        for col in EXTRA_COLUMNS:
            if col in filtered.columns:
                offer[col] = filtered[col]
    return offer


def write_offer_excel(offer_df: pd.DataFrame, title: str = "Under Armour Product Offer") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Offer"

    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)
    order_fill = PatternFill("solid", fgColor="FFF2CC")
    title_font = Font(size=14, bold=True, color="111827")
    thin_side = Side(style="thin", color="D1D5DB")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws.cell(row=1, column=1, value=title).font = title_font
    ws.cell(row=2, column=1, value=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    start_row = 4
    for col_idx, column_name in enumerate(offer_df.columns, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=column_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    number_columns = {
        "MOC CZK",
        "MOC EUR",
        "Celkem ks styl/barva",
        "Dostupnost ve vybraných skladech",
        "Total available",
        "Dostupné velikosti styl/barva",
    }
    for row_idx, row in enumerate(offer_df.itertuples(index=False), start=start_row + 1):
        for col_idx, value in enumerate(row, start=1):
            header = offer_df.columns[col_idx - 1]
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if header == "ORDER":
                cell.fill = order_fill
            if header == "EAN":
                cell.number_format = "@"
            if header in number_columns:
                cell.number_format = "#,##0"

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = (
        f"A{start_row}:{get_column_letter(len(offer_df.columns))}{start_row + len(offer_df)}"
    )

    widths = {
        "Artikl": 16,
        "Size": 10,
        "Název": 34,
        "Local warehouse 101": 18,
        "Local warehouse 501": 18,
        "Central warehouse": 18,
        "Celkem ks styl/barva": 21,
        "Dostupnost ve vybraných skladech": 24,
        "Dostupné velikosti styl/barva": 25,
        "Plný sizerun": 14,
        "Top style": 12,
        "Materiál": 16,
        "ORDER": 12,
        "MOC CZK": 12,
        "MOC EUR": 12,
        "EAN": 18,
        "Gender": 12,
        "Silhouette": 14,
        "Fit": 16,
        "End use": 16,
        "Season": 12,
        "C/O": 10,
        "Color group": 14,
        "Color name": 22,
        "Color code": 12,
        "Detail silhouette": 18,
        "Composition": 34,
        "Total available": 15,
        "Style code": 14,
    }
    for idx, column_name in enumerate(offer_df.columns, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(column_name, 14)

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[start_row].height = 30

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def unique_sorted(data: pd.DataFrame, col: str) -> list[str]:
    if col not in data.columns:
        return []
    values = [
        value
        for value in clean_text(data[col]).unique().tolist()
        if value and value.casefold() not in {"0", "nan", "none"}
    ]
    return sorted(values)
