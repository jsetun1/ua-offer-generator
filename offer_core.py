"""Core logic for UA Offer Generator."""

from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Iterable, Sequence

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_COLUMNS = [
    "Artikl", "Size", "Název", "Local warehouse 101", "Local warehouse 501",
    "Central warehouse", "ORDER", "EUR RRP", "EAN", "Gender", "Silhouette",
    "Fit", "End use", "Season", "C/O",
]

EXTRA_COLUMNS = ["MOC CZK", "Color group", "Color name", "Color code", "Detail silhouette", "Composition", "Total available"]

SIZE_ORDER = {
    "XXS": 10, "XS": 20, "SM": 30, "S": 30, "MD": 40, "M": 40,
    "LG": 50, "L": 50, "XL": 60, "XXL": 70, "2XL": 70, "XXXL": 80,
    "3XL": 80, "4XL": 90, "5XL": 100, "OSFA": 900, "OSFM": 900,
}


def normalize_ean(value) -> str:
    """Return a safe text EAN without decimals/spaces."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 8:
        return digits
    return text


def read_excel(path_or_buffer) -> pd.DataFrame:
    """Read the first worksheet into a DataFrame.

    This uses openpyxl read-only mode because some stock exports are slow through
    pandas.read_excel. Numeric conversion happens later.
    """
    from openpyxl import load_workbook

    wb = load_workbook(path_or_buffer, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    try:
        header_row = next(rows)
    except StopIteration:
        return pd.DataFrame()
    headers = []
    seen = {}
    for idx, value in enumerate(header_row, start=1):
        name = str(value).strip() if value is not None else f"Column {idx}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        headers.append(name)
    data = list(rows)
    df = pd.DataFrame(data, columns=headers)
    return df


def find_col(df: pd.DataFrame, candidates: Sequence[str], required: bool = True) -> str | None:
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in normalized:
            return normalized[key]
    if required:
        raise KeyError(f"Missing required column. Tried: {', '.join(candidates)}")
    return None


def to_number(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
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


def first_available_cols_central(central_df: pd.DataFrame) -> list[str]:
    week_cols = [c for c in central_df.columns if re.match(r"^\s*week\s*\d+", str(c), flags=re.I)]
    if len(week_cols) < 2:
        raise KeyError("Central stock file must contain at least two Week columns, e.g. Week 25 and Week 26.")
    return week_cols[:2]


def aggregate_stock(df: pd.DataFrame, ean_candidates: Sequence[str], qty_candidates: Sequence[str], output_col: str) -> pd.DataFrame:
    ean_col = find_col(df, ean_candidates)
    qty_col = find_col(df, qty_candidates)
    tmp = pd.DataFrame({"EAN": df[ean_col].map(normalize_ean), output_col: to_number(df[qty_col])})
    tmp = tmp[tmp["EAN"] != ""]
    return tmp.groupby("EAN", as_index=False)[output_col].sum()


def aggregate_central(central_df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    ean_col = find_col(central_df, ["EAN"])
    week_cols = first_available_cols_central(central_df)
    tmp = pd.DataFrame({"EAN": central_df[ean_col].map(normalize_ean)})
    tmp["Central warehouse"] = sum(to_number(central_df[c]) for c in week_cols)
    tmp = tmp[tmp["EAN"] != ""]
    return tmp.groupby("EAN", as_index=False)["Central warehouse"].sum(), week_cols


def prepare_master(master_df: pd.DataFrame) -> pd.DataFrame:
    col_map = {
        "EAN": ["EAN"],
        "SKU": ["SKU"],
        "Artikl": ["Artikl", "Article", "ARTICLE_GENERIC"],
        "Size": ["Size US", "US Size", "Size"],
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
    }
    out = pd.DataFrame()
    for canonical, candidates in col_map.items():
        col = find_col(master_df, candidates, required=canonical in ["EAN", "Artikl", "Size", "Název"])
        out[canonical] = "" if col is None else master_df[col]

    out["EAN"] = out["EAN"].map(normalize_ean)
    text_cols = [c for c in out.columns if c not in {"MOC CZK", "MOC EUR"}]
    for col in text_cols:
        out[col] = clean_text(out[col])
    out["MOC CZK"] = to_number(out["MOC CZK"])
    out["MOC EUR"] = to_number(out["MOC EUR"])
    out = out[out["EAN"] != ""].drop_duplicates(subset=["EAN"], keep="first")
    return out


def build_dataset(master_df: pd.DataFrame, local101_df: pd.DataFrame, local501_df: pd.DataFrame, central_df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    master = prepare_master(master_df)
    stock101 = aggregate_stock(local101_df, ["EAN poslední", "EAN"], ["Qty", "Quantity"], "Local warehouse 101")
    stock501 = aggregate_stock(local501_df, ["EAN poslední", "EAN"], ["Qty", "Quantity"], "Local warehouse 501")
    central, central_cols = aggregate_central(central_df)
    data = master.merge(stock101, on="EAN", how="left")
    data = data.merge(stock501, on="EAN", how="left")
    data = data.merge(central, on="EAN", how="left")
    for col in ["Local warehouse 101", "Local warehouse 501", "Central warehouse"]:
        data[col] = data[col].fillna(0).round(0).astype(int)
    data["Total available"] = data[["Local warehouse 101", "Local warehouse 501", "Central warehouse"]].sum(axis=1)
    return data, central_cols


def product_type_mask(data: pd.DataFrame, product_types: Sequence[str], exclude_cotton: bool) -> pd.Series:
    if not product_types:
        return pd.Series(True, index=data.index)
    masks = []
    division_apparel = data["Division"].str.casefold().eq("apparel")
    for product_type in product_types:
        if product_type == "Technical T-shirts":
            top_mask = data["Silhouette"].str.contains("tops", case=False, na=False)
            detail_mask = data["Detail silhouette"].str.contains("sleeve|sleeveless|tee|t-shirt", case=False, na=False, regex=True)
            mask = division_apparel & top_mask & detail_mask
            if exclude_cotton:
                mask = mask & ~data["Composition"].str.contains("cotton", case=False, na=False)
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
    for mask in masks[1:]:
        result = result | mask
    return result


def color_mask(data: pd.DataFrame, selected_colors: Sequence[str]) -> pd.Series:
    if not selected_colors:
        return pd.Series(True, index=data.index)
    color_group = clean_text(data["Color group"])
    color_name = clean_text(data["Color name"])
    masks = []
    for color in selected_colors:
        if color == "Black":
            masks.append(color_group.str.contains("black", case=False, na=False) | color_name.str.contains("black", case=False, na=False))
        elif color == "Dark Blue":
            masks.append(color_name.str.contains("navy|midnight|academy", case=False, na=False, regex=True))
        else:
            masks.append(color_group.str.contains(re.escape(color), case=False, na=False, regex=True))
    result = masks[0].copy()
    for mask in masks[1:]:
        result = result | mask
    return result


def size_sort_value(size: str) -> int:
    s = str(size).strip().upper()
    if s in SIZE_ORDER:
        return SIZE_ORDER[s]
    try:
        return int(float(s) * 10)
    except Exception:
        return 1000


def apply_filters(
    data: pd.DataFrame,
    product_types: Sequence[str] = ("Technical T-shirts",),
    genders: Sequence[str] = ("Mens",),
    colors: Sequence[str] = ("Black", "Dark Blue"),
    price_min: float = 0,
    price_max: float = 999,
    selected_warehouses: Sequence[str] = ("Local warehouse 101", "Local warehouse 501", "Central warehouse"),
    min_total_available: int = 1,
    seasons: Sequence[str] = (),
    end_uses: Sequence[str] = (),
    detail_silhouettes: Sequence[str] = (),
    fits: Sequence[str] = (),
    co_values: Sequence[str] = (),
    exclude_cotton: bool = True,
    min_article_qty: int = 1,
    min_article_sizes: int = 1,
) -> pd.DataFrame:
    mask = pd.Series(True, index=data.index)
    mask &= product_type_mask(data, product_types, exclude_cotton)
    if genders:
        mask &= data["Gender"].isin(genders)
    if colors:
        mask &= color_mask(data, colors)
    mask &= data["MOC CZK"].between(price_min, price_max, inclusive="both")
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
    availability_cols = list(selected_warehouses) if selected_warehouses else ["Local warehouse 101", "Local warehouse 501", "Central warehouse"]
    selected_available = data[availability_cols].sum(axis=1)
    mask &= selected_available >= min_total_available
    filtered = data[mask].copy()
    if filtered.empty:
        return filtered
    article_stats = filtered.groupby("Artikl").agg(article_qty=("Total available", "sum"), article_sizes=("Size", "nunique"))
    valid_articles = article_stats[(article_stats["article_qty"] >= min_article_qty) & (article_stats["article_sizes"] >= min_article_sizes)].index
    filtered = filtered[filtered["Artikl"].isin(valid_articles)].copy()
    filtered["_size_sort"] = filtered["Size"].map(size_sort_value)
    filtered = filtered.sort_values(["Artikl", "_size_sort", "Size", "EAN"]).drop(columns=["_size_sort"])
    return filtered


def display_qty(value: int | float) -> int | str:
    try:
        number = int(round(float(value)))
    except Exception:
        return ""
    return "100+" if number > 100 else number


def to_offer_table(filtered: pd.DataFrame, include_extra_columns: bool = False) -> pd.DataFrame:
    offer = pd.DataFrame({
        "Artikl": filtered["Artikl"],
        "Size": filtered["Size"],
        "Název": filtered["Název"],
        "Local warehouse 101": filtered["Local warehouse 101"].map(display_qty),
        "Local warehouse 501": filtered["Local warehouse 501"].map(display_qty),
        "Central warehouse": filtered["Central warehouse"].map(display_qty),
        "ORDER": "",
        "EUR RRP": filtered["MOC EUR"],
        "EAN": filtered["EAN"].astype(str),
        "Gender": filtered["Gender"],
        "Silhouette": filtered["Silhouette"],
        "Fit": filtered["Fit"],
        "End use": filtered["End use"],
        "Season": filtered["Season"],
        "C/O": filtered["C/O"],
    })
    if include_extra_columns:
        for col in EXTRA_COLUMNS:
            if col in filtered.columns:
                offer[col] = filtered[col]
    return offer


def write_offer_excel(offer_df: pd.DataFrame, title: str = "UA Offer") -> bytes:
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
            if header in {"EUR RRP", "MOC CZK", "Total available"}:
                cell.number_format = "#,##0"
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(offer_df.columns))}{start_row + len(offer_df)}"
    widths = {
        "Artikl": 16, "Size": 10, "Název": 34, "Local warehouse 101": 18,
        "Local warehouse 501": 18, "Central warehouse": 18, "ORDER": 12,
        "EUR RRP": 11, "EAN": 18, "Gender": 12, "Silhouette": 14,
        "Fit": 16, "End use": 16, "Season": 12, "C/O": 10,
        "MOC CZK": 12, "Color group": 14, "Color name": 22,
        "Color code": 12, "Detail silhouette": 18, "Composition": 34,
        "Total available": 15,
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
    vals = [v for v in clean_text(data[col]).unique().tolist() if v and v.lower() not in {"0", "nan", "none"}]
    return sorted(vals)
