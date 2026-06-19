"""UA Offer Generator - Streamlit prototype."""

from __future__ import annotations

import streamlit as st

from offer_core import (
    apply_filters,
    build_dataset,
    read_excel,
    to_offer_table,
    unique_sorted,
    write_offer_excel,
)

APP_TITLE = "UA Offer Generator"


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Prototype for generating Under Armour product offers from master data and warehouse files.")

    with st.sidebar:
        st.header("1. Upload files")
        master_file = st.file_uploader("Master data", type=["xlsx"], help="Must include EAN, Artikl, MOC CZK, MOC EUR, Color group/name and Composition.")
        local101_file = st.file_uploader("Local warehouse 101", type=["xlsx"])
        local501_file = st.file_uploader("Local warehouse 501", type=["xlsx"])
        central_file = st.file_uploader("Central warehouse", type=["xlsx"], help="The app sums the first two Week columns.")

    if not all([master_file, local101_file, local501_file, central_file]):
        st.info("Upload all four files to start.")
        st.stop()

    try:
        with st.spinner("Loading and merging data by EAN..."):
            master_df = read_excel(master_file)
            local101_df = read_excel(local101_file)
            local501_df = read_excel(local501_file)
            central_df = read_excel(central_file)
            data, central_cols = build_dataset(master_df, local101_df, local501_df, central_df)
    except Exception as exc:
        st.error(f"File loading failed: {exc}")
        st.stop()

    st.success(f"Loaded {len(data):,} EAN rows from master data. Central availability = {' + '.join(central_cols)}.")

    st.header("2. Offer criteria")
    col1, col2, col3, col4 = st.columns(4)
    gender_options = unique_sorted(data, "Gender")

    with col1:
        product_types = st.multiselect(
            "Product type",
            ["Technical T-shirts", "Shorts", "Tops", "Footwear", "Accessories"],
            default=["Technical T-shirts"],
        )
        genders = st.multiselect("Gender", gender_options, default=["Mens"] if "Mens" in gender_options else [])
        exclude_cotton = st.checkbox("Exclude cotton", value=True)
    with col2:
        colors = st.multiselect(
            "Color group / color logic",
            ["Black", "Dark Blue", "Blue", "Gray", "White", "Red", "Green", "Orange", "Yellow", "Brown"],
            default=["Black", "Dark Blue"],
        )
        price_min = st.number_input("MOC CZK from", min_value=0, value=0, step=100)
        price_max = st.number_input("MOC CZK to", min_value=0, value=999, step=100)
    with col3:
        selected_warehouses = st.multiselect(
            "Availability source used for filtering",
            ["Local warehouse 101", "Local warehouse 501", "Central warehouse"],
            default=["Local warehouse 101", "Local warehouse 501", "Central warehouse"],
        )
        min_total_available = st.number_input("Minimum selected availability per EAN", min_value=0, value=1, step=1)
        min_article_qty = st.number_input("Minimum total availability per article", min_value=0, value=1, step=1)
        min_article_sizes = st.number_input("Minimum available sizes per article", min_value=1, value=1, step=1)
    with col4:
        seasons = st.multiselect("Season", unique_sorted(data, "Season"), default=[])
        end_uses = st.multiselect("End use", unique_sorted(data, "End use"), default=[])
        co_values = st.multiselect("C/O", unique_sorted(data, "C/O"), default=[])
        include_extra_columns = st.checkbox("Include helper columns in export", value=False)

    filtered = apply_filters(
        data=data,
        product_types=product_types,
        genders=genders,
        colors=colors,
        price_min=float(price_min),
        price_max=float(price_max),
        selected_warehouses=selected_warehouses,
        min_total_available=int(min_total_available),
        seasons=seasons,
        end_uses=end_uses,
        co_values=co_values,
        exclude_cotton=exclude_cotton,
        min_article_qty=int(min_article_qty),
        min_article_sizes=int(min_article_sizes),
    )

    offer_df = to_offer_table(filtered, include_extra_columns=include_extra_columns)

    st.header("3. Preview")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("EAN rows", f"{len(offer_df):,}")
    kpi2.metric("Unique articles", f"{offer_df['Artikl'].nunique() if not offer_df.empty else 0:,}")
    kpi3.metric("Total available", f"{int(filtered['Total available'].sum()) if not filtered.empty else 0:,}")
    kpi4.metric("Average MOC CZK", f"{filtered['MOC CZK'].mean():,.0f}" if not filtered.empty else "-")

    st.dataframe(offer_df.head(500), use_container_width=True, hide_index=True)
    if len(offer_df) > 500:
        st.caption("Preview shows first 500 rows only. Export contains all filtered rows.")

    st.header("4. Export")
    file_name = st.text_input("Output filename", value="ua_offer.xlsx")
    if not file_name.lower().endswith(".xlsx"):
        file_name += ".xlsx"

    if offer_df.empty:
        st.warning("No products match the current filters.")
    else:
        excel_bytes = write_offer_excel(offer_df, title="Under Armour Product Offer")
        st.download_button(
            label="Download Excel offer",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.expander("Current filter logic"):
        st.markdown(
            """
            **Technical T-shirts** = `Division = Apparel` + `Silhouette = Tops` + `Detail silhouette contains sleeve/sleeveless/tee`.

            **Exclude cotton** removes rows where `Composition` contains `cotton`.

            **Dark Blue** currently means `Color name` contains `Navy`, `Midnight`, or `Academy`.

            **Central warehouse** = sum of the first two `Week` columns in the central warehouse file.
            """
        )


if __name__ == "__main__":
    main()

