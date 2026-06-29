"""UA Offer Generator - Streamlit app."""

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
    st.caption("Generate Under Armour offers from master data and warehouse files.")

    with st.sidebar:
        st.header("1. Upload files")
        master_file = st.file_uploader(
            "Master data",
            type=["xlsx"],
            help=(
                "Must include EAN, Artikl, size, name, MOC CZK, MOC EUR and Composition. "
                "For the Top style filter, use a Top style / Top styles column and mark any row of the base style with x."
            ),
        )
        local101_file = st.file_uploader("Local warehouse 101", type=["xlsx"])
        local501_file = st.file_uploader("Local warehouse 501", type=["xlsx"])
        central_file = st.file_uploader(
            "Central warehouse",
            type=["xlsx"],
            help="The app sums the first two Week columns.",
        )

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

    st.success(
        f"Loaded {len(data):,} EAN rows from master data. "
        f"Central availability = {' + '.join(central_cols)}."
    )

    st.header("2. Offer criteria")
    col1, col2, col3, col4 = st.columns(4)

    gender_options = unique_sorted(data, "Gender")
    detail_silhouette_options = unique_sorted(data, "Detail silhouette")
    fit_options = unique_sorted(data, "Fit")
    moc_eur_max_default = int(max(float(data["MOC EUR"].max()), 999))
    moc_czk_max_default = int(max(float(data["MOC CZK"].max()), 999))

    with col1:
        product_types = st.multiselect(
            "Product type",
            ["Technical T-shirts", "Shorts", "Tops", "Footwear", "Accessories"],
            default=["Technical T-shirts"],
        )
        genders = st.multiselect(
            "Gender",
            gender_options,
            default=["Mens"] if "Mens" in gender_options else [],
        )

        st.markdown("**Material**")
        technical_material = st.checkbox("Technické", value=False)
        cotton_material = st.checkbox("Bavlna", value=False)
        only_top_styles = st.checkbox("Only Top styles", value=False)

    with col2:
        colors = st.multiselect(
            "Color group / color logic",
            ["Black", "Dark Blue", "Blue", "Gray", "White", "Red", "Green", "Orange", "Yellow", "Brown"],
            default=["Black", "Dark Blue"],
        )
        price_min = st.number_input("MOC CZK from", min_value=0, value=0, step=100)
        price_max = st.number_input(
            "MOC CZK to",
            min_value=0,
            value=min(999, moc_czk_max_default),
            step=100,
        )
        price_eur_min = st.number_input("MOC EUR from", min_value=0, value=0, step=5)
        price_eur_max = st.number_input(
            "MOC EUR to",
            min_value=0,
            value=moc_eur_max_default,
            step=5,
        )

    with col3:
        selected_warehouses = st.multiselect(
            "Availability source used for filtering",
            ["Local warehouse 101", "Local warehouse 501", "Central warehouse"],
            default=["Local warehouse 101", "Local warehouse 501", "Central warehouse"],
        )
        min_total_available = st.number_input(
            "Minimum selected availability per EAN",
            min_value=0,
            value=1,
            step=1,
        )
        min_style_color_qty = st.number_input(
            "Minimum total pieces per style / colour",
            min_value=0,
            value=1,
            step=1,
            help=(
                "Sum of all EAN rows with the same Artikl (style + colour), "
                "calculated from the selected warehouses."
            ),
        )
        min_article_sizes = st.number_input(
            "Minimum available sizes per style / colour",
            min_value=1,
            value=1,
            step=1,
        )
        only_full_sizerun = st.checkbox(
            "Only full sizerun",
            value=False,
            help="Mens: S–2XL. Womens: XS–XL. Every core size must have stock in the selected warehouses.",
        )

    with col4:
        seasons = st.multiselect("Season", unique_sorted(data, "Season"), default=[])
        end_uses = st.multiselect("End use", unique_sorted(data, "End use"), default=[])
        detail_silhouettes = st.multiselect(
            "Sleeve / detail silhouette",
            detail_silhouette_options,
            default=[],
        )
        fits = st.multiselect("Fit", fit_options, default=[])
        co_values = st.multiselect("C/O", unique_sorted(data, "C/O"), default=[])
        include_extra_columns = st.checkbox("Include helper columns in export", value=False)

    filtered = apply_filters(
        data=data,
        product_types=product_types,
        genders=genders,
        colors=colors,
        price_min=float(price_min),
        price_max=float(price_max),
        price_eur_min=float(price_eur_min),
        price_eur_max=float(price_eur_max),
        selected_warehouses=selected_warehouses,
        min_total_available=int(min_total_available),
        min_style_color_qty=int(min_style_color_qty),
        min_article_sizes=int(min_article_sizes),
        seasons=seasons,
        end_uses=end_uses,
        detail_silhouettes=detail_silhouettes,
        fits=fits,
        co_values=co_values,
        technical_material=technical_material,
        cotton_material=cotton_material,
        only_top_styles=only_top_styles,
        only_full_sizerun=only_full_sizerun,
    )

    offer_df = to_offer_table(filtered, include_extra_columns=include_extra_columns)

    st.header("3. Preview")
    full_sizerun_count = (
        filtered.loc[filtered["Plný sizerun"].eq("Ano"), "Artikl"].nunique()
        if not filtered.empty
        else 0
    )
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("EAN rows", f"{len(offer_df):,}")
    kpi2.metric("Style / colour", f"{offer_df['Artikl'].nunique() if not offer_df.empty else 0:,}")
    kpi3.metric(
        "Selected availability",
        f"{int(filtered['Dostupnost ve vybraných skladech'].sum()) if not filtered.empty else 0:,}",
    )
    kpi4.metric("Full sizeruns", f"{full_sizerun_count:,}")
    kpi5.metric(
        "Average MOC CZK",
        f"{filtered['MOC CZK'].mean():,.0f}" if not filtered.empty else "-",
    )
    kpi6.metric(
        "Average MOC EUR",
        f"{filtered['MOC EUR'].mean():,.0f}" if not filtered.empty else "-",
    )

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
        excel_bytes = write_offer_excel(offer_df)
        st.download_button(
            label="Download Excel offer",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.expander("Current filter logic"):
        st.markdown(
            """
            **Top styles**: the master column can be named `Top style` or `Top styles`.
            Values `x`, `ano`, `yes`, `true` or `1` are accepted. A marker in any
            EAN row makes the entire base style a Top style across all colours and sizes.

            **Technické / Bavlna**: a row containing `cotton` or `bavlna` in
            `Composition` is Bavlna. Other compositions containing common performance
            fibres such as polyester, elastane, nylon or polyamide are Technické.
            Both selected includes both groups; neither selected leaves material unrestricted.
            The product type **Technical T-shirts** now defines only the product category,
            not the material.

            **Celkem ks styl/barva**: total stock across all EAN rows sharing `Artikl`
            (UA style + colour), calculated from the currently selected warehouses.
            It is shown in every export.

            **Full sizerun**: Mens requires S, M, L, XL and 2XL; Womens requires
            XS, S, M, L and XL. Each of those sizes must have stock in the selected
            warehouses. UA size labels SM / MD / LG and XXL are converted automatically.

            **Central warehouse** = sum of the first two `Week` columns in the central
            warehouse file.
            """
        )


if __name__ == "__main__":
    main()
