"""UA Offer Generator - Streamlit app."""

from __future__ import annotations

import streamlit as st

from offer_core import (
    ALTERNATIVE_ENGINE_VERSION,
    apply_filters,
    build_dataset,
    read_excel,
    read_style_selectors_excel,
    imported_unavailable_with_alternatives,
    to_offer_table,
    unique_sorted,
    unmatched_style_selectors,
    write_import_substitution_excel,
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
        style_import_file = st.file_uploader(
            "Import stylů / stylů-barvy (XLSX)",
            type=["xlsx"],
            help=(
                "V prvním listu vložte styly do buněk pod sebe. Hodnota 1326799 vybere všechny "
                "barvy stylu; 1326799-036 (případně 1326799/036) vybere pouze konkrétní styl/barvu. "
                "Při importu mají zadané styly přednost před všemi standardními filtry; vybraný sklad "
                "slouží pouze pro výpočet a zobrazení dostupnosti."
            ),
            key="style_import_file",
        )
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

    selected_style_refs: list[str] = []
    if style_import_file is not None:
        try:
            selected_style_refs = read_style_selectors_excel(style_import_file)
        except Exception as exc:
            st.error(f"Style import failed: {exc}")
            st.stop()

        if not selected_style_refs:
            st.warning(
                "No valid styles found in the imported file. Use values such as 1326799 or 1326799-036."
            )
            st.stop()
        else:
            unmatched = unmatched_style_selectors(data, selected_style_refs)
            alternative_strategy_label = st.radio(
                "Strategie doporučených alternativ",
                options=[
                    "Obě varianty",
                    "Stejný styl – jiná barva",
                    "Stejná barva – jiný styl",
                ],
                index=0,
                help=(
                    "Stejný styl zachová střih a obvykle i cenu, ale nabídne jinou dostupnou barvu. "
                    "Stejná barva hledá jiný dostupný styl, například jiné pánské šedé golfové polo. "
                    "Volba Obě varianty zobrazí vyvážený mix obou cest."
                ),
            )
            alternative_strategy = {
                "Obě varianty": "both",
                "Stejný styl – jiná barva": "same_style",
                "Stejná barva – jiný styl": "same_color_new_style",
            }[alternative_strategy_label]
            alternative_count = st.selectbox(
                "Počet automaticky navržených alternativ",
                options=[3, 5, 8],
                index=1,
                help=(
                    "Alternativy se hledají pouze mezi produkty dostupnými ve skladech, "
                    "které jsou zvolené v poli Availability source used for filtering. U volby "
                    "Obě varianty se počet rozdělí mezi obě cesty; nevyužitá místa doplní druhá cesta."
                ),
            )
            unavailable_import_df, alternatives_df = imported_unavailable_with_alternatives(
                data=data,
                style_selectors=selected_style_refs,
                selected_warehouses=selected_warehouses,
                max_alternatives=int(alternative_count),
                alternative_strategy=alternative_strategy,
            )
            st.success(
                f"Import mode is active: {len(selected_style_refs)} unique selection(s) loaded. "
                "The import overrides every standard offer filter. Base styles include all colours; "
                "exact style-colour values include only that colour. Selected warehouses are used only "
                "to calculate and display availability, not to remove imported products."
            )

            if not unavailable_import_df.empty:
                st.warning(
                    f"{len(unavailable_import_df)} imported style / colour item(s) exist in master data "
                    "but have zero stock across all uploaded warehouses."
                )
                with st.expander("Nedostupné položky z importu a doporučené alternativy", expanded=False):
                    st.caption(
                        f"Alternativy v{ALTERNATIVE_ENGINE_VERSION.replace('v', '')}: přesně nedostupný artikl "
                        "je vždy vyloučen. Podle zvolené strategie aplikace buď nabídne stejný styl v jiné "
                        "dostupné barvě, nebo jiný dostupný styl ve stejné barvě. U jiného stylu nejprve drží "
                        "typ produktu (např. Polo), pohlaví, division, použití a střih; poté řadí fit, materiál, "
                        "segment, MOC a zásobu. Dostupnost alternativ odpovídá aktuálně vybraným skladům."
                    )
                    st.subheader("Nedostupné z importu")
                    st.dataframe(unavailable_import_df, use_container_width=True, hide_index=True)
                    st.subheader("Doporučené alternativy")
                    if alternatives_df.empty:
                        st.info("Pro tyto položky nebyla ve vybraných skladech nalezena vhodná dostupná alternativa.")
                    else:
                        st.dataframe(alternatives_df, use_container_width=True, hide_index=True)
            if unmatched:
                preview_unmatched = ", ".join(unmatched[:12])
                suffix = " …" if len(unmatched) > 12 else ""
                st.warning(f"Not found in master data: {preview_unmatched}{suffix}")

            if not unavailable_import_df.empty or unmatched:
                substitution_bytes = write_import_substitution_excel(
                    unavailable_import_df,
                    alternatives_df,
                    unmatched,
                )
                st.download_button(
                    label="Download nedostupných produktů a alternativ",
                    data=substitution_bytes,
                    file_name="ua_import_nedostupne_a_alternativy.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help=(
                        "Soubor obsahuje list Nedostupné z importu, list Alternativy a případně list Mimo master."
                    ),
                )

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
        selected_style_refs=selected_style_refs,
    )

    # Preview remains neutral: it shows both MOC currencies and does not yet
    # add a negotiated discount. Final currency and discount are chosen below.
    preview_df = to_offer_table(
        filtered,
        include_extra_columns=include_extra_columns,
        export_currency="CZK + EUR",
        discount_percent=None,
    )

    st.header("3. Preview")
    full_sizerun_count = (
        filtered.loc[filtered["Plný sizerun"].eq("Ano"), "Artikl"].nunique()
        if not filtered.empty
        else 0
    )
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("EAN rows", f"{len(preview_df):,}")
    kpi2.metric("Style / colour", f"{preview_df['Artikl'].nunique() if not preview_df.empty else 0:,}")
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

    st.dataframe(preview_df.head(500), use_container_width=True, hide_index=True)
    if len(preview_df) > 500:
        st.caption("Preview shows first 500 rows only. Export contains all filtered rows.")

    st.header("4. Final pricing and export")
    export_col1, export_col2, export_col3 = st.columns([1, 1, 1])
    with export_col1:
        discount_percent = st.number_input(
            "Sleva z MOC včetně DPH (%)",
            min_value=0.0,
            max_value=100.0,
            value=0.0,
            step=0.5,
            format="%.1f",
            help=(
                "Sleva se propíše až do finální nabídky a nemění seznam vyfiltrovaných produktů."
            ),
        )
    with export_col2:
        export_currency = st.radio(
            "Měna ve finálním exportu",
            options=["CZK", "EUR", "CZK + EUR"],
            index=2,
            horizontal=True,
            help="Vyberte, zda má export obsahovat ceny v CZK, EUR, nebo obou měnách.",
        )
    with export_col3:
        price_vat_mode = st.radio(
            "Cena po slevě v exportu",
            options=["Bez DPH", "S DPH"],
            index=0,
            horizontal=True,
            help="Volí, zda se výsledná nabídková cena po slevě zobrazí bez DPH nebo včetně DPH.",
        )

    if price_vat_mode == "S DPH":
        st.caption(
            "Nabídková cena s DPH = MOC × (1 − sleva). "
            "DPH je 21 % a MOC se již bere včetně DPH. Příklad: 999 CZK při slevě 40 % = 599,40 CZK s DPH."
        )
    else:
        st.caption(
            "Nabídková cena bez DPH = MOC × (1 − sleva) / 1,21. "
            "DPH je pevně 21 %. Příklad: 999 CZK při slevě 40 % = 495,37 CZK bez DPH."
        )

    final_offer_df = to_offer_table(
        filtered,
        include_extra_columns=include_extra_columns,
        export_currency=export_currency,
        discount_percent=float(discount_percent),
        vat_rate=0.21,
        price_vat_mode=price_vat_mode,
    )

    file_name = st.text_input("Output filename", value="ua_offer.xlsx")
    if not file_name.lower().endswith(".xlsx"):
        file_name += ".xlsx"

    if final_offer_df.empty:
        st.warning("No products match the current filters.")
    else:
        excel_bytes = write_offer_excel(
            final_offer_df,
            discount_percent=float(discount_percent),
            export_currency=export_currency,
            vat_rate=0.21,
            price_vat_mode=price_vat_mode,
        )
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

            **Style import (XLSX)**: insert styles in cells under one another in the first
            worksheet. `1326799` returns all colours of the base style; `1326799-036`
            (or `1326799/036`) returns only that exact style/colour. **Import mode has precedence over
            all regular offer filters**: product type, gender, material, Top styles, colour, MOC ranges,
            season, end use, silhouette, fit, C/O, minimum availability, minimum style/colour stock,
            minimum number of sizes and Full sizerun. Warehouse selection remains active only for calculating
            and displaying availability; it does not remove imported products from the result.
            If imported products are present in the master but have **zero stock in all three uploaded warehouses**,
            an immediate list and download appear below the import. The file contains the unavailable imported products,
            ranked available alternatives, and a separate sheet for valid imported values not found in master data.
            Alternatives use the warehouses selected for the current offer. The **exact unavailable style/colour** is always
            excluded, but a different available colourway of the same style is a valid option when the customer wants to keep
            the fit and price. The alternative strategy lets you choose between **same style / another colour**, **same colour /
            another style**, or a balanced mix. When another style is requested and a product type can be recognized from the
            name (for example, **Polo**), it is preserved first. The app then prioritizes colour, gender, division, end use and
            cut; fit, material, segment and MOC determine the order among usable matches.

            **Final discount, currency and VAT mode**: choose CZK, EUR or both currencies,
            then whether the final discounted price is **Bez DPH** or **S DPH**. The first
            variant is `MOC × (1 − discount) / 1.21`; the second is `MOC × (1 − discount)`.
            VAT is fixed at 21 %.

            **Central warehouse** = sum of the first two `Week` columns in the central
            warehouse file.
            """
        )


if __name__ == "__main__":
    main()
