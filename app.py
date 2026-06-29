"""UA Offer Generator – Streamlit app."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from offer_core import (
    ALTERNATIVE_ENGINE_VERSION,
    add_selected_alternatives_to_offer,
    apply_filters,
    build_dataset,
    imported_unavailable_with_alternatives,
    read_excel,
    read_style_selectors_excel,
    to_english_display,
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
    st.caption("Generate Under Armour product offers from master data and warehouse files.")

    with st.sidebar:
        st.header("1. Upload files")
        master_file = st.file_uploader(
            "Master data",
            type=["xlsx"],
            help=(
                "Must include EAN, Artikl / Article, size, product name, RRP CZK, RRP EUR and Composition. "
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
    rrp_eur_max_default = int(max(float(data["MOC EUR"].max()), 999))
    rrp_czk_max_default = int(max(float(data["MOC CZK"].max()), 999))

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
        performance_material = st.checkbox("Polyester & performance fibres", value=False)
        cotton_material = st.checkbox("Cotton", value=False)
        only_top_styles = st.checkbox("Only Top styles", value=False)

    with col2:
        colors = st.multiselect(
            "Color group / color logic",
            ["Black", "Dark Blue", "Blue", "Gray", "White", "Red", "Green", "Orange", "Yellow", "Brown"],
            default=["Black", "Dark Blue"],
        )
        price_min = st.number_input("RRP CZK from", min_value=0, value=0, step=100)
        price_max = st.number_input(
            "RRP CZK to",
            min_value=0,
            value=min(999, rrp_czk_max_default),
            step=100,
        )
        price_eur_min = st.number_input("RRP EUR from", min_value=0, value=0, step=5)
        price_eur_max = st.number_input(
            "RRP EUR to",
            min_value=0,
            value=rrp_eur_max_default,
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
            "Minimum total units per style / colour",
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
            "Only full size run",
            value=False,
            help="Mens: S–2XL. Womens: XS–XL. Every core size must have stock in the selected warehouses.",
        )

    with col4:
        style_import_file = st.file_uploader(
            "Import styles / style-colours (XLSX)",
            type=["xlsx"],
            help=(
                "On the first worksheet, enter styles in cells below one another. A value such as 1326799 selects all "
                "available colours of that style; 1326799-036 (or 1326799/036) selects one exact style / colour. "
                "When an import is used, imported values override all standard filters. Selected warehouses are used only "
                "to calculate and display availability."
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
    selected_alternatives_df = None
    replacement_mode = "replace_selected"

    if style_import_file is not None:
        try:
            selected_style_refs = read_style_selectors_excel(style_import_file)
        except Exception as exc:
            st.error(f"Style import failed: {exc}")
            st.stop()

        if not selected_style_refs:
            st.warning(
                "No valid styles were found in the imported file. Use values such as 1326799 or 1326799-036."
            )
            st.stop()

        unmatched = unmatched_style_selectors(data, selected_style_refs)
        alternative_strategy_label = st.radio(
            "Alternative recommendation strategy",
            options=[
                "Both approaches",
                "Same style – another colour",
                "Same colour – different style",
            ],
            index=0,
            help=(
                "Same style retains the cut and usually the price, but offers a different available colour. "
                "Same colour looks for another available style, for example another mens gray golf polo. "
                "Both approaches shows a balanced mix of both routes."
            ),
        )
        alternative_strategy = {
            "Both approaches": "both",
            "Same style – another colour": "same_style",
            "Same colour – different style": "same_color_new_style",
        }[alternative_strategy_label]
        alternative_count = st.selectbox(
            "Number of automatic alternatives",
            options=[3, 5, 8],
            index=1,
            help=(
                "Alternatives are searched only among products available in the warehouses selected in "
                "Availability source used for filtering. With Both approaches, the requested count is split between "
                "the two routes; unused places are filled by the other route."
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
            with st.expander("Unavailable imported products and recommended alternatives", expanded=False):
                st.caption(
                    f"Alternatives {ALTERNATIVE_ENGINE_VERSION}: the exact unavailable style / colour is always excluded. "
                    "Depending on the selected strategy, the app offers either the same style in another available colour, "
                    "or another available style in the same colour. For a different style, it prioritises product family "
                    "(for example, Polo), gender, division, end use and silhouette; then it ranks fit, material, segment, "
                    "RRP and stock. Alternative availability reflects the currently selected warehouses."
                )
                st.subheader("Unavailable imports")
                st.dataframe(to_english_display(unavailable_import_df), use_container_width=True, hide_index=True)
                st.subheader("Recommended alternatives")
                if alternatives_df.empty:
                    st.info("No suitable available alternative was found in the selected warehouses.")
                else:
                    st.dataframe(to_english_display(alternatives_df), use_container_width=True, hide_index=True)

        if not alternatives_df.empty:
            st.subheader("Select alternatives for the final offer")
            st.caption(
                "Tick the products to include in the final offer. The original unavailable RRP is shown beside each option "
                "so you can compare it directly with the alternative RRP. Normally select one preferred replacement per "
                "unavailable item, although multiple alternatives may be selected. Each selected product is included with all "
                "of its EANs, sizes and current availability in the selected warehouses."
            )
            selection_signature = "|".join(
                alternatives_df[["Nedostupný artikl", "Alternativa – artikl"]]
                .fillna("")
                .astype(str)
                .agg("→".join, axis=1)
                .tolist()
            )
            if st.session_state.get("alternative_selection_signature_v13") != selection_signature:
                st.session_state["alternative_selection_signature_v13"] = selection_signature
                st.session_state.pop("alternative_offer_selection_v13", None)

            selection_table = st.data_editor(
                # A dedicated display table keeps all working columns in the core logic unchanged.
                # Its index matches alternatives_df and is used below to recover selected core rows.
                _build_alternative_selection_table(alternatives_df),
                key="alternative_offer_selection_v13",
                hide_index=True,
                use_container_width=True,
                height=min(520, max(180, 42 * (len(alternatives_df) + 1))),
                disabled=[
                    "Imported request", "Unavailable style / colour", "Unavailable product",
                    "Original RRP CZK", "Original RRP EUR", "Alternative style / colour",
                    "Alternative product", "Alternative type", "Alternative RRP CZK", "Alternative RRP EUR",
                    "Selected warehouse availability", "Available sizes in selected warehouses",
                    "Recommendation reason",
                ],
                column_config={
                    "Add to offer": st.column_config.CheckboxColumn(
                        "Add to offer",
                        help="The selected product will be included in the final XLSX export.",
                        default=False,
                    ),
                },
            )
            selected_indices = selection_table.index[
                selection_table["Add to offer"].fillna(False).astype(bool)
            ]
            selected_alternatives_df = alternatives_df.loc[selected_indices].copy()
            if not selected_alternatives_df.empty:
                replacement_mode_label = st.radio(
                    "How should the original unavailable product be handled in the final export?",
                    options=[
                        "Replace it with the selected alternative",
                        "Keep it in the offer with zero stock",
                    ],
                    index=0,
                    horizontal=True,
                    help=(
                        "The first option removes only the exact unavailable style / colour for which you selected "
                        "an alternative. Other imported products remain unchanged."
                    ),
                )
                replacement_mode = (
                    "replace_selected"
                    if replacement_mode_label == "Replace it with the selected alternative"
                    else "keep_original"
                )
                st.info(
                    f"{selected_alternatives_df['Alternativa – artikl'].nunique()} unique alternative "
                    "style / colour item(s) will be included in the final offer."
                )

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
                label="Download unavailable products and alternatives",
                data=substitution_bytes,
                file_name="ua_import_unavailable_and_alternatives.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="The file contains Unavailable imports, Alternatives and, if applicable, Not in master data.",
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
        technical_material=performance_material,
        cotton_material=cotton_material,
        only_top_styles=only_top_styles,
        only_full_sizerun=only_full_sizerun,
        selected_style_refs=selected_style_refs,
    )

    offer_rows = filtered
    if selected_alternatives_df is not None and not selected_alternatives_df.empty:
        offer_rows = add_selected_alternatives_to_offer(
            data=data,
            base_offer=filtered,
            selected_warehouses=selected_warehouses,
            selected_alternatives=selected_alternatives_df,
            replace_selected_unavailable=replacement_mode == "replace_selected",
        )

    # Preview remains neutral: it shows both RRP currencies and does not yet
    # add a negotiated discount. Final currency and discount are chosen below.
    preview_df = to_offer_table(
        offer_rows,
        include_extra_columns=include_extra_columns,
        export_currency="CZK + EUR",
        discount_percent=None,
    )

    st.header("3. Preview")
    full_sizerun_count = (
        offer_rows.loc[offer_rows["Plný sizerun"].eq("Yes"), "Artikl"].nunique()
        if not offer_rows.empty
        else 0
    )
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("EAN rows", f"{len(preview_df):,}")
    kpi2.metric("Style / colour", f"{preview_df['Style / colour'].nunique() if not preview_df.empty else 0:,}")
    kpi3.metric(
        "Selected availability",
        f"{int(offer_rows['Dostupnost ve vybraných skladech'].sum()) if not offer_rows.empty else 0:,}",
    )
    kpi4.metric("Full size runs", f"{full_sizerun_count:,}")
    kpi5.metric(
        "Average RRP CZK",
        f"{offer_rows['MOC CZK'].mean():,.0f}" if not offer_rows.empty else "-",
    )
    kpi6.metric(
        "Average RRP EUR",
        f"{offer_rows['MOC EUR'].mean():,.0f}" if not offer_rows.empty else "-",
    )

    st.dataframe(preview_df.head(500), use_container_width=True, hide_index=True)
    if len(preview_df) > 500:
        st.caption("Preview shows the first 500 rows only. The export contains all selected offer rows.")

    st.header("4. Final pricing and export")
    export_col1, export_col2, export_col3 = st.columns([1, 1, 1])
    with export_col1:
        discount_percent = st.number_input(
            "Discount from VAT-inclusive RRP (%)",
            min_value=0.0,
            max_value=100.0,
            value=0.0,
            step=0.5,
            format="%.1f",
            help="The discount is applied only in the final offer and does not change the selected product list.",
        )
    with export_col2:
        export_currency = st.radio(
            "Currency in final export",
            options=["CZK", "EUR", "CZK + EUR"],
            index=2,
            horizontal=True,
            help="Choose whether the export should contain prices in CZK, EUR or both currencies.",
        )
    with export_col3:
        price_vat_mode = st.radio(
            "Final discounted price",
            options=["Excl. VAT", "Incl. VAT"],
            index=0,
            horizontal=True,
            help="Choose whether the final discounted price is shown excluding or including VAT.",
        )

    if price_vat_mode == "Incl. VAT":
        st.caption(
            "Offer price incl. VAT = RRP × (1 − discount). VAT is 21% and the RRP is already VAT-inclusive. "
            "Example: 999 CZK at a 40% discount = 599.40 CZK incl. VAT."
        )
    else:
        st.caption(
            "Offer price excl. VAT = RRP × (1 − discount) / 1.21. VAT is fixed at 21%. "
            "Example: 999 CZK at a 40% discount = 495.37 CZK excl. VAT."
        )

    final_offer_df = to_offer_table(
        offer_rows,
        include_extra_columns=include_extra_columns,
        export_currency=export_currency,
        discount_percent=float(discount_percent),
        vat_rate=0.21,
        price_vat_mode=price_vat_mode,
    )

    file_name = st.text_input("Output file name", value="ua_offer.xlsx")
    if not file_name.lower().endswith(".xlsx"):
        file_name += ".xlsx"

    if final_offer_df.empty:
        st.warning("No products match the current criteria.")
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

            **Polyester & performance fibres / Cotton**: a row containing `cotton` or `bavlna` in
            `Composition` is Cotton. Polyester & performance fibres means a composition without cotton
            that contains common performance fibres such as polyester, elastane, nylon or polyamide.
            Both selected includes both groups; neither selected leaves material unrestricted.
            The product type **Technical T-shirts** defines only the product category, not the material.

            **Total units / style-colour**: total stock across all EAN rows sharing `Artikl`
            (UA style + colour), calculated from the currently selected warehouses. It is shown in every export.

            **Full size run**: Mens requires S, M, L, XL and 2XL; Womens requires XS, S, M, L and XL.
            Each of those sizes must have stock in the selected warehouses. UA size labels SM / MD / LG and XXL
            are converted automatically.

            **Style import (XLSX)**: insert styles in cells below one another in the first worksheet.
            `1326799` returns all colours of the base style; `1326799-036` (or `1326799/036`) returns only that
            exact style / colour. **Import mode takes precedence over all regular offer filters**: product type,
            gender, material, Top styles, colour, RRP ranges, season, end use, silhouette, fit, C/O, minimum
            availability, minimum style/colour stock, minimum number of sizes and Full size run. Warehouse selection
            remains active only for calculating and displaying availability; it does not remove imported products
            from the result. If imported products are present in the master but have **zero stock in all three uploaded
            warehouses**, an immediate list and a download appear below the import. The file contains unavailable imported
            products, ranked available alternatives and a separate sheet for valid imported values not found in master data.
            Alternatives use the warehouses selected for the current offer. The **exact unavailable style / colour** is
            always excluded, but a different available colourway of the same style is valid when the customer wants to
            preserve fit and RRP. The alternative strategy lets you choose between **same style / another colour**,
            **same colour / different style**, or a balanced mix. **Selected alternatives** can be checked directly
            in the app and flow into the preview and final XLSX report with their current EANs, available sizes and warehouse
            availability. You can choose whether a checked alternative replaces its exact zero-stock import or is added
            alongside it. When a different style is requested and a product family can be recognized from the name
            (for example, **Polo**), it is preserved first. The app then prioritises colour, gender, division, end use
            and silhouette; fit, material, segment and RRP determine the order among usable matches.

            **Final discount, currency and VAT mode**: choose CZK, EUR or both currencies, then whether the final
            discounted price is **Excl. VAT** or **Incl. VAT**. The first variant is `RRP × (1 − discount) / 1.21`;
            the second is `RRP × (1 − discount)`. VAT is fixed at 21%.

            **Central warehouse** = sum of the first two `Week` columns in the central warehouse file.
            """
        )


def _build_alternative_selection_table(alternatives_df):
    """Build the editable, English table used to select final-offer alternatives."""
    # Retain the original index. It links checked UI rows back to the internal
    # alternative records used by add_selected_alternatives_to_offer().
    return pd.DataFrame(
        {
            "Add to offer": False,
            "Imported request": alternatives_df["Importovaný požadavek"],
            "Unavailable style / colour": alternatives_df["Nedostupný artikl"],
            "Unavailable product": alternatives_df["Nedostupný název"],
            "Original RRP CZK": alternatives_df["MOC nedostupného CZK"],
            "Original RRP EUR": alternatives_df["MOC nedostupného EUR"],
            "Alternative style / colour": alternatives_df["Alternativa – artikl"],
            "Alternative product": alternatives_df["Alternativa – název"],
            "Alternative type": alternatives_df["Typ alternativy"],
            "Alternative RRP CZK": alternatives_df["MOC CZK"],
            "Alternative RRP EUR": alternatives_df["MOC EUR"],
            "Selected warehouse availability": alternatives_df["Dostupnost ve vybraných skladech"],
            "Available sizes in selected warehouses": alternatives_df["Dostupné velikosti ve vybraných skladech"],
            "Recommendation reason": alternatives_df["Důvod doporučení"],
        },
        index=alternatives_df.index,
    )


if __name__ == "__main__":
    main()
