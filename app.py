import hashlib
import os
import pandas as pd
import streamlit as st

import pipeline as pl

st.set_page_config(page_title="Shopify Import Assistant", page_icon="🛍️", layout="wide")

EDIT_COLS = ["title", "vendor", "category", "price", "sku", "image_url"]
EDIT_LABELS = {
    "title": "Nazev", "vendor": "Znacka", "category": "Kategorie",
    "price": "Cena", "sku": "SKU", "image_url": "Obrazek (URL)",
}


def get_api_key():
    """Klic z prostredi, jinak z .streamlit/secrets.toml (negitovana, viz
    .gitignore) - v obou pripadech se propaguje do os.environ, protoze
    pipeline.py cte primo os.environ, ne st.secrets."""
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        try:
            key = st.secrets.get("ANTHROPIC_API_KEY", "")
        except Exception:
            key = ""
        if key:
            os.environ["ANTHROPIC_API_KEY"] = key
    return key


def init_state():
    for k in ["df", "validations", "enriched", "processed"]:
        if k not in st.session_state:
            st.session_state[k] = None
    if st.session_state["processed"] is None:
        st.session_state["processed"] = False
    if "editor_revision" not in st.session_state:
        st.session_state["editor_revision"] = 0


def format_price_for_edit(price) -> str:
    if price is None:
        return ""
    return str(int(price)) if float(price).is_integer() else f"{price:.2f}"


def issues_text(v: dict) -> str:
    parts = []
    if v["hard_issues"]:
        parts.append("BLOKUJE: " + "; ".join(v["hard_issues"]))
    if v["soft_issues"]:
        parts.append("Zkontrolovat: " + "; ".join(v["soft_issues"]))
    return " | ".join(parts)


def run_processing(uploaded_file):
    progress = st.progress(0.0, text="Zpracovavam...")

    def cb(done, total):
        if total:
            progress.progress(min(done / total, 1.0), text=f"AI zpracovava produkty... {done}/{total}")

    df = pl.load_input(uploaded_file)
    validations = pl.recompute_validations(df)
    enriched = pl.ai_enrich_all(df, validations, progress_cb=cb)
    validations = pl.flag_missing_ai_content(validations, enriched)
    progress.empty()

    st.session_state["df"] = df
    st.session_state["validations"] = validations
    st.session_state["enriched"] = enriched
    st.session_state["processed"] = True
    st.session_state["editor_revision"] = 0


def rerun_selected(edited_indices, edited_df):
    df = st.session_state["df"]
    for idx in edited_indices:
        for col in EDIT_COLS:
            value = edited_df.at[idx, col]
            if col == "price":
                value = pl.clean_price(value)
            df.at[idx, col] = value

    progress = st.progress(0.0, text="Znovu kontroluji opravene radky...")

    def cb(done, total):
        if total:
            progress.progress(min(done / total, 1.0), text=f"AI zpracovava opravene produkty... {done}/{total}")

    validations, enriched = pl.revalidate_and_reenrich(
        df, st.session_state["enriched"], edited_indices, progress_cb=cb
    )
    progress.empty()

    st.session_state["validations"] = validations
    st.session_state["enriched"] = enriched


# ---------------------------------------------------------------------------

init_state()

st.title("🛍️ Shopify Import Assistant")
st.caption("Nahrajte produktovy Excel. AI zkontroluje udaje, doplni popis/SEO/AEO obsah a oznaci, co jeste potrebuje pozornost, nez to pujde do Shopify.")

get_api_key()  # zajisti, ze klic ze secrets.toml (pokud tam je) je v os.environ pro pipeline.py

uploaded = st.file_uploader("Vstupni Excel s produkty (.xlsx)", type=["xlsx"])

if uploaded is not None:
    # Hash obsahu (ne jmeno souboru) - odhali i kdyz uzivatel nahraje jiny soubor
    # se stejnym nazvem, nebo upravenou verzi puvodniho souboru. Bez tohohle by
    # vysledky ze STAREHO souboru zustaly zobrazene, dokud by uzivatel znovu
    # rucne neklikl "Zpracovat soubor" - a vypadalo by to jako spatna/vymyslena data.
    current_hash = hashlib.md5(uploaded.getvalue()).hexdigest()
    if st.session_state.get("last_uploaded_hash") != current_hash:
        st.session_state["processed"] = False
        st.session_state["last_uploaded_hash"] = current_hash

col_a, col_b = st.columns([1, 4])
with col_a:
    process_clicked = st.button("▶️ Zpracovat soubor", type="primary", disabled=uploaded is None)

if process_clicked and uploaded is not None:
    run_processing(uploaded)

if uploaded is not None and not st.session_state["processed"] and not process_clicked:
    st.info("Vybran novy soubor - klikni na 'Zpracovat soubor' pro nacteni aktualnich dat.")

if st.session_state["processed"]:
    df = st.session_state["df"]
    validations = st.session_state["validations"]
    enriched = st.session_state["enriched"]
    qc_df = pl.build_qc_dataframe(df, validations, enriched)

    ready_mask = qc_df["Ready_for_Import"] == "ANO"
    warn_mask = qc_df["Warnings / Review"] != ""
    needs_attention_mask = (~ready_mask) | warn_mask

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Celkem produktu", len(df))
    m2.metric("✅ Pripraveno k importu", int(ready_mask.sum()))
    m3.metric("⛔ Blokovano chybou", int((~ready_mask).sum()))
    m4.metric("⚠️ S varovanim", int(warn_mask.sum()))

    st.divider()

    needs_attention_positions = [i for i in range(len(validations)) if needs_attention_mask.iloc[i]]

    if needs_attention_positions:
        st.subheader("⚠️ Vyzaduje pozornost")
        st.caption("Radky s blokujici chybou nebo varovanim. Kategorie/znacka jsou predvyplnene "
                   "AI navrhem tam, kde chybely - staci zkontrolovat, pripadne upravit, a ulozit.")

        if st.button("🪄 Automaticky opravit, co lze (popis, SEO title/description/tagy)"):
            modified = pl.auto_fix_fixable(df, validations, enriched)

            # radky, kde AI SEO title/description/tagy/popis chybi - vcetne radku
            # bez nazvu produktu. Pro ty pouzijeme jasne oznaceny docasny nazev jen
            # pro ucely AI generovani (do skutecnych dat se nezapisuje), aby pole
            # nezustavala navzdy prazdna - vysledek je jasne poznat jako placeholder
            # a je potreba ho po doplneni skutecneho nazvu znovu zkontrolovat.
            ai_missing_positions = [
                i for i, v in enumerate(validations)
                if v["soft_fields"] & {"ai_seo_title", "ai_seo_description", "ai_tags", "ai_body_html"}
            ]
            if ai_missing_positions:
                subset_validations = [validations[i] for i in ai_missing_positions]
                df_for_ai = df.copy()
                for v in subset_validations:
                    idx = v["index"]
                    if not pl.clean_text(df_for_ai.loc[idx, "title"]):
                        df_for_ai.loc[idx, "title"] = f"Neznamy produkt (radek {idx}) - DOPLNTE NAZEV"

                progress = st.progress(0.0, text="Doplnuji chybejici AI obsah...")

                def cb(done, total):
                    if total:
                        progress.progress(min(done / total, 1.0), text=f"AI doplnuje obsah... {done}/{total}")

                subset_enriched = pl.ai_enrich_all(df_for_ai, subset_validations, progress_cb=cb)
                progress.empty()
                for i, ai in zip(ai_missing_positions, subset_enriched):
                    enriched[i] = ai
                modified |= {validations[i]["index"] for i in ai_missing_positions}

            if modified:
                new_validations = pl.flag_missing_ai_content(pl.recompute_validations(df), enriched)
                st.session_state["validations"] = new_validations
                st.session_state["enriched"] = enriched
                st.session_state["editor_revision"] += 1
                st.rerun()
            else:
                st.info("Nic k automatickemu opravu tady neni - zbyvajici problemy "
                        "(chybejici SKU, spatna cena, duplicity) vyzaduji rucni zasah.")
        st.caption("Doplni popis produktu a AI SEO title/description/tagy tam, kde chybi - i u radku "
                   "bez nazvu (pouzije se docasny placeholder oznaceny 'DOPLNTE NAZEV', ktery je nutne "
                   "po zadani skutecneho nazvu zkontrolovat). SKU se NEvymysli - chybejici/duplicitni "
                   "SKU vyzaduje vzdy rucni zadani.")

        edit_rows = {}
        problem_rows = {}
        for pos in needs_attention_positions:
            v = validations[pos]
            idx = v["index"]
            ai = enriched[pos]
            row = df.loc[idx]
            edit_rows[idx] = {
                "title": pl.clean_text(row.get("title")),
                "vendor": pl.clean_text(row.get("vendor")) or pl.STORE_DEFAULT_VENDOR,
                "category": pl.clean_text(row.get("category")) or ai.get("ai_suggested_category", ""),
                "price": format_price_for_edit(v["price_clean"]),
                "sku": pl.clean_text(row.get("sku")),
                "image_url": pl.clean_text(row.get("image_url")),
                "AI SEO Title": ai.get("ai_seo_title", ""),
                "AI SEO Description": ai.get("ai_seo_description", ""),
                "AI Tags": ai.get("ai_tags", ""),
            }
            problem_rows[idx] = issues_text(v)

        edit_df = pd.DataFrame.from_dict(edit_rows, orient="index")
        edit_df["problemy"] = pd.Series(problem_rows)
        edit_df = edit_df.rename(columns={**EDIT_LABELS, "problemy": "Problemy"})
        AI_READONLY_COLS = ["AI SEO Title", "AI SEO Description", "AI Tags"]

        label_to_field = {label: field for field, label in EDIT_LABELS.items()}
        label_to_field.update({
            "AI SEO Title": "ai_seo_title",
            "AI SEO Description": "ai_seo_description",
            "AI Tags": "ai_tags",
        })
        hard_fields_by_idx = {v["index"]: v["hard_fields"] for v in validations}
        soft_fields_by_idx = {v["index"]: v["soft_fields"] for v in validations}

        def highlight_problem_cells(row):
            hard = hard_fields_by_idx.get(row.name, set())
            soft = soft_fields_by_idx.get(row.name, set())
            styles = []
            for col in row.index:
                field = label_to_field.get(col)
                if field and field in hard:
                    styles.append("background-color: #f2cbcb; color: #1a1a1a")
                elif field and field in soft:
                    styles.append("background-color: #fce8b2; color: #1a1a1a")
                else:
                    styles.append("")
            return styles

        st.caption("Barevne zvyraznene bunky ukazuji presne, KDE je problem "
                   "(cervena = blokujici chyba, zluta = varovani). Uprava probiha v tabulce pod tim.")
        st.dataframe(edit_df.style.apply(highlight_problem_cells, axis=1), use_container_width=True)

        # Klic zavisi na revizi: sada radku "vyzadujicich pozornost" se po kazdem
        # ulozeni meni (ridici radky mizi, poradi se posouva). Pri stejnem klici
        # by Streamlit mohl znovupouzit stare pozicni edity na jine (nove) radky -
        # zmenou klice se widget pri kazde revizi vytvori od znova.
        # Cena je textove pole (ne NumberColumn) - v teto verzi Streamlit
        # NumberColumn vykresluje chybejici cenu jako doslovny text "None"
        # misto prazdne bunky.
        st.caption("✏️ Zde oprav konkretni hodnoty (nazev, znacka, kategorie, cena, SKU, obrazek) "
                   "primo v bunkach. Po dokonceni klikni na 'Ulozit opravy a znovu zkontrolovat' - "
                   "radek se prevaliduje a zmizi odsud, jakmile bude v poradku.")
        edited = st.data_editor(
            edit_df,
            column_config={
                "Problemy": st.column_config.TextColumn("Problemy", disabled=True, width="large"),
                **{col: st.column_config.TextColumn(col, disabled=True) for col in AI_READONLY_COLS},
            },
            use_container_width=True,
            key=f"editor_{st.session_state['editor_revision']}",
        )

        if st.button("💾 Ulozit opravy a znovu zkontrolovat", type="primary"):
            reverse_labels = {v_: k for k, v_ in EDIT_LABELS.items()}
            edited_internal = edited.rename(columns=reverse_labels)[EDIT_COLS]
            rerun_selected(list(edit_df.index), edited_internal)
            st.session_state["editor_revision"] += 1
            st.rerun()
    else:
        st.success("Vsechny produkty jsou pripravene k importu. 🎉")

    st.divider()

    if int(ready_mask.sum()) > 0:
        shopify_df = pl.build_shopify_dataframe(df, validations, enriched)
        shopify_df_all = pl.build_shopify_dataframe(df, validations, enriched, include_blocked=True)
        st.subheader("✅ Nahled pripravenych produktu")
        st.dataframe(
            qc_df.loc[ready_mask].drop(columns=["_hard_fields", "_soft_fields"], errors="ignore"),
            use_container_width=True,
        )

        st.caption("Shopify import (CSV): jen pripravene produkty, ve strukture pro import. "
                   "QC report (XLSX): prehledny seznam ke kontrole. Varianta 'vse' obsahuje i blokovane radky pro audit.")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ Shopify import - pripravene (CSV)",
                data=pl.shopify_dataframe_to_csv_bytes(shopify_df),
                file_name="shopify_import.csv",
                mime="text/csv",
                type="primary",
            )
        with c2:
            st.download_button(
                "⬇️ Shopify import - vse (CSV)",
                data=pl.shopify_dataframe_to_csv_bytes(shopify_df_all),
                file_name="shopify_import_vse.csv",
                mime="text/csv",
                help="Vcetne blokovanych radku - pro kontrolu, NE pro primy import do Shopify.",
            )

        c3, c4 = st.columns(2)
        with c3:
            st.download_button(
                "⬇️ QC report - pripravene (XLSX)",
                data=pl.qc_dataframe_to_xlsx_bytes(qc_df.loc[ready_mask]),
                file_name="pripravene_produkty.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Jen radky Ready_for_Import = ANO, prehledne sloupce jako v QC reportu.",
            )
        with c4:
            st.download_button(
                "⬇️ QC report - vse (XLSX)",
                data=pl.qc_dataframe_to_xlsx_bytes(qc_df),
                file_name="qc_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Kompletni audit vsech produktu vcetne blokovanych - pro kontrolu, ne pro import.",
            )
    else:
        st.info("Zatim neni pripraveny zadny produkt k exportu - opravte oznacene radky vyse.")
