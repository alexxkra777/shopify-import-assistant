import os
import sys

import pipeline as pl


def main():
    input_path = sys.argv[1] if len(sys.argv) > 1 else "sample_input.xlsx"
    prefix = sys.argv[2] if len(sys.argv) > 2 else ""
    out_dir = "output"
    os.makedirs(out_dir, exist_ok=True)

    print(f"[INFO] Nacitam vstup: {input_path}")
    print(f"[INFO] AI_MODE = {pl.get_ai_mode()} | Vendor = {pl.STORE_DEFAULT_VENDOR}")

    def progress(done, total):
        print(f"[INFO] AI enrichment: {done}/{total}")

    df, validations, enriched, qc_df, shopify_df = pl.run_pipeline(input_path, progress_cb=progress)

    qc_path = os.path.join(out_dir, f"{prefix}qc_report.xlsx" if prefix else "qc_report.xlsx")
    csv_path = os.path.join(out_dir, f"{prefix}shopify_import.csv" if prefix else "shopify_import.csv")

    with open(qc_path, "wb") as f:
        f.write(pl.qc_dataframe_to_xlsx_bytes(qc_df))
    with open(csv_path, "wb") as f:
        f.write(pl.shopify_dataframe_to_csv_bytes(shopify_df))

    ready_count = (qc_df["Ready_for_Import"] == "ANO").sum()
    not_ready_count = (qc_df["Ready_for_Import"] == "NE").sum()
    warn_count = (qc_df["Warnings / Review"] != "").sum()

    print("\n===== SOUHRN =====")
    print(f"Celkem produktu ve vstupu:            {len(df)}")
    print(f"Pripraveno k importu (bez blok. chyb): {ready_count}")
    print(f"Neni pripraveno (blokujici chyby):     {not_ready_count}")
    print(f"S varovanim k rucni kontrole:          {warn_count}")
    print(f"\nQC report:      {qc_path}")
    print(f"Shopify import: {csv_path}  ({len(shopify_df)} radku)")


if __name__ == "__main__":
    main()
