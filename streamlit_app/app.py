#!/usr/bin/env python3
"""
Streamlit aplikacija za generisanje PDF službenih putnih naloga.
Čita podatke iz Google Sheets-a (javni link) ili uploadovanog Excel fajla.
"""

import streamlit as st
import openpyxl
import pandas as pd
from pypdf import PdfReader, PdfWriter
from datetime import datetime, timedelta
import os
import re
import zipfile
import io
import tempfile
from pathlib import Path

# ================== KONFIGURACIJA ==================
TEMPLATE_PATH = Path(__file__).parent / "template.pdf"
# ==================================================

st.set_page_config(
    page_title="Generisanje putnih naloga",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Generisanje službenih putnih naloga")
st.markdown("Učitaj podatke iz **Google Sheets**-a ili Excel fajla, pa pritisni dugme da generišeš PDF-ove.")


def parse_date(dstr):
    """15.08.2026 ili 15.8.2026 → datetime"""
    dstr = str(dstr).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(dstr, fmt)
        except ValueError:
            continue
    raise ValueError(f"Nepoznat format datuma: {dstr}")


def format_date(dt):
    """datetime → 15.8.26.g"""
    return f"{dt.day}.{dt.month}.{str(dt.year)[2:]}.g"


def format_amount(val):
    """3471 → 3.471,00"""
    try:
        num = float(val)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def load_workers_from_excel(file_bytes):
    """Učitava radnike iz Excel fajla (sheet Timovi)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "Timovi" not in wb.sheetnames:
        st.error("U Excel fajlu ne postoji sheet sa imenom **Timovi**.")
        return []
    ws = wb["Timovi"]
    workers = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        workers.append({
            "tim_id": row[0],
            "ime": str(row[1]).strip(),
            "pozicija": str(row[2]).strip() if row[2] else "Službenik obezbeđenja",
            "datum_pocetka": row[3],
            "datum_kraja": row[4],
            "mesto": str(row[8]).strip() if len(row) > 8 and row[8] else "Zaječar",
            "zadatak": str(row[9]).strip() if len(row) > 9 and row[9] else "Obavljanje službenog zadatka na terenu",
            "prevoz": str(row[10]).strip() if len(row) > 10 and row[10] else "Ford Fokus 1.4",
            "registracija": str(row[11]).strip() if len(row) > 11 and row[11] else "ZA095LE",
            "dnevni_iznos": row[12] if len(row) > 12 else 3471,
            "teret": str(row[14]).strip() if len(row) > 14 and row[14] else "poslodavca",
            "organizacija": str(row[15]).strip() if len(row) > 15 and row[15] else "Business Centre - Gold Guard DOO Zaječar",
        })
    return workers


def load_workers_from_google_sheets(sheet_url: str):
    """
    Učitava podatke iz javnog Google Sheets-a preko CSV export-a.
    Sheet mora biti podešen na 'Anyone with the link can view'.
    """
    # Izvuci spreadsheet ID
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        st.error("Neispravan Google Sheets URL.")
        return []

    spreadsheet_id = match.group(1)

    # Pokušaj da nađeš gid za sheet "Timovi" (opciono)
    # Najjednostavnije: koristimo prvi sheet ili eksplicitni gid=0
    # Korisnik može da doda &gid=XXXXX u URL ako zna

    csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid=0"

    try:
        df = pd.read_csv(csv_url)
    except Exception as e:
        st.error(f"Ne mogu da učitam Google Sheet. Proveri da li je deljen kao 'Anyone with the link'. Greška: {e}")
        return []

    # Normalizuj imena kolona
    df.columns = [str(c).strip() for c in df.columns]

    # Mapiranje očekivanih kolona
    col_map = {
        "ime": ["Član tima", "Član tima", "Ime", "Radnik", "Name"],
        "pozicija": ["Pozicija", "Pozicija", "Radno mesto"],
        "datum_pocetka": ["Datum početka", "Datum pocetka", "Početak"],
        "datum_kraja": ["Datum kraja", "Datum kraja", "Kraj"],
        "mesto": ["Mesto", "Lokacija"],
        "zadatak": ["Zadatak", "Opis posla"],
        "prevoz": ["Prevozno sredstvo", "Prevoz"],
        "registracija": ["Registracija", "Reg. oznaka"],
        "dnevni_iznos": ["Dnevni iznos (RSD)", "Dnevni iznos", "Dnevnica"],
        "teret": ["Troškovi padaju na teret", "Teret"],
        "organizacija": ["Organizacija", "Firma"],
    }

    def find_col(possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        return None

    workers = []
    for _, row in df.iterrows():
        ime_col = find_col(col_map["ime"])
        if not ime_col or pd.isna(row.get(ime_col)):
            continue

        def get(key, default=""):
            col = find_col(col_map[key])
            if col and not pd.isna(row.get(col)):
                return str(row[col]).strip()
            return default

        workers.append({
            "tim_id": row.get("Tim ID", ""),
            "ime": get("ime"),
            "pozicija": get("pozicija", "Službenik obezbeđenja"),
            "datum_pocetka": get("datum_pocetka"),
            "datum_kraja": get("datum_kraja"),
            "mesto": get("mesto", "Zaječar"),
            "zadatak": get("zadatak", "Obavljanje službenog zadatka na terenu"),
            "prevoz": get("prevoz", "Ford Fokus 1.4"),
            "registracija": get("registracija", "ZA095LE"),
            "dnevni_iznos": get("dnevni_iznos", 3471),
            "teret": get("teret", "poslodavca"),
            "organizacija": get("organizacija", "Business Centre - Gold Guard DOO Zaječar"),
        })
    return workers


def generate_pdfs(workers) -> bytes:
    """Generiše sve PDF-ove i vraća ZIP kao bytes."""
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template nije pronađen: {TEMPLATE_PATH}")

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for w in workers:
            start = parse_date(w["datum_pocetka"])
            end = parse_date(w["datum_kraja"])
            prevoz_full = f"{w['prevoz']}, {w['registracija']}"
            amount_str = format_amount(w["dnevni_iznos"])
            amount_rsd = f"{amount_str} RSD"
            org_short = "Business Centre - Gold Guard"
            org_full = w["organizacija"]
            mesto = w["mesto"]
            mesto_loc = "Zaječaru" if mesto == "Zaječar" else (mesto + "u" if mesto and not mesto.endswith("u") else mesto)

            current = start
            while current <= end:
                date_str = format_date(current)

                reader = PdfReader(str(TEMPLATE_PATH))
                writer = PdfWriter()
                writer.append(reader)

                field_values = {
                    "fill_2": org_short,
                    "fill_7": w["ime"],
                    "fill_8": w["pozicija"],
                    "fill_10": date_str,
                    "fill_17": prevoz_full,
                    "fill_19": amount_rsd,
                    "fill_20": date_str,
                    "fill_21": org_full,
                    "fill_3": date_str,
                    "fill_6": date_str,
                    "fill_4": "07",
                    "fill_5": "20",
                    "fill_90": "13",
                    "fill_91": "1",
                    "fill_45": amount_str,
                    "fill_29": amount_str,
                    "fill_41": amount_str,
                    "fill_87": amount_str,
                    "fill_89": amount_str,
                    "fill_24": mesto_loc,
                    "fill_25": date_str,
                    "fill_34": mesto_loc,
                    "fill_35": date_str,
                    "fill_33": w["teret"],
                    "fill_3_2": org_short,
                    "fill_6_2": w["ime"],
                    "fill_8_2": w["pozicija"],
                    "fill_11_2": date_str,
                    "fill_24_2": prevoz_full,
                    "fill_27": amount_rsd,
                    "fill_29_2": date_str,
                    "fill_33_2": org_full,
                    "fill_2_2": "Službeni put protekao po planu i bez vanrednih događaja.",
                    "fill_13": w["zadatak"],
                    "fill_16_2": w["zadatak"],
                }

                for page in writer.pages:
                    writer.update_page_form_field_values(page, field_values, auto_regenerate=False)

                safe_name = re.sub(r"[^\w\s-]", "", w["ime"]).replace(" ", "_")
                filename = f"{safe_name}_{current.strftime('%Y%m%d')}.pdf"

                pdf_bytes = io.BytesIO()
                writer.write(pdf_bytes)
                zipf.writestr(filename, pdf_bytes.getvalue())

                current += timedelta(days=1)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# ================== UI ==================

tab1, tab2 = st.tabs(["📊 Google Sheets / Excel", "ℹ️ Uputstvo"])

with tab1:
    st.subheader("1. Izvor podataka")

    source = st.radio(
        "Odakle učitavam podatke?",
        ["Upload Excel fajla (.xlsx)", "Google Sheets (javni link)"],
        horizontal=True,
    )

    workers = []

    if source == "Upload Excel fajla (.xlsx)":
        uploaded = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])
        if uploaded:
            workers = load_workers_from_excel(uploaded.getvalue())
            st.success(f"Učitano **{len(workers)}** radnika iz Excel fajla.")
    else:
        sheet_url = st.text_input(
            "Nalepi link ka Google Sheets-u",
            placeholder="https://docs.google.com/spreadsheets/d/XXXXXXXXXXXX/edit",
            help="Sheet mora biti deljen kao: Anyone with the link → Viewer",
        )
        if sheet_url:
            with st.spinner("Učitavam podatke iz Google Sheets-a..."):
                workers = load_workers_from_google_sheets(sheet_url)
            if workers:
                st.success(f"Učitano **{len(workers)}** radnika iz Google Sheets-a.")

    if workers:
        st.subheader("2. Pregled podataka")
        preview_df = pd.DataFrame(workers)[["ime", "pozicija", "datum_pocetka", "datum_kraja", "mesto", "prevoz"]]
        st.dataframe(preview_df, use_container_width=True)

        # Izračunaj koliko PDF-ova će biti
        total_pdfs = 0
        for w in workers:
            try:
                start = parse_date(w["datum_pocetka"])
                end = parse_date(w["datum_kraja"])
                total_pdfs += (end - start).days + 1
            except Exception:
                pass

        st.info(f"Biće generisano **{total_pdfs}** PDF naloga ({len(workers)} radnika).")

        st.subheader("3. Generisanje")
        if st.button("🚀 GENERIŠI PDF NALOGE", type="primary", use_container_width=True):
            with st.spinner(f"Generišem {total_pdfs} PDF-ova... Ovo može potrajati 20–60 sekundi."):
                try:
                    zip_bytes = generate_pdfs(workers)
                    st.success(f"✅ Uspešno generisano **{total_pdfs}** PDF naloga!")
                    st.download_button(
                        label="⬇️ Preuzmi ZIP sa svim PDF-ovima",
                        data=zip_bytes,
                        file_name=f"putni_nalozi_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Greška prilikom generisanja: {e}")
                    st.exception(e)

with tab2:
    st.markdown("""
### Kako koristiti aplikaciju

1. **Pripremi Google Sheet** (ili Excel):
   - Mora postojati sheet/tabela sa kolonama:  
     `Član tima`, `Pozicija`, `Datum početka`, `Datum kraja`, `Mesto`, `Zadatak`,  
     `Prevozno sredstvo`, `Registracija`, `Dnevni iznos (RSD)`, `Troškovi padaju na teret`, `Organizacija`
   - Datumi u formatu `15.08.2026`

2. **Ako koristiš Google Sheets**:
   - Otvori Sheet → **Share** → **Anyone with the link** → **Viewer**
   - Kopiraj link i nalepi ga u aplikaciju

3. **Pritisni dugme „GENERIŠI PDF NALOGE“**
   - Aplikacija će napraviti po jedan PDF za svaki dan svakog radnika
   - Dobiješ ZIP fajl za preuzimanje

### Napomene
- Template PDF je ugrađen u aplikaciju.
- Generisanje 100 PDF-ova traje obično 20–50 sekundi.
- Specijalni karakteri (ć, č, đ, š, ž) su podržani.
""")

st.markdown("---")
st.caption("Automatizacija službenih putnih naloga • Business Centre - Gold Guard")
