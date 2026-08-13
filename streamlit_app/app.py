#!/usr/bin/env python3
"""
Streamlit – popunjavanje ORIGINALNOG PDF obrasca.
Ugrađuje DejaVu font da bi srpska slova i ceo tekst bili vidljivi.
"""

import streamlit as st
import openpyxl
import pandas as pd
import pymupdf
from datetime import datetime, timedelta
from pathlib import Path
import re
import zipfile
import io

APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "template.pdf"
FONT_PATH = APP_DIR / "DejaVuSans.ttf"

st.set_page_config(page_title="Generisanje putnih naloga", page_icon="📄", layout="centered")
st.title("📄 Generisanje službenih putnih naloga")
st.markdown("Originalni obrazac + ispravno prikazivanje teksta.")


def parse_date(dstr):
    dstr = str(dstr).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(dstr, fmt)
        except ValueError:
            continue
    raise ValueError(f"Nepoznat format datuma: {dstr}")


def format_date(dt: datetime) -> str:
    return f"{dt.day}.{dt.month}.{str(dt.year)[2:]}.g"


def format_amount(val) -> str:
    try:
        num = float(val)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def create_nalog_bytes(worker: dict, day: datetime) -> bytes:
    doc = pymupdf.open(str(TEMPLATE_PATH))

    # Ugradi Unicode font na svaku stranu
    for page in doc:
        page.insert_font(fontname="F0", fontfile=str(FONT_PATH))

    date_str = format_date(day)
    amount = format_amount(worker["dnevni_iznos"])
    amount_rsd = f"{amount} RSD"
    prevoz = f"{worker['prevoz']}, {worker['registracija']}"
    org_short = "Business Centre - Gold Guard"
    org_full = worker["organizacija"]
    mesto_loc = "Zaječaru" if worker["mesto"] == "Zaječar" else worker["mesto"]

    values = {
        "fill_2": org_short,
        "fill_7": worker["ime"],
        "fill_8": worker["pozicija"],
        "fill_10": date_str,
        "fill_17": prevoz,
        "fill_19": amount_rsd,
        "fill_20": date_str,
        "fill_21": org_full,
        "fill_3": date_str,
        "fill_6": date_str,
        "fill_4": "07",
        "fill_5": "20",
        "fill_90": "13",
        "fill_91": "1",
        "fill_45": amount,
        "fill_29": amount,
        "fill_41": amount,
        "fill_87": amount,
        "fill_89": amount,
        "fill_24": mesto_loc,
        "fill_25": date_str,
        "fill_34": mesto_loc,
        "fill_35": date_str,
        "fill_33": worker["teret"],
        "fill_3_2": org_short,
        "fill_6_2": worker["ime"],
        "fill_8_2": worker["pozicija"],
        "fill_11_2": date_str,
        "fill_24_2": prevoz,
        "fill_27": amount_rsd,
        "fill_29_2": date_str,
        "fill_33_2": org_full,
        "fill_2_2": "Službeni put protekao po planu i bez vanrednih događaja.",
        "fill_13": worker["zadatak"],
        "fill_16_2": worker["zadatak"],
    }

    # Polja koja moraju da koriste Unicode font
    unicode_fields = {
        "fill_7", "fill_8", "fill_21", "fill_24", "fill_34", "fill_33",
        "fill_6_2", "fill_8_2", "fill_33_2", "fill_2_2", "fill_13", "fill_16_2",
        "fill_2", "fill_3_2", "fill_17", "fill_24_2", "fill_19", "fill_27",
    }

    for page in doc:
        for w in page.widgets():
            if w.field_name not in values:
                continue
            w.field_value = values[w.field_name]

            if w.field_name in unicode_fields:
                w.text_font = "F0"
                w.text_fontsize = 8.5
            else:
                # datumi i brojevi – malo manji da stanu
                if w.field_name in ("fill_3", "fill_6", "fill_10", "fill_11_2", "fill_20", "fill_25", "fill_35", "fill_29_2"):
                    w.text_fontsize = 8.0
                else:
                    w.text_fontsize = 9.0

            try:
                w.update()
            except Exception:
                pass

    doc.need_appearances = True
    pdf_bytes = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    return pdf_bytes


def load_workers_from_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "Timovi" not in wb.sheetnames:
        st.error("Nema sheet-a **Timovi**.")
        return []
    ws = wb["Timovi"]
    workers = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        workers.append({
            "ime": str(row[1]).strip(),
            "pozicija": str(row[2]).strip() if row[2] else "Službenik obezbeđenja",
            "datum_pocetka": row[3],
            "datum_kraja": row[4],
            "mesto": str(row[8]).strip() if len(row) > 8 and row[8] else "Zaječar",
            "zadatak": str(row[9]).strip() if len(row) > 9 and row[9] else "Obavljanje službenog zadatka na terenu",
            "prevoz": str(row[10]).strip() if len(row) > 10 and row[10] else "Ford Fokus 1.4",
            "registracija": str(row[11]).strip() if len(row) > 11 and row[11] else "ZA095LE",
            "dnevni_iznos": row[12] if len(row) > 12 and row[12] else 3471,
            "teret": str(row[14]).strip() if len(row) > 14 and row[14] else "poslodavca",
            "organizacija": str(row[15]).strip() if len(row) > 15 and row[15] else "Business Centre - Gold Guard DOO Zaječar",
        })
    return workers


def load_workers_from_google_sheets(sheet_url: str):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        st.error("Neispravan Google Sheets URL.")
        return []
    spreadsheet_id = match.group(1)
    csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid=0"
    try:
        df = pd.read_csv(csv_url)
    except Exception as e:
        st.error(f"Ne mogu da učitam Sheet: {e}")
        return []
    df.columns = [str(c).strip() for c in df.columns]
    col_map = {
        "ime": ["Član tima", "Ime", "Radnik", "Name"],
        "pozicija": ["Pozicija", "Radno mesto"],
        "datum_pocetka": ["Datum početka", "Datum pocetka", "Početak"],
        "datum_kraja": ["Datum kraja", "Kraj"],
        "mesto": ["Mesto", "Lokacija"],
        "zadatak": ["Zadatak", "Opis posla"],
        "prevoz": ["Prevozno sredstvo", "Prevoz"],
        "registracija": ["Registracija", "Reg. oznaka"],
        "dnevni_iznos": ["Dnevni iznos (RSD)", "Dnevni iznos", "Dnevnica"],
        "teret": ["Troškovi padaju na teret", "Teret"],
        "organizacija": ["Organizacija", "Firma"],
    }
    def find_col(names):
        for n in names:
            if n in df.columns: return n
        return None
    workers = []
    for _, row in df.iterrows():
        ime_col = find_col(col_map["ime"])
        if not ime_col or pd.isna(row.get(ime_col)): continue
        def get(key, default=""):
            col = find_col(col_map[key])
            if col and not pd.isna(row.get(col)): return str(row[col]).strip()
            return default
        workers.append({
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
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for w in workers:
            start = parse_date(w["datum_pocetka"])
            end = parse_date(w["datum_kraja"])
            current = start
            while current <= end:
                pdf_bytes = create_nalog_bytes(w, current)
                safe_name = re.sub(r"[^\w\s-]", "", w["ime"]).replace(" ", "_")
                filename = f"{safe_name}_{current.strftime('%Y%m%d')}.pdf"
                zipf.writestr(filename, pdf_bytes)
                current += timedelta(days=1)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# ===== UI =====
tab1, tab2 = st.tabs(["📊 Podaci", "ℹ️ Uputstvo"])

with tab1:
    source = st.radio("Izvor", ["Upload Excel (.xlsx)", "Google Sheets link"], horizontal=True)
    workers = []
    if source.startswith("Upload"):
        up = st.file_uploader("Excel fajl", type=["xlsx"])
        if up:
            workers = load_workers_from_excel(up.getvalue())
            st.success(f"Učitano {len(workers)} radnika")
    else:
        url = st.text_input("Google Sheets URL")
        if url:
            with st.spinner("Učitavam..."):
                workers = load_workers_from_google_sheets(url)
            if workers:
                st.success(f"Učitano {len(workers)} radnika")

    if workers:
        st.dataframe(pd.DataFrame(workers)[["ime", "pozicija", "datum_pocetka", "datum_kraja", "mesto"]], use_container_width=True)
        total = sum((parse_date(w["datum_kraja"]) - parse_date(w["datum_pocetka"])).days + 1 for w in workers)
        st.info(f"Generisaće se **{total}** PDF-ova")
        if st.button("🚀 GENERIŠI PDF NALOGE", type="primary", use_container_width=True):
            with st.spinner("Generišem..."):
                try:
                    z = generate_pdfs(workers)
                    st.success(f"✅ {total} PDF-ova spremno")
                    st.download_button("⬇️ Preuzmi ZIP", data=z,
                        file_name=f"putni_nalozi_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                        mime="application/zip", use_container_width=True)
                except Exception as e:
                    st.error(str(e))
                    st.exception(e)

with tab2:
    st.markdown("Koristi se **originalni PDF obrazac**. Font DejaVu je ugrađen da bi se sav tekst (uključujući ć, č, š, ž, đ) lepo video.")

st.caption("Originalni template + ugrađeni Unicode font")
