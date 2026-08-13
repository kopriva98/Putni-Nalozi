# Generisanje službenih putnih naloga (Streamlit)

## Šta radi
- Učitava podatke iz **Google Sheets**-a (javni link) ili Excel fajla
- Generiše po jedan PDF nalog za svakog radnika i svaki dan
- Daje ZIP fajl za preuzimanje

## Kako pokrenuti lokalno

```bash
cd streamlit_app
pip install -r requirements.txt
streamlit run app.py
```

## Kako objaviti besplatno na Streamlit Community Cloud

1. Napravi GitHub nalog (ako nemaš)
2. Napravi novi repository i upload-uj ceo folder `streamlit_app` (app.py, requirements.txt, template.pdf)
3. Idi na https://share.streamlit.io
4. Poveži GitHub nalog → New app → izaberi repository i `app.py`
5. Klikni Deploy

Posle toga dobijaš javni link tipa:
`https://tvoje-ime-generisi-naloge.streamlit.app`

## Google Sheets podešavanje

Da bi aplikacija mogla da čita tvoj Sheet:
1. Otvori Google Sheet
2. Share → **Anyone with the link** → Viewer
3. Kopiraj link i nalepi ga u aplikaciju

## Struktura podataka (sheet "Timovi")

| Tim ID | Član tima | Pozicija | Datum početka | Datum kraja | ... | Prevozno sredstvo | Registracija | Dnevni iznos (RSD) | ... |
|--------|-----------|----------|---------------|-------------|-----|-------------------|--------------|--------------------|-----|
| 1      | Marko Petrović | Službenik obezbeđenja | 15.08.2026 | 24.08.2026 | ... | Ford Fokus 1.4 | ZA095LE | 3471 | ... |
