import streamlit as st
import pandas as pd
import io
from docx import Document
from docx.shared import Pt, Inches

# Ustawienie strony
st.set_page_config(layout="wide", page_title="Przeglądarka Dyskursu", page_icon="📚")

st.title("📚 Zaawansowana Przeglądarka Dyskursu")
st.markdown("Filtruj bazę danych, przeszukuj teksty i generuj eleganckie raporty do Worda/PDF.")
st.markdown("Autor wtyczki Emil C.")

# --- FUNKCJA GENEROWANIA RAPORTU WORD (DOCX) ---
def stworz_raport_docx(dataframe, aktywne_kolumny_tematów):
    doc = Document()
    
    # Stylizacja dokumentu Word
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)
    
    title = doc.add_heading('Raport z Analizy Dyskursu', level=0)
    title.alignment = 1 # Wyśrodkowanie
    
    doc.add_paragraph(f"Wygenerowano automatycznie. Liczba znalezionych dokumentów: {len(dataframe)}\n")
    doc.add_paragraph("-" * 50)
    
    for idx, row in dataframe.iterrows():
        tytul = row.get(next((c for c in dataframe.columns if 'tytuł' in str(c).lower()), 'Bez tytułu'), 'Bez tytułu')
        autor = row.get(next((c for c in dataframe.columns if 'autor' in str(c).lower()), '-'), '-')
        rok = row.get(next((c for c in dataframe.columns if 'rok' in str(c).lower()), '-'), '-')
        zrodlo = row.get(next((c for c in dataframe.columns if 'źródło' in str(c).lower() or 'zrodlo' in str(c).lower()), '-'), '-')
        
        # Nagłówek pojedynczego wpisu
        p_head = doc.add_heading(f"📄 {tytul}", level=2)
        p_meta = doc.add_paragraph()
        p_meta.add_run(f"Autor: {autor} | Rok: {rok} | Źródło: {zrodlo}").italic = True
        
        # Dodawanie treści z kategorii tematycznych
        for col_name in dataframe.columns:
            if " -> " in col_name:
                kategoria, typ = col_name.split(" -> ")
                
                # Jeśli użytkownik filtruje po konkretnej kategorii, wyciągamy głównie ją
                wartosc = row[col_name]
                if wartosc != "Brak danych" and str(wartosc).strip() != "":
                    p_content = doc.add_paragraph()
                    if typ.lower() == "cytat":
                        p_content.add_run(f"[{kategoria.upper()} - CYTAT]: ").bold = True
                        p_content.add_run(f"\"{wartosc}\"")
                    elif typ.lower() == "opis":
                        p_content.add_run(f"[{kategoria.upper()} - OPIS/ANALIZA]: ").bold = True
                        p_content.add_run(str(wartosc))
                    elif 'etykieta' in typ.lower():
                        p_content.add_run(f"🏷️ Etykieta ({kategoria}): ").italic = True
                        p_content.add_run(str(wartosc))
                        
        doc.add_paragraph("\n" + "_"*30 + "\n")
        
    # Zapis do pamięci podręcznej (by Streamlit mógł to pobrać)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# --- MODUŁ WGRYWANIA PLIKU ---
wgrany_plik = st.file_uploader("Wgraj swój plik Excel (.xlsx) lub CSV:", type=["csv", "xlsx"])

if wgrany_plik is not None:
    try:
        if wgrany_plik.name.endswith('.xlsx'):
            df = pd.read_excel(wgrany_plik, header=[0, 1])
        else:
            try:
                df = pd.read_csv(wgrany_plik, header=[0, 1], sep=';')
                if len(df.columns) < 3:
                    wgrany_plik.seek(0)
                    df = pd.read_csv(wgrany_plik, header=[0, 1], sep=',')
            except:
                wgrany_plik.seek(0)
                df = pd.read_csv(wgrany_plik, header=[0, 1])
                
        # Naprawa nagłówków
        nowe_kolumny = []
        ostatnia_kategoria = "" 
        
        for col in df.columns:
            kat_glowna = str(col[0]).replace('\n', ' ').strip()
            podkategoria = str(col[1]).replace('\n', ' ').strip()
            if "Unnamed" not in kat_glowna:
                ostatnia_kategoria = kat_glowna
            if "Unnamed" in podkategoria:
                nowe_kolumny.append(kat_glowna)
            else:
                nowe_kolumny.append(f"{ostatnia_kategoria} -> {podkategoria}")
                
        df.columns = nowe_kolumny
        df = df.fillna("Brak danych")

        # Bezpieczne szukanie kolumn meta-danych
        kol_autor = next((col for col in df.columns if 'autor' in str(col).lower()), None)
        kol_rok = next((col for col in df.columns if 'rok' in str(col).lower()), None)

        # Dynamiczne wyciągnięcie dostępnych kategorii tematycznych z nagłówków
        lista_kategorii = sorted(list(set([col.split(" -> ")[0] for col in df.columns if " -> " in col])))

        # --- PASEK BOCZNY (ZAAWANSOWANE FILTRY) ---
        st.sidebar.header("🔍 Filtry i Wyszukiwanie")

        # 1. Wyszukiwarka tekstowa
        szukana_fraza = st.sidebar.text_input("Szukaj słowa w tekstach:")

        # 2. Filtr Autora
        autorzy = ["Wszyscy"] + sorted([str(a) for a in df[kol_autor].unique() if str(a) != "Brak danych"])
        wybrany_autor = st.sidebar.selectbox("Wybierz autora:", autorzy)

        # 3. Filtr Roku
        lata = ["Wszystkie"] + sorted([str(r) for r in df[kol_rok].unique() if str(r) != "Brak danych"])
        wybrany_rok = st.sidebar.selectbox("Wybierz rok:", lata)

        # 4. NOWOŚĆ: Filtr kategorii tematycznych
        wybrana_kategoria = st.sidebar.selectbox("Wybierz kategorię badawczą:", ["Wszystkie"] + lista_kategorii)

        # --- SEKCJA FILTROWANIA DANYCH ---
        df_filtered = df.copy()
        
        if wybrany_autor != "Wszyscy":
            df_filtered = df_filtered[df_filtered[kol_autor].astype(str) == wybrany_autor]
        if wybrany_rok != "Wszystkie":
            df_filtered = df_filtered[df_filtered[kol_rok].astype(str) == wybrany_rok]
            
        # Filtrowanie po kategorii (warunek: w danej kategorii musi być wpis inny niż 'Brak danych')
        if wybrana_kategoria != "Wszystkie":
            powiazane_kolumny = [c for c in df.columns if c.startswith(f"{wybrana_kategoria} ->")]
            # Sprawdzamy czy którykolwiek z podfolderów kategorii (cytat/opis) ma zawartość
            warunek = df_filtered[powiazane_kolumny].apply(lambda row: any(str(x) != "Brak danych" and str(x).strip() != "" for x in row), axis=1)
            df_filtered = df_filtered[warunek]

        # Filtrowanie wyszukiwarką tekstową
        if szukana_fraza:
            szukana_fraza = szukana_fraza.lower()
            tekstowy_warunek = df_filtered.apply(lambda row: any(szukana_fraza in str(x).lower() for x in row), axis=1)
            df_filtered = df_filtered[tekstowy_warunek]

        # --- PANEL AKCJI (DRUKOWANIE/EXPORT) ---
        st.sidebar.write("---")
        st.sidebar.header("🖨️ Eksport danych")
        
        if len(df_filtered) > 0:
            # Generowanie pliku docx w locie
            file_docx = stworz_raport_docx(df_filtered, lista_kategorii)
            
            st.sidebar.download_button(
                label="📥 Pobierz jako plik WORD (DOCX)",
                data=file_docx,
                file_name=f"Raport_Dyskursu_{wybrana_kategoria}_{wybrany_autor}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.sidebar.caption("💡 Po pobraniu otwórz plik w Wordzie i wydrukuj lub zapisz jako PDF.")
        else:
            st.sidebar.warning("Brak rekordów do eksportu.")

        st.subheader(f"📊 Znaleziono rekordów spełniających kryteria: {len(df_filtered)}")

        # --- WYŚWIETLANIE REKORDÓW ---
        st.write("---")

        for index, row in df_filtered.iterrows():
            tytul = row.get(next((c for c in df.columns if 'tytuł' in str(c).lower()), 'Brak tytułu'), 'Brak tytułu')
            autor = row.get(kol_autor, '-')
            rok = row.get(kol_rok, '-')
            zrodlo = row.get(next((c for c in df.columns if 'źródło' in str(c).lower() or 'zrodlo' in str(c).lower()), '-'), '-')
            
            # NOWOŚĆ: Używamy expandera, aby strona była czysta i zwarta
            with st.expander(f"📄 {tytul} ({autor} - {rok})"):
                st.caption(f"**Pełne źródło:** {zrodlo}")
                
                # Przechodzimy po kolumnach
                for col_name in df.columns:
                    if " -> " in col_name:
                        kategoria, typ = col_name.split(" -> ")
                        
                        # Jeśli użytkownik wybrał konkretną kategorię, wyróżnijmy ją wizualnie
                        jest_wybrana = (kategoria == wybrana_kategoria)
                        
                        wartosc = row[col_name]
                        if wartosc != "Brak danych" and str(wartosc).strip() != "":
                            # Dodatkowy efekt wizualny dla poszukiwanej kategorii
                            prefix = "⭐ " if jest_wybrana else ""
                            
                            if typ.lower() == "cytat":
                                st.info(f"**{prefix}[{kategoria.upper()}] - Cytat:**\n\n{wartosc}")
                            elif typ.lower() == "opis":
                                st.success(f"**{prefix}[{kategoria.upper()}] - Opis / Analiza:**\n\n{wartosc}")
                            elif 'etykieta' in typ.lower():
                                st.warning(f"**{prefix}🏷️ Etykieta ({kategoria}):** {wartosc}")
                                
    except Exception as e:
        st.error(f"Wystąpił błąd podczas analizy struktury pliku: {e}")
else:
    st.info("👆 Czekam na wgranie pliku .xlsx z analizą dyskursu.")