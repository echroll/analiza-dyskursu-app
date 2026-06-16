import streamlit as st
import pandas as pd
import os
import io
from docx import Document
from docx.shared import Pt, Inches
import matplotlib.pyplot as plt
from wordcloud import WordCloud, STOPWORDS
from transformers import pipeline

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
        
        typ_dok = row.get(next((c for c in dataframe.columns if 'typ dokumentu' in str(c).lower()), '-'), '-')
        odbiorcy = row.get(next((c for c in dataframe.columns if 'odbiorcy' in str(c).lower()), '-'), '-')
        filozof = row.get(next((c for c in dataframe.columns if 'filozof' in str(c).lower()), '-'), '-')
        
        p_meta = doc.add_paragraph()
        p_meta.add_run(f"Autor: {autor} | Rok: {rok} | Źródło: {zrodlo}\n").italic = True
        p_meta.add_run(f"Typ dokumentu: {typ_dok} | Odbiorcy: {odbiorcy} | Filozofowie: {filozof}").italic = True
        
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
        # Naprawa skomplikowanych nagłówków w tabeli
        nowe_kolumny = []
        ostatnia_kategoria = "" 
        
        for col in df.columns:
            kat_glowna = str(col[0]).replace('\n', ' ').strip()
            podkategoria = str(col[1]).replace('\n', ' ').strip()
            
            # Wzmocnione wykrywanie pustych komórek w nagłówkach
            if "unnamed" not in kat_glowna.lower() and kat_glowna != "" and kat_glowna.lower() != "nan":
                ostatnia_kategoria = kat_glowna
                
            if "unnamed" in podkategoria.lower() or podkategoria == "" or podkategoria.lower() == "nan":
                nowe_kolumny.append(kat_glowna)
            else:
                nowe_kolumny.append(f"{ostatnia_kategoria} -> {podkategoria}")
                
        df.columns = nowe_kolumny
        df = df.fillna("Brak danych")

        # Bezpieczne szukanie kolumn meta-danych
        kol_autor = next((col for col in df.columns if 'autor' in str(col).lower()), None)
        kol_rok = next((col for col in df.columns if 'rok' in str(col).lower()), None)

        # --- ŁATKA 1: NORMALIZACJA AUTORÓW I DAT ---
        import re
        def normalizuj_autora(a):
            a_low = str(a).lower()
            wynik = []
            if 'putin' in a_low: wynik.append('Władimir Putin')
            if 'kirill' in a_low or 'cyryl' in a_low: wynik.append('Patriarcha Cyryl')
            if 'lavrov' in a_low or 'ławrow' in a_low: wynik.append('Siergiej Ławrow')
            if 'medvedev' in a_low or 'miedwiediew' in a_low: wynik.append('Dmitrij Miedwiediew')
            if 'zakharova' in a_low or 'zacharowa' in a_low: wynik.append('Maria Zacharowa')
            if 'peskov' in a_low or 'pieskow' in a_low: wynik.append('Dmitrij Pieskow')
            if 'gorbaczow' in a_low or 'gorbachev' in a_low: wynik.append('Michaił Gorbaczow')
            
            if wynik: return ", ".join(wynik)
            return str(a).strip()
            
        def wyciagnij_rok(r):
            m = re.search(r'\d{4}', str(r))
            return m.group(0) if m else str(r)

        if kol_autor: df[kol_autor] = df[kol_autor].apply(normalizuj_autora)
        if kol_rok: df[kol_rok] = df[kol_rok].apply(wyciagnij_rok)

        # -------------------------------------------

        # Dynamiczne wyciągnięcie dostępnych kategorii tematycznych z nagłówków
        lista_kategorii = sorted(list(set([col.split(" -> ")[0] for col in df.columns if " -> " in col])))
        # --- PASEK BOCZNY (ZAAWANSOWANE FILTRY) ---
        st.sidebar.header("🔍 Filtry i Wyszukiwanie")

        # 1. Wyszukiwarka tekstowa (Wiele słów)
        szukana_fraza = st.sidebar.text_input("Szukaj słów (np. zachód wojna):")

        # 2. Szybkie przyciski i Filtr Autora
        st.sidebar.markdown("**🔥 Najczęstsi autorzy (Top 15):**")
        top_15 = df[kol_autor].value_counts().head(15).index.tolist()
        top_15 = [a for a in top_15 if a not in ["Brak danych", "-"]]
        
        # Tworzymy przyciski w kolumnach
        kols = st.sidebar.columns(3)
        for i, aut in enumerate(top_15):
            krotka_nazwa = aut.split()[-1] if ' ' in aut else aut
            if kols[i % 3].button(f"👤 {krotka_nazwa}", key=f"btn_{i}"):
                st.session_state['moj_autor'] = aut

        autorzy = ["Wszyscy"] + sorted([str(a) for a in df[kol_autor].unique() if str(a) != "Brak danych"])
        
        if 'moj_autor' not in st.session_state:
            st.session_state['moj_autor'] = "Wszyscy"

        wybrany_autor = st.sidebar.selectbox("Wybierz autora:", autorzy, key='moj_autor')
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

        # Filtrowanie wyszukiwarką tekstową (Wiele słów)
        if szukana_fraza:
            slowa_kluczowe = szukana_fraza.lower().split()
            # Musi zawierać KAŻDE wpisane słowo (w dowolnej z kolumn)
            tekstowy_warunek = df_filtered.apply(lambda row: all(any(slowo in str(x).lower() for x in row) for slowo in slowa_kluczowe), axis=1)
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

        # --- ZAKŁADKI Z WIZUALIZACJAMI (CHMURA I EMOCJE) ---
        st.write("---")
        st.header("🧠 Analiza NLP (Chmura i Afekt)")
        tab1, tab2 = st.tabs(["☁️ Chmura słów", "📈 Krzywa Afektu"])
        
        with tab1:
            st.markdown("**Wygeneruj chmurę słów z tekstów.**")
            
            # --- BEZPOŚREDNI IMPORT PLIKÓW TXT ZE STOP-LISTĄ ---
            wgrane_stop_pliki = st.file_uploader("📥 Wgraj swoje pliki stop-words (.txt) [Możesz wrzucić english, polish i russian naraz!]", type=["txt"], accept_multiple_files=True)
            
            stop_z_plikow = []
            if wgrane_stop_pliki:
                for plik_txt in wgrane_stop_pliki:
                    tresc = plik_txt.getvalue().decode("utf-8").splitlines()
                    stop_z_plikow.extend([linia.strip().lower() for linia in tresc if linia.strip()])
                st.success(f"✅ Wczytano {len(stop_z_plikow)} słów wykluczonych z Twoich plików!")
            # -------------------------------------------------------------
            
            col1, col2 = st.columns(2)
            zrodlo_danych = col1.radio("Wybierz zakres:", ["Tylko odfiltrowane rekordy", "Wszystkie dokumenty w bazie"])
            max_slow = col2.number_input("Maksymalna liczba słów:", min_value=10, max_value=500, value=100, step=10)
            
            tylko_cytaty_chmura = st.checkbox("🎯 Analizuj TYLKO cytaty (pomiń opisy i etykiety badaczy)", value=True)
            dodatkowe_stop = st.text_area("Możesz tutaj dopisać dodatkowe słowa wykluczone 'w locie' (oddzielone przecinkiem):", "putin,rosja")
            
            if st.button("Generuj chmurę słów"):
                import re # Narzędzie do czyszczenia tekstu
                df_do_chmury = df_filtered if zrodlo_danych == "Tylko odfiltrowane rekordy" else df
                
                wszystkie_teksty = []
                for _, r in df_do_chmury.iterrows():
                    for c in df_do_chmury.columns:
                        if "->" in c and str(r[c]) not in ["Brak danych", "", "nan"]:
                            kategoria, typ = c.split(" -> ")
                            
                            if tylko_cytaty_chmura and typ.lower() != "cytat":
                                continue
                                
                            wszystkie_teksty.append(str(r[c]))
                
                tekst_polaczony = " ".join(wszystkie_teksty)
                # KRYTYCZNE: Usuwamy znaki interpunkcyjne (kropki, przecinki, cudzysłowy), żeby stop-lista zadziałała!
                tekst_polaczony = re.sub(r'[^\w\s]', '', tekst_polaczony)
                
                if not tekst_polaczony.strip():
                    st.warning("Brak tekstów do wygenerowania chmury.")
                else:
                    stop_words = set(STOPWORDS)
                    moje_stop = [s.strip().lower() for s in dodatkowe_stop.split(',') if s.strip()]
                    stop_words.update(moje_stop)
                    stop_words.update(stop_z_plikow)
                    
                    wc = WordCloud(width=800, height=400, background_color='white', stopwords=stop_words, max_words=max_slow).generate(tekst_polaczony)
                    
                    fig, ax = plt.subplots(figsize=(10, 5))
                    ax.imshow(wc, interpolation='bilinear')
                    ax.axis("off")
                    st.pyplot(fig)
        with tab2:
            st.markdown("**Krzywa afektu (Affect Curve) - dynamika nastrojów w czasie.**")
            st.caption("*Analiza wspierana potężnym wielojęzycznym modelem AI (XLM-RoBERTa), który natywnie rozumie język rosyjski (cyrylicę), polski oraz angielski. Odczytuje ładunek emocjonalny od -1.0 (bardzo agresywny/negatywny) do +1.0 (bardzo pozytywny).*")
            
            # Zapamiętujemy model w pamięci podręcznej, żeby nie ładował się 5 minut przy każdym kliknięciu
            @st.cache_resource
            def wczytaj_model_nlp():
                return pipeline("sentiment-analysis", model="cardiffnlp/twitter-xlm-roberta-base-sentiment", truncation=True, max_length=512)
            
            if len(df_filtered) > 0 and st.button("Generuj krzywę afektu dla kategorii"):
                with st.spinner("Uruchamiam sieć neuronową i analizuję teksty... (to potrwa chwilę)"):
                    analizator = wczytaj_model_nlp()
                    dane_emocje = []
                    
                    for idx, r in df_filtered.iterrows():
                        rok = str(r.get(kol_rok, 'Brak'))
                        tytul = str(r.get(next((c for c in df.columns if 'tytuł' in str(c).lower()), 'Brak tytułu'), 'Brak tytułu'))
                        
                        if rok == 'Brak' or not rok.isdigit():
                            continue
                            
                        for c in df_filtered.columns:
                            if "->" in c and str(r[c]) not in ["Brak danych", "", "nan"]:
                                kategoria, typ = c.split(" -> ")
                                
                                if typ.lower() == "cytat":
                                    tekst_rekordu = str(r[c])
                                    
                                    # Magia AI: Model sam rozpoznaje język i ocenia tekst
                                    wynik = analizator(tekst_rekordu)[0]
                                    etykieta = wynik['label']
                                    pewnosc = wynik['score']
                                    
                                    # Model z Cardiff zwraca: LABEL_0 (negatywny), LABEL_1 (neutralny), LABEL_2 (pozytywny)
                                    if etykieta == 'LABEL_0' or 'negative' in etykieta.lower():
                                        sentyment = -pewnosc
                                    elif etykieta == 'LABEL_2' or 'positive' in etykieta.lower():
                                        sentyment = pewnosc
                                    else:
                                        sentyment = 0.0 # Neutralny
                                    
                                    dane_emocje.append({
                                        'Rok': int(rok), 
                                        'Kategoria': kategoria.upper(), 
                                        'Sentyment': sentyment,
                                        'Dokument': tytul
                                    })
                    
                    if dane_emocje:
                        df_emocje = pd.DataFrame(dane_emocje)
                        srednia_roczna = df_emocje.groupby(['Rok', 'Kategoria'])['Sentyment'].mean().reset_index()
                        wykres_data = srednia_roczna.pivot(index='Rok', columns='Kategoria', values='Sentyment')
                        
                        st.write("---")
                        st.markdown("**Wybierz kategorie do nałożenia na wykres:**")
                        
                        dostepne_kategorie = list(wykres_data.columns)
                        kolumny_chk = st.columns(4)
                        zaznaczone_kategorie = []
                        
                        for i, kat in enumerate(dostepne_kategorie):
                            if kolumny_chk[i % 4].checkbox(kat, value=True, key=f"chk_afekt_{kat}"):
                                zaznaczone_kategorie.append(kat)
                                
                        st.write("---")
                        
                        if zaznaczone_kategorie:
                            st.line_chart(wykres_data[zaznaczone_kategorie])
                            with st.expander("🔍 Zobacz szczegółowe wyniki dla poszczególnych dokumentów (tabela)"):
                                st.dataframe(df_emocje.sort_values(by=['Rok', 'Sentyment']), use_container_width=True)
                        else:
                            st.warning("Zaznacz przynajmniej jedną kategorię z listy powyżej, aby wygenerować wykres.")
                    else:
                        st.warning("W odfiltrowanych danych nie znaleziono żadnych cytatów dla wybranych osób i lat.")

        # --- WYŚWIETLANIE REKORDÓW ---
        st.write("---")

        for index, row in df_filtered.iterrows():
            tytul = row.get(next((c for c in df.columns if 'tytuł' in str(c).lower()), 'Brak tytułu'), 'Brak tytułu')
            autor = row.get(kol_autor, '-')
            rok = row.get(kol_rok, '-')
            zrodlo = row.get(next((c for c in df.columns if 'źródło' in str(c).lower() or 'zrodlo' in str(c).lower()), '-'), '-')
            
            # NOWOŚĆ: Używamy expandera, aby strona była czysta i zwarta
            with st.expander(f"📄 {tytul} ({autor} - {rok})"):
                # Wyciągamy brakujące metadane, ignorując spacje
                typ_dok = row.get(next((c for c in df.columns if 'typ dokumentu' in str(c).lower()), '-'), '-')
                odbiorcy = row.get(next((c for c in df.columns if 'odbiorcy' in str(c).lower()), '-'), '-')
                filozof = row.get(next((c for c in df.columns if 'filozof' in str(c).lower()), '-'), '-')
                
                st.markdown(f"**Źródło:** {zrodlo} | **Typ dokumentu:** {typ_dok}")
                st.markdown(f"**Odbiorcy:** {odbiorcy} | **Odniesienia do filozofów:** {filozof}")
                st.write("---")
                
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
