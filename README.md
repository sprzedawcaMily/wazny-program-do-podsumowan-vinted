Program: eksport maili Vinted z WP do Excela

Co robi:
- Loguje się do skrzynki WP przez IMAP.
- Szuka maili od Zespol Vinted.
- Bierze tylko dwa typy tematow:
  - Zamowienie zostalo zakonczone (sprzedaze)
  - Twoje potwierdzenie zakupu (zakupy)
- Wyciaga dane: tytul, kwota, data, blok kraju/firmy.
- Automatycznie tlumaczy tytul na polski.
- Zapisuje do Excela w 2 arkuszach: Sprzedaze i Zakupy.
- Przy kolejnym uruchomieniu dopisuje tylko nowe rekordy (bez duplikatow).
 - Tworzy tez arkusz Uslugi elektroniczne i Podsumowanie.

Wymagania:
- Python 3.10+
- Konto WP z dostepem IMAP
- (opcjonalnie) Node.js 18+ dla scrapera Puppeteer

Pliki:
- vinted_mail_to_excel.py
- requirements.txt
- .env.example

Konfiguracja:
1. Skopiuj .env.example do .env
2. Ustaw w .env:
   - WP_EMAIL
   - WP_PASSWORD
   - START_DATE (domyslnie 2026-03-16)
   - OUTPUT_XLSX (domyslnie vinted_podsumowanie.xlsx)
  - OUTPUT_HTML (domyslnie vinted_podsumowanie.html)
  - DYNAMIC_FILENAME=1 (domyslnie: dynamiczne nazwy plikow)
  - OUTPUT_PREFIX (np. podsumowanie Kamochi)
  - (opcjonalnie) VINTED_PROFILE_DIR, VINTED_CHROME_PATH, VINTED_WAIT_FOR_LOGIN, VINTED_HEADLESS

Uwaga do hasla:
- Jesli zwykle haslo nie dziala, ustaw haslo aplikacji w WP i uzyj go w WP_PASSWORD.

Instalacja bibliotek:
pip install -r requirements.txt

Opcjonalny scraper (kraje kupujacych z Vinted):
1. Zainstaluj zaleznosci Node:
  npm install
2. Przy uruchomieniu programu zaakceptuj pytanie o scraping.
  Otworzy sie przegladarka, zaloguj sie do Vinted.
  Skrypt przejdzie do zakladki Zakonczone i pobierze kraj kupujacego.

Jesli logowanie przez Google jest blokowane:
- Zaloguj sie na Vinted loginem/haslem (nie przez Google), albo
- Uzyj profilu Chrome z zapamietana sesja i ustaw w .env:
  - VINTED_PROFILE_DIR=C:\Users\<twoj_user>\AppData\Local\Google\Chrome\User Data\Profile 1
  - VINTED_CHROME_PATH=C:\Program Files\Google\Chrome\Application\chrome.exe
  - VINTED_WAIT_FOR_LOGIN=0 (gdy sesja juz jest aktywna)

Uruchomienie:
python vinted_mail_to_excel.py

Wynik:
- Powstaje lub aktualizuje sie plik Excel:
  - Arkusz Zakupy (pierwszy)
  - Arkusz Sprzedaze
  - Arkusz Uslugi elektroniczne
  - Arkusz Podsumowanie (sumy)

Pola w tabelach:
- Zakupy:
  - data
  - tytul_oryginal
  - tytul_pl
  - kwota_lacznie
- Sprzedaze:
  - data
  - tytul_oryginal
  - kwota
  - nr_vat
  - numer_transakcji
  - kraj_kupujacego
  - wysylka_zagraniczna
- Uslugi elektroniczne:
  - data
  - usluga
  - kwota
- Podsumowanie:
  - Suma sprzedazy
  - Suma sprzedazy zagraniczna
  - Suma zakupow
  - Suma uslug elektronicznych

Filtrowanie daty:
- Program bierze maile od daty START_DATE wzwyz.
- Przy starcie mozesz podac miesiac w formacie YYYY-MM, wtedy bierze tylko ten miesiac.

Dynamiczne nazwy plikow:
- Domyslnie, gdy OUTPUT_XLSX i OUTPUT_HTML maja wartosci "vinted_podsumowanie.xlsx/html", pliki beda dynamiczne (np. "podsumowanie Kamochi 2026-03.xlsx").
- Mozesz wymusic to przez DYNAMIC_FILENAME=1 i OUTPUT_PREFIX.
- Alternatywnie uzyj placeholderow w OUTPUT_XLSX/OUTPUT_HTML: {month} lub {period}.
