INSTALACJA:
1. Pobieramy Pythona 3.6.4 ze strony https://www.python.org/downloads/
2. Plik pinger.py wrzucamy do folderu, w którym znajduje lub będzie się znajdował plik z pingami weekendowymi.
3. Opdalamy powershella
4. za pomocą komend "cd" przechodzimy do danego folderu.
5. Instalujemy pakiet openpyxl komendą "pip install openpyxl"

INSTRUKCJA
1. skrypt uruchamiamy komendą "python pinger.py"
2. skrypt przegląda plik pingów, jeśli jest sformatowany ok, to wyświetli nam treści wszystkich maili wraz z adresatem.
3. Wpisując "T" akceptujemy wysyłke pingu, "N" wyłacza skrypt
4. Bedziemy poproszeni o podanie swojego adresu mailowego i hasła do SMTP (do wyciągniecia z thunderbirda -> opcje -> zabezpieczenia)
5. Przed zamknięciem programu powinien wyświetlić się napis "Rozesłano wszystkie pingi.", co oznacza że wszystko poszło ok.

HOWTO:
Generalnie skrypcik ma pewną elastyczność jeśli chodzi o interpretowanie formatu pliku, jednak dla bezpieczeństwa najlepiej podawać mu pliki pingów weekendowych sformatowanych zgodnie z plikiem example.xlsx

FAQ:
1. Wpisując hasło, nie widzę, żebym cokolwiek pisał.
odp: Jest to spowodowane zastosowaniem modułu getpass, który m. in. służy do zamaskowania wpisywanego hasła, dlatego go nie widać.
2. Program zapętlił się, help!
Generalnie w skrajnie nieprawdopodobnym przypadku w którym program non stop coś pokazuje, by przerwać jego działanie, użyj ctrl+c
3. 