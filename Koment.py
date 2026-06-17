# Import sys jest potrzebny do uruchomienia i poprawnego zamknięcia aplikacji PyQt
import sys

# Pandas służy tutaj do wczytywania danych z pliku Excel
import pandas as pd

# Importujemy elementy GUI z biblioteki PyQt5
from PyQt5.QtWidgets import (
    QApplication,        # główny obiekt aplikacji
    QMainWindow,         # główne okno programu
    QWidget,             # kontener na inne elementy GUI
    QVBoxLayout,         # układ pionowy
    QHBoxLayout,         # układ poziomy
    QPushButton,         # przycisk
    QTableWidget,        # tabela
    QTableWidgetItem,    # pojedyncza komórka tabeli
    QFileDialog,         # okno wyboru pliku
    QMessageBox,         # okno komunikatu/błędu
    QLabel,              # zwykły napis
    QLineEdit,           # pole tekstowe
    QComboBox,           # lista rozwijana
    QDoubleSpinBox       # pole liczbowe z wartościami po przecinku
)

# Element potrzebny do osadzenia wykresu Matplotlib w oknie PyQt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# Figure to główny obiekt wykresu Matplotlib
from matplotlib.figure import Figure


# Klasa reprezentująca pojedynczy wydatek
class Transaction:
    # Konstruktor uruchamia się przy tworzeniu obiektu Transaction(...)
    def __init__(self, data, kategoria, kwota, opis):
        # Zapisujemy datę wydatku w obiekcie
        self.data = data

        # Zapisujemy kategorię, np. Jedzenie, Transport, Rachunki
        self.kategoria = kategoria

        # Zapisujemy kwotę jako float, czyli liczbę zmiennoprzecinkową
        self.kwota = float(kwota)

        # Zapisujemy opis wydatku
        self.opis = opis


# Klasa zarządzająca wydatkami
# Oddziela logikę danych od wyglądu aplikacji
class BudgetManager:
    # Konstruktor managera
    def __init__(self):
        # Lista, w której będą przechowywane obiekty klasy Transaction
        self.wydatki = []

    # Metoda dodająca jeden wydatek do listy
    def dodaj_wydatek(self, wydatek):
        self.wydatki.append(wydatek)

    # Metoda czyszcząca wszystkie dane
    def czysc_dane(self):
        self.wydatki.clear()

    # Metoda licząca sumę wydatków dla każdej kategorii
    def oblicz_sume_kategorii(self):
        # Tworzymy pusty słownik
        # Przykładowy wynik: {"Jedzenie": 120.50, "Transport": 80.00}
        suma = {}

        # Przechodzimy po wszystkich wydatkach zapisanych w managerze
        for wydatek in self.wydatki:
            # Jeśli danej kategorii nie ma jeszcze w słowniku, tworzymy ją z wartością 0
            if wydatek.kategoria not in suma:
                suma[wydatek.kategoria] = 0

            # Dodajemy kwotę wydatku do odpowiedniej kategorii
            suma[wydatek.kategoria] += wydatek.kwota

        # Zwracamy gotowy słownik z sumami
        return suma

    # Metoda wczytująca dane z Excela
    def wczytaj_z_excela(self, sciezka_pliku):
        # Przed wczytaniem nowego pliku czyścimy stare dane
        self.czysc_dane()

        # Wczytujemy plik Excel do obiektu DataFrame
        df = pd.read_excel(sciezka_pliku)

        # Lista kolumn, które muszą istnieć w Excelu
        wymagane_kolumny = ["Data", "Kategoria", "Kwota", "Opis"]

        # Sprawdzamy, czy każda wymagana kolumna znajduje się w pliku
        for kolumna in wymagane_kolumny:
            if kolumna not in df.columns:
                # Jeśli brakuje kolumny, przerywamy działanie i zgłaszamy błąd
                raise ValueError(f"Brakuje kolumny: {kolumna}")

        # Przechodzimy po każdym wierszu Excela
        for _, row in df.iterrows():
            # Z jednego wiersza Excela tworzymy jeden obiekt Transaction
            wydatek = Transaction(
                row["Data"],
                row["Kategoria"],
                row["Kwota"],
                row["Opis"]
            )

            # Dodajemy utworzony wydatek do managera
            self.dodaj_wydatek(wydatek)


# Klasa odpowiedzialna za wykres
# Dziedziczy po FigureCanvas, aby można było wstawić wykres do okna PyQt
class ChartCanvas(FigureCanvas):
    def __init__(self):
        # Tworzymy figurę Matplotlib o rozmiarze 5x4 cale
        self.figure = Figure(figsize=(5, 4))

        # Dodajemy obszar wykresu
        self.ax = self.figure.add_subplot(111)

        # Uruchamiamy konstruktor klasy FigureCanvas
        super().__init__(self.figure)

    # Metoda rysująca wykres na podstawie danych
    def rysuj_wykres(self, dane):
        # Czyścimy poprzedni wykres
        self.ax.clear()

        # Jeśli nie ma danych, pokazujemy informację
        if not dane:
            self.ax.set_title("Brak danych do wyświetlenia")
            self.draw()
            return

        # Pobieramy nazwy kategorii ze słownika
        kategorie = list(dane.keys())

        # Pobieramy wartości, czyli sumy kwot
        kwoty = list(dane.values())

        # Tworzymy wykres słupkowy
        self.ax.bar(kategorie, kwoty)

        # Ustawiamy tytuł wykresu
        self.ax.set_title("Wydatki według kategorii")

        # Opis osi X
        self.ax.set_xlabel("Kategoria")

        # Opis osi Y
        self.ax.set_ylabel("Kwota [zł]")

        # Obracamy podpisy kategorii, żeby były czytelniejsze
        self.ax.tick_params(axis="x", rotation=30)

        # Dopasowujemy układ wykresu
        self.figure.tight_layout()

        # Odświeżamy wykres w oknie
        self.draw()


# Główna klasa okna aplikacji
class MainWindow(QMainWindow):
    def __init__(self):
        # Uruchamiamy konstruktor klasy QMainWindow
        super().__init__()

        # Ustawiamy tytuł okna
        self.setWindowTitle("Domowy budżet")

        # Ustawiamy pozycję i rozmiar okna
        # setGeometry(x, y, szerokość, wysokość)
        self.setGeometry(100, 100, 900, 600)

        # Tworzymy manager danych
        self.manager = BudgetManager()

        # Tworzymy przycisk do wczytywania Excela
        self.button_wczytaj = QPushButton("Wczytaj plik Excel")

        # Po kliknięciu przycisku uruchomi się metoda wybierz_plik
        self.button_wczytaj.clicked.connect(self.wybierz_plik)

        # Tworzymy przycisk do czyszczenia danych
        self.button_wyczysc = QPushButton("Wyczyść dane")

        # Po kliknięciu uruchomi się metoda wyczysc_dane
        self.button_wyczysc.clicked.connect(self.wyczysc_dane)

        # Tworzymy tabelę
        self.table = QTableWidget()

        # Ustawiamy liczbę kolumn w tabeli
        self.table.setColumnCount(4)

        # Ustawiamy nazwy kolumn
        self.table.setHorizontalHeaderLabels(["Data", "Kategoria", "Kwota", "Opis"])

        # Tworzymy obiekt wykresu
        self.chart = ChartCanvas()

        # Pole tekstowe do wpisania daty
        self.input_data = QLineEdit()
        self.input_data.setPlaceholderText("Data, np. 2026-06-14")

        # Lista rozwijana z kategoriami
        self.input_kategoria = QComboBox()
        self.input_kategoria.addItems([
            "Jedzenie",
            "Rozrywka",
            "Rachunki",
            "Transport",
            "Inne"
        ])

        # Pole liczbowe do wpisania kwoty
        self.input_kwota = QDoubleSpinBox()
        self.input_kwota.setRange(0, 100000)
        self.input_kwota.setDecimals(2)
        self.input_kwota.setSuffix(" zł")

        # Pole tekstowe do wpisania opisu
        self.input_opis = QLineEdit()
        self.input_opis.setPlaceholderText("Opis wydatku")

        # Przycisk do ręcznego dodania wydatku
        self.button_dodaj = QPushButton("Dodaj wydatek ręcznie")
        self.button_dodaj.clicked.connect(self.dodaj_recznie)

        # Główny układ pionowy całego okna
        layout_glowny = QVBoxLayout()

        # Układ poziomy dla przycisków
        layout_przyciski = QHBoxLayout()
        layout_przyciski.addWidget(self.button_wczytaj)
        layout_przyciski.addWidget(self.button_wyczysc)

        # Układ poziomy dla formularza ręcznego dodawania wydatków
        layout_formularz = QHBoxLayout()
        layout_formularz.addWidget(QLabel("Data:"))
        layout_formularz.addWidget(self.input_data)
        layout_formularz.addWidget(QLabel("Kategoria:"))
        layout_formularz.addWidget(self.input_kategoria)
        layout_formularz.addWidget(QLabel("Kwota:"))
        layout_formularz.addWidget(self.input_kwota)
        layout_formularz.addWidget(QLabel("Opis:"))
        layout_formularz.addWidget(self.input_opis)
        layout_formularz.addWidget(self.button_dodaj)

        # Dodajemy układ przycisków do głównego układu
        layout_glowny.addLayout(layout_przyciski)

        # Dodajemy formularz do głównego układu
        layout_glowny.addLayout(layout_formularz)

        # Dodajemy tabelę do głównego układu
        layout_glowny.addWidget(self.table)

        # Dodajemy wykres do głównego układu
        layout_glowny.addWidget(self.chart)

        # Tworzymy kontener QWidget
        # QMainWindow nie przyjmuje bezpośrednio layoutu, więc potrzebny jest QWidget
        container = QWidget()

        # Ustawiamy główny layout na kontenerze
        container.setLayout(layout_glowny)

        # Ustawiamy kontener jako główną zawartość okna
        self.setCentralWidget(container)

    # Metoda uruchamiana po kliknięciu przycisku "Wczytaj plik Excel"
    def wybierz_plik(self):
        # Otwieramy okno wyboru pliku
        sciezka, _ = QFileDialog.getOpenFileName(
            self,
            "Wybierz plik Excel",
            "",
            "Excel Files (*.xlsx)"
        )

        # Jeśli użytkownik wybrał plik
        if sciezka:
            try:
                # Wczytujemy dane z Excela do managera
                self.manager.wczytaj_z_excela(sciezka)

                # Odświeżamy tabelę
                self.odswiez_tabele()

                # Odświeżamy wykres
                self.odswiez_wykres()

            # Jeśli wystąpi błąd, pokazujemy okno z komunikatem
            except Exception as e:
                QMessageBox.critical(self, "Błąd", str(e))

    # Metoda odświeżająca tabelę na podstawie danych z managera
    def odswiez_tabele(self):
        # Ustawiamy liczbę wierszy równą liczbie wydatków
        self.table.setRowCount(len(self.manager.wydatki))

        # enumerate daje numer wiersza oraz obiekt wydatku
        for row, wydatek in enumerate(self.manager.wydatki):
            # Wstawiamy datę do kolumny 0
            self.table.setItem(row, 0, QTableWidgetItem(str(wydatek.data)))

            # Wstawiamy kategorię do kolumny 1
            self.table.setItem(row, 1, QTableWidgetItem(str(wydatek.kategoria)))

            # Wstawiamy kwotę do kolumny 2
            self.table.setItem(row, 2, QTableWidgetItem(str(wydatek.kwota)))

            # Wstawiamy opis do kolumny 3
            self.table.setItem(row, 3, QTableWidgetItem(str(wydatek.opis)))

    # Metoda odświeżająca wykres
    def odswiez_wykres(self):
        # Pobieramy słownik z sumami kategorii
        dane = self.manager.oblicz_sume_kategorii()

        # Przekazujemy dane do klasy wykresu
        self.chart.rysuj_wykres(dane)

    # Metoda dodająca wydatek wpisany ręcznie w formularzu
    def dodaj_recznie(self):
        # Pobieramy tekst z pola daty
        data = self.input_data.text()

        # Pobieramy wybraną kategorię z listy rozwijanej
        kategoria = self.input_kategoria.currentText()

        # Pobieramy wartość z pola liczbowego
        kwota = self.input_kwota.value()

        # Pobieramy opis z pola tekstowego
        opis = self.input_opis.text()

        # Sprawdzamy, czy data i opis nie są puste
        if data == "" or opis == "":
            QMessageBox.warning(self, "Błąd", "Uzupełnij datę i opis.")
            return

        # Tworzymy nowy obiekt Transaction z danych wpisanych przez użytkownika
        wydatek = Transaction(data, kategoria, kwota, opis)

        # Dodajemy wydatek do managera
        self.manager.dodaj_wydatek(wydatek)

        # Odświeżamy tabelę
        self.odswiez_tabele()

        # Odświeżamy wykres
        self.odswiez_wykres()

        # Czyścimy pole daty
        self.input_data.clear()

        # Zerujemy kwotę
        self.input_kwota.setValue(0)

        # Czyścimy opis
        self.input_opis.clear()

    # Metoda czyszcząca wszystkie dane
    def wyczysc_dane(self):
        # Czyścimy dane w managerze
        self.manager.czysc_dane()

        # Odświeżamy tabelę, żeby była pusta
        self.odswiez_tabele()

        # Odświeżamy wykres, żeby pokazał brak danych
        self.odswiez_wykres()


# Ten warunek oznacza:
# uruchom poniższy kod tylko wtedy, gdy ten plik jest uruchamiany bezpośrednio
if __name__ == "__main__":
    # Tworzymy aplikację PyQt
    app = QApplication(sys.argv)

    # Tworzymy główne okno
    window = MainWindow()

    # Pokazujemy okno
    window.show()

    # Uruchamiamy aplikację i poprawnie ją zamykamy po wyjściu
    sys.exit(app.exec_())
