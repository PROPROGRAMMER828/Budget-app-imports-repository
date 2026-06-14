import sys
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QFileDialog,
    QMessageBox,
    QLabel,
    QLineEdit,
    QComboBox,
    QDoubleSpinBox
)

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


class Transaction:
    def __init__(self, data, kategoria, kwota, opis):
        self.data = data
        self.kategoria = kategoria
        self.kwota = float(kwota)
        self.opis = opis


class BudgetManager:
    def __init__(self):
        self.wydatki = []

    def dodaj_wydatek(self, wydatek):
        self.wydatki.append(wydatek)

    def czysc_dane(self):
        self.wydatki.clear()

    def oblicz_sume_kategorii(self):
        suma = {}

        for wydatek in self.wydatki:
            if wydatek.kategoria not in suma:
                suma[wydatek.kategoria] = 0

            suma[wydatek.kategoria] += wydatek.kwota

        return suma

    def wczytaj_z_excela(self, sciezka_pliku):
        self.czysc_dane()

        df = pd.read_excel(sciezka_pliku)

        wymagane_kolumny = ["Data", "Kategoria", "Kwota", "Opis"]

        for kolumna in wymagane_kolumny:
            if kolumna not in df.columns:
                raise ValueError(f"Brakuje kolumny: {kolumna}")

        for _, row in df.iterrows():
            wydatek = Transaction(
                row["Data"],
                row["Kategoria"],
                row["Kwota"],
                row["Opis"]
            )

            self.dodaj_wydatek(wydatek)


class ChartCanvas(FigureCanvas):
    def __init__(self):
        self.figure = Figure(figsize=(5, 4))
        self.ax = self.figure.add_subplot(111)
        super().__init__(self.figure)

    def rysuj_wykres(self, dane):
        self.ax.clear()

        if not dane:
            self.ax.set_title("Brak danych do wyświetlenia")
            self.draw()
            return

        kategorie = list(dane.keys())
        kwoty = list(dane.values())

        self.ax.bar(kategorie, kwoty)
        self.ax.set_title("Wydatki według kategorii")
        self.ax.set_xlabel("Kategoria")
        self.ax.set_ylabel("Kwota [zł]")
        self.ax.tick_params(axis="x", rotation=30)

        self.figure.tight_layout()
        self.draw()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Domowy budżet")
        self.setGeometry(100, 100, 900, 600)

        self.manager = BudgetManager()

        self.button_wczytaj = QPushButton("Wczytaj plik Excel")
        self.button_wczytaj.clicked.connect(self.wybierz_plik)

        self.button_wyczysc = QPushButton("Wyczyść dane")
        self.button_wyczysc.clicked.connect(self.wyczysc_dane)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Data", "Kategoria", "Kwota", "Opis"])

        self.chart = ChartCanvas()

        self.input_data = QLineEdit()
        self.input_data.setPlaceholderText("Data, np. 2026-06-14")

        self.input_kategoria = QComboBox()
        self.input_kategoria.addItems([
            "Jedzenie",
            "Rozrywka",
            "Rachunki",
            "Transport",
            "Inne"
        ])

        self.input_kwota = QDoubleSpinBox()
        self.input_kwota.setRange(0, 100000)
        self.input_kwota.setDecimals(2)
        self.input_kwota.setSuffix(" zł")

        self.input_opis = QLineEdit()
        self.input_opis.setPlaceholderText("Opis wydatku")

        self.button_dodaj = QPushButton("Dodaj wydatek ręcznie")
        self.button_dodaj.clicked.connect(self.dodaj_recznie)

        layout_glowny = QVBoxLayout()

        layout_przyciski = QHBoxLayout()
        layout_przyciski.addWidget(self.button_wczytaj)
        layout_przyciski.addWidget(self.button_wyczysc)

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

        layout_glowny.addLayout(layout_przyciski)
        layout_glowny.addLayout(layout_formularz)
        layout_glowny.addWidget(self.table)
        layout_glowny.addWidget(self.chart)

        container = QWidget()
        container.setLayout(layout_glowny)

        self.setCentralWidget(container)

    def wybierz_plik(self):
        sciezka, _ = QFileDialog.getOpenFileName(
            self,
            "Wybierz plik Excel",
            "",
            "Excel Files (*.xlsx)"
        )

        if sciezka:
            try:
                self.manager.wczytaj_z_excela(sciezka)
                self.odswiez_tabele()
                self.odswiez_wykres()

            except Exception as e:
                QMessageBox.critical(self, "Błąd", str(e))

    def odswiez_tabele(self):
        self.table.setRowCount(len(self.manager.wydatki))

        for row, wydatek in enumerate(self.manager.wydatki):
            self.table.setItem(row, 0, QTableWidgetItem(str(wydatek.data)))
            self.table.setItem(row, 1, QTableWidgetItem(str(wydatek.kategoria)))
            self.table.setItem(row, 2, QTableWidgetItem(str(wydatek.kwota)))
            self.table.setItem(row, 3, QTableWidgetItem(str(wydatek.opis)))

    def odswiez_wykres(self):
        dane = self.manager.oblicz_sume_kategorii()
        self.chart.rysuj_wykres(dane)

    def dodaj_recznie(self):
        data = self.input_data.text()
        kategoria = self.input_kategoria.currentText()
        kwota = self.input_kwota.value()
        opis = self.input_opis.text()

        if data == "" or opis == "":
            QMessageBox.warning(self, "Błąd", "Uzupełnij datę i opis.")
            return

        wydatek = Transaction(data, kategoria, kwota, opis)

        self.manager.dodaj_wydatek(wydatek)

        self.odswiez_tabele()
        self.odswiez_wykres()

        self.input_data.clear()
        self.input_kwota.setValue(0)
        self.input_opis.clear()

    def wyczysc_dane(self):
        self.manager.czysc_dane()
        self.odswiez_tabele()
        self.odswiez_wykres()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())

