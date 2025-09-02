import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QFileDialog,
    QTableWidget, QTableWidgetItem, QMessageBox
)
from analysis import check_normality
from export_word import export_results

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD - Статистичний аналіз даних")
        self.resize(900, 600)

        # Таблиця
        self.table = QTableWidget(5, 5)  # стартова таблиця 5x5
        self.table.setHorizontalHeaderLabels([f"Фактор {i+1}" for i in range(5)])
        self.setCentralWidget(self.table)

        # Меню
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Файл")
        analysis_menu = menubar.addMenu("Аналіз даних")
        about_menu = menubar.addMenu("Про програму")

        # Кнопки
        exit_action = QAction("Вихід", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        normality_action = QAction("Перевірка на нормальність (Шапіро–Вілка)", self)
        normality_action.triggered.connect(self.run_normality_test)
        analysis_menu.addAction(normality_action)

        about_action = QAction("Розробник", self)
        about_action.triggered.connect(self.show_about)
        about_menu.addAction(about_action)

        # Додаємо підтримку Ctrl+C, Ctrl+V
        self.table.installEventFilter(self)

    def eventFilter(self, source, event):
        if source == self.table:
            if event.type() == event.KeyPress:
                if event.matches(event.Copy):
                    self.copy_selection()
                    return True
                elif event.matches(event.Paste):
                    self.paste_selection()
                    return True
        return super().eventFilter(source, event)

    def copy_selection(self):
        selected = self.table.selectedRanges()
        if selected:
            r = selected[0]
            text = ""
            for i in range(r.rowCount()):
                row_data = []
                for j in range(r.columnCount()):
                    item = self.table.item(r.topRow() + i, r.leftColumn() + j)
                    row_data.append(item.text() if item else "")
                text += "\t".join(row_data) + "\n"
            QApplication.clipboard().setText(text)

    def paste_selection(self):
        text = QApplication.clipboard().text()
        rows = text.splitlines()
        r = self.table.currentRow()
        c = self.table.currentColumn()
        for i, row in enumerate(rows):
            for j, val in enumerate(row.split("\t")):
                item = QTableWidgetItem(val)
                self.table.setItem(r + i, c + j, item)

    def run_normality_test(self):
        data = []
        for col in range(self.table.columnCount()):
            col_data = []
            for row in range(self.table.rowCount()):
                item = self.table.item(row, col)
                if item and item.text().strip():
                    try:
                        col_data.append(float(item.text()))
                    except ValueError:
                        pass
            if col_data:
                data.append(col_data)

        if not data:
            QMessageBox.warning(self, "Помилка", "Немає даних для аналізу.")
            return

        results = check_normality(data)
        export_results(data, results)
        QMessageBox.information(self, "Готово", "Результати збережено у Word-документі.")

    def show_about(self):
        QMessageBox.information(self, "Про програму",
            "SAD - Статистичний аналіз даних\n"
            "Розробник: Чаплоуцький А.М.\n"
            "Кафедра плодівництва і виноградарства УНУ\n"
            "© Усі права захищено"
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
