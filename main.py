import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QTableWidget,
    QTableWidgetItem, QInputDialog, QMessageBox
)
from PyQt5.QtCore import Qt
from analysis import check_normality
from export_word import export_results_to_word


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD - Статистичний аналіз даних")
        self.resize(1000, 600)

        # Початкова таблиця 5x5
        self.table = QTableWidget(5, 5)
        self.table.setHorizontalHeaderLabels([f"Фактор {i+1}" for i in range(5)])
        self.setCentralWidget(self.table)

        # Робимо всі клітинки редагованими
        for i in range(self.table.rowCount()):
            for j in range(self.table.columnCount()):
                item = QTableWidgetItem("")
                item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.table.setItem(i, j, item)

        # Меню
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Файл")
        table_menu = menubar.addMenu("Таблиця")
        analysis_menu = menubar.addMenu("Аналіз даних")
        help_menu = menubar.addMenu("Про програму")

        # Файл → Вихід
        exit_action = QAction("Вихід", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Таблиця → додати / видалити
        add_row = QAction("Додати рядок", self)
        add_row.triggered.connect(self.add_row)
        table_menu.addAction(add_row)

        del_row = QAction("Видалити рядок", self)
        del_row.triggered.connect(self.delete_row)
        table_menu.addAction(del_row)

        add_col = QAction("Додати стовпчик", self)
        add_col.triggered.connect(self.add_col)
        table_menu.addAction(add_col)

        del_col = QAction("Видалити стовпчик", self)
        del_col.triggered.connect(self.delete_col)
        table_menu.addAction(del_col)

        # Аналіз → Перевірка нормальності
        norm_action = QAction("Перевірити нормальність (Шапіро-Вілк)", self)
        norm_action.triggered.connect(self.run_normality_test)
        analysis_menu.addAction(norm_action)

        # Про програму
        about_action = QAction("Інформація", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        # Додаємо підтримку Ctrl+C / Ctrl+V
        self.table.installEventFilter(self)

    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for j in range(self.table.columnCount()):
            item = QTableWidgetItem("")
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.table.setItem(row, j, item)

    def delete_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def add_col(self):
        col = self.table.columnCount()
        self.table.insertColumn(col)
        self.table.setHorizontalHeaderItem(col, QTableWidgetItem(f"Фактор {col+1}"))
        for i in range(self.table.rowCount()):
            item = QTableWidgetItem("")
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.table.setItem(i, col, item)

    def delete_col(self):
        col = self.table.currentColumn()
        if col >= 0:
            self.table.removeColumn(col)

    # Ctrl+C / Ctrl+V
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
                item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.table.setItem(r + i, c + j, item)

    def run_normality_test(self):
        indicator, ok = QInputDialog.getText(self, "Назва показника", "Введіть назву показника:")
        units, ok2 = QInputDialog.getText(self, "Одиниці виміру", "Введіть одиниці виміру:")

        if not ok or not ok2:
            return

        # Збираємо дані
        data = []
        for i in range(self.table.rowCount()):
            for j in range(self.table.columnCount()):
                item = self.table.item(i, j)
                if item and item.text().strip():
                    try:
                        data.append(float(item.text()))
                    except ValueError:
                        pass

        if not data:
            QMessageBox.warning(self, "Помилка", "Дані відсутні або некоректні")
            return

        result = check_normality(data)

        # Показуємо результат
        QMessageBox.information(self, "Результат перевірки",
                                f"Статистика W = {result['W']:.4f}, p = {result['p']:.4f}")

        # Експортуємо у Word
        export_results_to_word(indicator, units, data, result)

    def show_about(self):
        QMessageBox.information(self, "Про програму",
                                "SAD - Статистичний аналіз даних\n"
                                "Розробник: ....\n"
                                "Кафедра плодівництва і виноградарства УНУ")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
