import pandas as pd
from functools import partial
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QShortcut, \
    QPushButton, QLabel
from PyQt5.QtGui import QKeySequence
from PyQt5.QtCore import Qt


class CustomTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

    def paste(self, cols_lim=int):
        clipboard_text = QApplication.clipboard().text()
        rows = clipboard_text.split('\n')

        # Remove qualquer linha vazia do final da lista
        while len(rows) > 0 and not rows[-1]:
            rows.pop()

        self.setRowCount(len(rows))
        self.setColumnCount(cols_lim)  # Número fixo de colunas (por exemplo, 3 colunas)

        for row_idx, row in enumerate(rows):
            cols = row.split('\t')[:cols_lim]  # Limitar o número de colunas ao valor fixo (por exemplo, 3 colunas)
            for col_idx, col in enumerate(cols):
                item = QTableWidgetItem(col)
                self.setItem(row_idx, col_idx, item)

        # Definir os cabeçalhos personalizados


class TableWithExport(QMainWindow):
    def __init__(self, rows, cols, headers, arq):
        super().__init__()
        self.rows = rows
        self.cols = cols
        self.headers = headers
        self.arq = arq
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Tabela com Exportação para Excel')
        self.setGeometry(0, 0, 900, 900)

        layout = QVBoxLayout()

        label = QLabel("* Obrigatório para o funcionamento correto.")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        self.table_widget = CustomTableWidget(self)
        self.table_widget.setRowCount(self.rows)
        self.table_widget.setColumnCount(self.cols)
        self.table_widget.setHorizontalHeaderLabels(self.headers)

        # Preenchendo a tabela com dados de exemplo
        for row in range(self.rows):
            for col in range(self.cols):
                item = QTableWidgetItem("")
                self.table_widget.setItem(row, col, item)

        layout.addWidget(self.table_widget)

        # Botão OK para exportar a tabela para Excel
        btn_export = QPushButton("OK", self)
        btn_export.clicked.connect(self.export_to_excel)
        layout.addWidget(btn_export)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Adicionando atalho para colar (Ctrl+V) na tabela
        shortcut_paste = QShortcut(QKeySequence.Paste, self.table_widget)
        shortcut_paste.activated.connect(partial(self.table_widget.paste, cols_lim=self.cols))

    def export_to_excel(self):
        rows = self.table_widget.rowCount()
        cols = self.table_widget.columnCount()
        data = []

        for row in range(rows):
            row_data = []
            for col in range(cols):
                item = self.table_widget.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    row_data.append('')
            data.append(row_data)

        df = pd.DataFrame(data, columns=self.headers)
        df.to_csv(self.arq, sep=',', index=False)
        self.close()