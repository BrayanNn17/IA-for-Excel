import sys
import os
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QComboBox, QLineEdit, QMessageBox)

class ExcelGraphApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Generador de Gráficos')
        self.setGeometry(100, 100, 400, 300)

        self.layout = QVBoxLayout()

        self.btn_abrir_archivo = QPushButton('Abrir Archivo Excel', self)
        self.btn_abrir_archivo.clicked.connect(self.abrir_archivo)
        self.layout.addWidget(self.btn_abrir_archivo)

        self.lbl_archivo = QLabel('No se ha seleccionado un archivo', self)
        self.layout.addWidget(self.lbl_archivo)

        self.lbl_filas = QLabel('Filas (separadas por comas):', self)
        self.layout.addWidget(self.lbl_filas)

        self.entry_filas = QLineEdit(self)
        self.layout.addWidget(self.entry_filas)

        self.lbl_columnas = QLabel('Columnas (separadas por comas):', self)
        self.layout.addWidget(self.lbl_columnas)

        self.entry_columnas = QLineEdit(self)
        self.layout.addWidget(self.entry_columnas)

        self.lbl_tipo_grafico = QLabel('Tipo de gráfico:', self)
        self.layout.addWidget(self.lbl_tipo_grafico)

        self.tipo_grafico = QComboBox(self)
        self.tipo_grafico.addItems(['linea', 'barra', 'scatter', 'pastel'])
        self.layout.addWidget(self.tipo_grafico)

        self.btn_crear_grafico = QPushButton('Crear Gráfico', self)
        self.btn_crear_grafico.clicked.connect(self.crear_grafico)
        self.layout.addWidget(self.btn_crear_grafico)

        self.setLayout(self.layout)

        self.file_path = None
        self.df = None

    def abrir_archivo(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel", "", "Archivos de Excel (*.xlsx *.xls);;Todos los archivos (*)", options=options)
        if file_path:
            self.file_path = file_path
            self.df = self.cargar_datos_desde_excel(file_path)
            if self.df is not None:
                self.lbl_archivo.setText(f"Archivo seleccionado: {os.path.basename(file_path)}")

    def cargar_datos_desde_excel(self, file_path):
        try:
            with xw.App(visible=False) as app:
                wb = xw.Book(file_path)
                sheet = wb.sheets[0]
                data = sheet.range('A1').expand().value
                df = pd.DataFrame(data[1:], columns=data[0])
                wb.close()
            return df
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar el archivo Excel: {e}")
            return None

    def crear_grafico(self):
        if not self.file_path:
            QMessageBox.warning(self, "Advertencia", "Primero seleccione un archivo Excel")
            return

        filas_seleccionadas = self.entry_filas.text()
        columnas_seleccionadas = self.entry_columnas.text()
        tipo_grafico = self.tipo_grafico.currentText()

        try:
            filas = [int(fila.strip()) - 1 for fila in filas_seleccionadas.split(",")]
            columnas = [int(col.strip()) - 1 for col in columnas_seleccionadas.split(",")]
            data = self.df.iloc[filas, columnas]

            plt.figure(figsize=(8, 6))
            if tipo_grafico == "linea":
                for i in range(len(filas)):
                    plt.plot(data.columns, data.iloc[i], label=f"Fila {filas[i]+1}")
                plt.xlabel("Columnas")
                plt.ylabel("Valor")
                plt.title("Gráfico de Líneas")
            elif tipo_grafico == "barra":
                for i in range(len(filas)):
                    plt.bar(data.columns, data.iloc[i], label=f"Fila {filas[i]+1}")
                plt.xlabel("Columnas")
                plt.ylabel("Valor")
                plt.title("Gráfico de Barras")
            elif tipo_grafico == "scatter":
                for i in range(len(filas)):
                    plt.scatter(data.columns, data.iloc[i], label=f"Fila {filas[i]+1}")
                plt.xlabel("Columnas")
                plt.ylabel("Valor")
                plt.title("Gráfico de Dispersión")
            elif tipo_grafico == "pastel":
                plt.pie(data.iloc[0], labels=data.columns, autopct='%1.1f%%')
                plt.title("Gráfico de Pastel")
                plt.axis('equal')
            else:
                QMessageBox.critical(self, "Error", "Tipo de gráfico no reconocido")
                return

            plt.legend()
            plt.grid(True)
            img_path = os.path.join(os.path.expanduser('~'), 'temp_chart.png')
            plt.savefig(img_path)
            plt.close()

            with xw.App(visible=False) as app:
                wb = xw.Book(self.file_path)
                sheet = wb.sheets[0]
                sheet.pictures.add(img_path, left=sheet.range("J1").left, top=sheet.range("J1").top)
                wb.save()
                wb.close()

            os.remove(img_path)
            QMessageBox.information(self, "Info", "Gráfico creado y guardado exitosamente")
        
            os.startfile(self.file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Se produjo un error: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelGraphApp()
    window.show()
    sys.exit(app.exec_())
