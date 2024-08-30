import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog, QComboBox, QMessageBox
from PyQt5.QtGui import QMovie
import tensorflow as tf
from sklearn.preprocessing import StandardScaler
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Dense
from tensorflow.keras.callbacks import EarlyStopping


class LoadingScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(50, 50)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.CustomizeWindowHint | Qt.FramelessWindowHint)
        
        self.label_animation = QLabel(self)
        
        self.movie = QMovie('Spinner.gif')
        self.label_animation.setMovie(self.movie)
        
        self.startAnimation()
        self.show()

    def startAnimation(self):
        self.movie.start()

    def stopAnimation(self):
        self.movie.stop()
        self.close()

class Worker(QThread):
    finished = pyqtSignal()
    plot_signal = pyqtSignal(dict)
                         
    def __init__(self, parent, archivo, operacion, columna1, columna2):
        super().__init__()
        self.archivo = archivo        
        self.columna1 = columna1
        self.columna2 = columna2
        self.operacion = operacion
        self.parent_window = parent

    def run(self):
        self.started.emit()

        self.parent_window.set_form_enabled(False)
        self.cargar_datos()
        self.parent_window.set_form_enabled(True)

        self.finished.emit()

    def cargar_datos(self):
        np.random.seed(0)
        tf.random.set_seed(0)

        datos = pd.read_excel(self.archivo)
        datos = datos.dropna(subset=[self.columna1, self.columna2])

        if self.operacion == "Suma":
            y = datos[self.columna1] + datos[self.columna2]
        elif self.operacion == "Multiplicacion":
            y = datos[self.columna1] * datos[self.columna2]
        elif self.operacion == "Resta":
            y = datos[self.columna1] - datos[self.columna2]
        elif self.operacion == "Division":
            y = datos[self.columna1] / datos[self.columna2]

        X = datos[[self.columna1, self.columna2]].values

        scaler = StandardScaler()
        X = scaler.fit_transform(X)

        model = Sequential([
            Dense(2, activation='relu', input_shape=(2,), kernel_regularizer=tf.keras.regularizers.l2(0.1)),
            Dense(1)])

        # Compilar el modelo
        model.compile(optimizer=tf.keras.optimizers.Adam(learning_rate=0.01), loss='mae')

        # Definir callbacks para detener el entrenamiento temprano si la pérdida en validación no mejora
        early_stopping = EarlyStopping(monitor='val_loss', patience=3)

        # Entrenar el modelo
        history = model.fit(X, y, epochs=600, batch_size=32, validation_split=0.2, callbacks=[early_stopping])

        # Graficar la pérdida durante el entrenamiento
        #plt.plot(history.history['loss'], label='Training Loss')
        #plt.plot(history.history['val_loss'], label='Validation Loss')
        #plt.xlabel('Epoch')
        #plt.ylabel('Loss')
        #plt.title('Loss during Training and Validation')
        #plt.legend()
        #plt.show()

        self.plot_signal.emit({'loss': history.history['loss'], 'val_loss': history.history['val_loss']})

        # Predicciones con los nuevos datos
        X_nuevos = scaler.transform(datos[[self.columna1, self.columna2]].values)
        predicciones = model.predict(X_nuevos)

        # Mostrar las predicciones de manera horizontal
        print("\nPredicciones:")
        resultados_prediccion = pd.DataFrame({self.operacion: predicciones.flatten()})
        print(resultados_prediccion)

        # Predicciones redondeadas
        print("\nPredicciones Redondeadas:")
        resultados_prediccion_redondeados = resultados_prediccion.applymap(round)
        print(resultados_prediccion_redondeados)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(300, 200) 
        self.setWindowTitle("Interfaz para Cargar Datos")

        layout = QVBoxLayout()

        # ComboBox para seleccionar la operación
        self.combo_operaciones = QComboBox()
        self.combo_operaciones.addItem("-- Seleccione --")
        self.combo_operaciones.addItems(["Suma", "Multiplicacion", "Resta", "Division"])
        layout.addWidget(QLabel("Operación:"))
        layout.addWidget(self.combo_operaciones)

        # ComboBox para seleccionar las columnas
        self.combo_columna1 = QComboBox()
        self.combo_columna2 = QComboBox()
        self.combo_columna1.addItem("-- Seleccione --")
        self.combo_columna2.addItem("-- Seleccione --")
        self.combo_columna1.addItems(["Datoprueba1", "Datoprueba2", "Datoprueba3", "Datoprueba4"])
        self.combo_columna2.addItems(["Datoprueba1", "Datoprueba2", "Datoprueba3", "Datoprueba4"])
        layout.addWidget(QLabel("Columna 1:"))
        layout.addWidget(self.combo_columna1)
        layout.addWidget(QLabel("Columna 2:"))
        layout.addWidget(self.combo_columna2)

        # Botón para cargar datos
        self.btn_cargar_datos = QPushButton("Cargar Datos")
        self.btn_cargar_datos.clicked.connect(self.cargar_datos_desde_interfaz)
        layout.addWidget(self.btn_cargar_datos)
        self.setLayout(layout)

    def cargar_datos_desde_interfaz(self):
        operacion_seleccionada = self.combo_operaciones.currentText()
        columna1_seleccionada = self.combo_columna1.currentText()
        columna2_seleccionada = self.combo_columna2.currentText()

        if operacion_seleccionada == "-- Seleccione --" or columna1_seleccionada == "-- Seleccione --" or columna2_seleccionada == "-- Seleccione --":
            QMessageBox.warning(self, "Alerta", "Debe seleccionar todas las opciones.")
        else:
            archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel", "", "Archivos de Excel (*.xlsx)")
            if archivo:
                self.loading_screen = LoadingScreen()
                self.worker = Worker(self, archivo, operacion_seleccionada, columna1_seleccionada, columna2_seleccionada)
                self.worker.started.connect(self.loading_screen.startAnimation)
                self.worker.plot_signal.connect(self.show_plot)
                self.worker.finished.connect(self.loading_screen.stopAnimation)
                self.worker.finished.connect(self.worker.deleteLater)
                self.worker.start()    

    def set_form_enabled(self, enabled):
        for widget in self.findChildren(QWidget):
            widget.setEnabled(enabled)

    def show_plot(self, data):
        plt.plot(data['loss'], label='Training Loss')
        plt.plot(data['val_loss'], label='Validation Loss')
        plt.xlabel('Epoch')
        plt.ylabel('Loss')
        plt.title('Loss during Training and Validation')
        plt.legend()
        plt.show()
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
