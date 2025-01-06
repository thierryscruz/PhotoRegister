import sys
import os
import cv2
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLineEdit,
    QCompleter,
    QPushButton,
    QLabel,
    QFileDialog,
    QMessageBox,
    QComboBox
)
from PyQt5.QtCore import QTimer, Qt, QStringListModel
from PyQt5.QtGui import QImage, QPixmap


class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Photo Register")
        self.setGeometry(400, 300, 440, 300)
        self.data = None
        self.current_record = None
        self.captured_frame = None

        # Campo de busca
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("Digite para buscar...")
        self.search_input.setGeometry(20, 30, 400, 30)

        # Configurar autocompletar
        self.completer = QCompleter()
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.search_input.setCompleter(self.completer)
        self.search_input.textChanged.connect(self.update_completer)

        # Botão de carregar Excel
        self.load_button = QPushButton("Carregar Excel", self)
        self.load_button.setGeometry(20, 70, 100, 30)
        self.load_button.clicked.connect(self.load_excel)

        # Botão de buscar
        self.search_button = QPushButton("Buscar", self)
        self.search_button.setGeometry(130, 70, 100, 30)
        self.search_button.clicked.connect(self.search_and_capture)

        # Botão de capturar imagem
        self.capture_button = QPushButton("Capturar", self)
        self.capture_button.setGeometry(20, 240, 100, 30)
        self.capture_button.clicked.connect(self.capture_image)
        self.capture_button.setEnabled(False)

        # Botão de salvar imagem
        self.save_button = QPushButton("Salvar", self)
        self.save_button.setGeometry(170, 240, 100, 30)
        self.save_button.clicked.connect(self.save_image)
        self.save_button.setVisible(False)

        # Botão de descartar imagem
        self.discard_button = QPushButton("Descartar", self)
        self.discard_button.setGeometry(290, 240, 100, 30)
        self.discard_button.clicked.connect(self.discard_image)
        self.discard_button.setVisible(False)

        # Labels para exibir os resultados
        self.matricula_label = QLabel("Matrícula: ", self)
        self.matricula_label.setGeometry(20, 120, 400, 30)

        self.nome_label = QLabel("Nome: ", self)
        self.nome_label.setGeometry(20, 150, 400, 30)

        self.setor_label = QLabel("Setor: ", self)
        self.setor_label.setGeometry(20, 180, 400, 30)

        # QLabel para exibir a miniatura da webcam
        self.webcam_label = QLabel(self)
        self.webcam_label.setGeometry(240, 70, 180, 160)
        self.webcam_label.setStyleSheet("border: 1px solid black;")

        # Timer para atualizar a imagem da webcam em tempo real
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_webcam_image)

        # ComboBox para alternar entre tema claro e escuro
        self.theme__label = QLabel("Tema: ", self)
        self.theme__label.setGeometry(320, 265, 70, 30)
        self.theme_combo = QComboBox(self)
        self.theme_combo.addItems(["Claro", "Escuro"])
        self.theme_combo.setGeometry(350, 270, 70, 20)
        self.theme_combo.currentIndexChanged.connect(self.toggle_theme)

        # Inicializar atributos para webcam
        self.cap = None

        # Aplicar o tema claro por padrão
        self.apply_dark_theme()

    def apply_dark_theme(self):
        # Tema Escuro usando setStyleSheet
        dark_stylesheet = """
                        QMainWindow {
                            background-color: #2e2e2e;
                            border: 1px solid #444444;
                        }
                        QLineEdit {
                            background-color: #555555;
                            color: white;
                            border: 1px solid #888888;
                        }
                        QLabel {
                            color: white;
                        }
                        QPushButton {
                            background-color: #444444;
                            color: white;
                            border: 1px solid #888888;
                        }
                        QPushButton:hover {
                            background-color: #666666;
                        }
                        QComboBox {
                            background-color: #555555;
                            color: white;
                            border: 1px solid #888888;
                        }
                        QComboBox QAbstractItemView {
                            background-color: #444444;
                            color: white;
                        }
                        QMessageBox {
                            background-color: #444444;
                            color: white;
                            border: 1px solid #888888;
                        }
                        QMessageBox QPushButton {
                        background-color: #555555;
                        color: white;
                        border: 1px solid #888888;
                        min-width: 80px;  /* Largura mínima */
                        min-height: 10px; /* Altura mínima */
                        font-size: 14px;  /* Tamanho da fonte */
                        padding: 2px;    /* Preenchimento interno */
                        }
                        QMessageBox QPushButton:hover {
                            background-color: #666666;
                        }
                        """
        self.setStyleSheet(dark_stylesheet)

    def apply_light_theme(self):
        # Tema Claro usando setStyleSheet
        light_stylesheet = """
                        QMainWindow {
                            background-color: #f5f5f5;
                            border: 1px solid #dcdcdc;
                        }
                        """
        self.setStyleSheet(light_stylesheet)

    def toggle_theme(self):
        if self.theme_combo.currentText() == "Escuro":
            self.apply_dark_theme()
        else:
            self.apply_light_theme()

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Abrir Arquivo Excel", "", "Arquivos Excel (*.xls *.xlsx)")
        if file_path:
            try:
                self.data = pd.read_excel(file_path)
                self.data.columns = self.data.columns.str.strip().str.upper()
                QMessageBox.information(self, "Sucesso", "Arquivo carregado com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao carregar o arquivo: {e}")

    def update_completer(self, text):
        if self.data is not None and len(text) >= 4:
            try:
                filtered_names = self.data[self.data["NOME"].fillna("").str.contains(text, case=False, na=False)]["NOME"]
                model = QStringListModel(filtered_names.tolist())
                self.completer.setModel(model)
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao atualizar sugestões: {e}")

    def search_and_capture(self):
        if self.data is None:
            QMessageBox.warning(self, "Aviso", "Por favor, carregue um arquivo Excel primeiro.")
            return

        query = self.search_input.text().strip()
        if not query:
            QMessageBox.warning(self, "Aviso", "Digite um valor para pesquisa.")
            return

        try:
            result = self.data[self.data["NOME"].fillna("").str.contains(query, case=False, na=False)]
            if not result.empty:
                record = result.iloc[0]
                self.current_record = record
                self.matricula_label.setText(f"Matrícula: {record['MATRICULA']}")
                self.nome_label.setText(f"Nome: {record['NOME']}")
                self.setor_label.setText(f"Setor: {record['SETOR']}")
                self.capture_button.setEnabled(True)
                self.start_webcam()
            else:
                QMessageBox.information(self, "Não encontrado", "Nenhum registro encontrado.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro durante a busca: {e}")

    def start_webcam(self):
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            QMessageBox.critical(self, "Erro", "Não foi possível acessar a webcam.")
            return
        self.timer.start(30)

    def update_webcam_image(self):
        if self.cap:
            ret, frame = self.cap.read()
            if ret:
                rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                h, w, _ = rgb_image.shape
                qimg = QImage(rgb_image.data, w, h, 3 * w, QImage.Format_RGB888)
                pixmap = QPixmap(qimg)
                self.webcam_label.setPixmap(pixmap.scaled(180, 160, Qt.KeepAspectRatio))

    def capture_image(self):
        if self.cap:
            ret, frame = self.cap.read()
            if ret:
                self.captured_frame = frame
                self.timer.stop()
                rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                h, w, _ = rgb_image.shape
                qimg = QImage(rgb_image.data, w, h, 3 * w, QImage.Format_RGB888)
                pixmap = QPixmap(qimg)
                self.webcam_label.setPixmap(pixmap.scaled(180, 160, Qt.KeepAspectRatio))
                self.capture_button.setEnabled(False)
                self.save_button.setVisible(True)
                self.discard_button.setVisible(True)

    def save_image(self):
        if self.captured_frame is not None:
            # Salvar a imagem na pasta do aplicativo
            app_directory = os.path.dirname(os.path.abspath(__file__))
            matricula = str(self.current_record["MATRICULA"])
            filename = f"{matricula}.png"
            file_path = os.path.join(app_directory, filename)
            cv2.imwrite(file_path, self.captured_frame)
            QMessageBox.information(self, "Imagem Salva", f"Imagem salva em '{file_path}'.")
            self.reset_capture()

    def discard_image(self):
        self.reset_capture()

    def reset_capture(self):
        self.captured_frame = None
        self.capture_button.setEnabled(True)
        self.save_button.setVisible(False)
        self.discard_button.setVisible(False)
        self.start_webcam()

    def closeEvent(self, event):
        if self.cap and self.cap.isOpened():
            self.cap.release()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
