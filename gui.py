import sys
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QHBoxLayout, QCheckBox, QMessageBox, QDateEdit, QPlainTextEdit
)
from PySide6.QtCore import QDate, QThread, Signal

# Importa as funções que farão a ponte com nosso back-end
from sap_interface import executar_consultas_sap
from excel_writer import exportar_para_excel

TRANSACOES = [
    "ZTMMQ123", "ME5A", "IW39","IWBK", "IW29", "SQVI"
]

class WorkerThread(QThread):
    log_signal = Signal(str)
    finished_signal = Signal(bool, str)

    def __init__(self, pasta, transacoes_selecionadas, centros, start_date, end_date):
        super().__init__()
        self.pasta = pasta
        self.transacoes = transacoes_selecionadas
        self.centros = centros
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        try:
            # Passa as datas da GUI para a lógica de automação
            executar_consultas_sap(
                pasta_destino=self.pasta,
                transacoes_selecionadas=self.transacoes,
                log_callback=self.log_signal.emit,
                centros=self.centros,
                start_date=self.start_date,
                end_date=self.end_date
            )
            self.finished_signal.emit(True, "")
        except Exception as e:
            self.log_signal.emit(f"❌ Erro fatal na automação: {str(e)}")
            self.finished_signal.emit(False, str(e))

class SAPGui(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automação de Consultas SAP v1.1")
        self.setFixedSize(600, 600)

        self.layout = QVBoxLayout(self)
        self.setStyleSheet("""
            QWidget { font-size: 11pt; }
            QPushButton { padding: 8px; background-color: #0078D7; color: white; border-radius: 4px; }
            QPushButton:hover { background-color: #005A9E; }
            QPushButton:disabled { background-color: #A0A0A0; }
            QPlainTextEdit { background-color: #2B2B2B; color: #BBBBBB; font-family: Consolas; border-radius: 4px; }
            QLabel { margin-top: 5px; }
            QDateEdit { padding: 4px; }
        """)

        self.layout.addWidget(QLabel("<b>1. Selecione a Pasta de Destino:</b>"))
        self.btn_pasta = QPushButton("Selecionar Pasta")
        self.btn_pasta.clicked.connect(self.selecionar_pasta)
        self.layout.addWidget(self.btn_pasta)
        self.lbl_pasta = QLabel("<i>Nenhuma pasta selecionada</i>")
        self.layout.addWidget(self.lbl_pasta)

        # --- SELEÇÃO DE DATAS ---
        self.layout.addWidget(QLabel("<b>2. Selecione o Período da Consulta:</b>"))
        datas_layout = QHBoxLayout()
        self.data_inicio = QDateEdit()
        self.data_inicio.setCalendarPopup(True)
        self.data_inicio.setDisplayFormat("dd/MM/yyyy")
        self.data_inicio.setDate(QDate(2023, 1, 1)) # Data de início padrão
        
        self.data_fim = QDateEdit()
        self.data_fim.setCalendarPopup(True)
        self.data_fim.setDisplayFormat("dd/MM/yyyy")
        self.data_fim.setDate(QDate.currentDate()) # Data de fim padrão (hoje)
        
        datas_layout.addWidget(QLabel("Início:"))
        datas_layout.addWidget(self.data_inicio)
        datas_layout.addWidget(QLabel("Fim:"))
        datas_layout.addWidget(self.data_fim)
        self.layout.addLayout(datas_layout)

        self.layout.addWidget(QLabel("<b>3. Selecione as Transações:</b>"))
        self.checkboxes = {}
        for t in TRANSACOES:
            cb = QCheckBox(t)
            cb.setChecked(True)
            self.layout.addWidget(cb)
            self.checkboxes[t] = cb

        btns_layout = QHBoxLayout()
        self.btn_executar = QPushButton("▶️ Executar Consultas SAP")
        self.btn_executar.clicked.connect(self.iniciar_execucao_thread)
        btns_layout.addWidget(self.btn_executar)

        self.btn_excel = QPushButton("📊 Exportar para Excel")
        self.btn_excel.clicked.connect(self.exportar_excel)
        btns_layout.addWidget(self.btn_excel)
        self.layout.addLayout(btns_layout)

        self.layout.addWidget(QLabel("<b>Log de Execução:</b>"))
        self.log_console = QPlainTextEdit()
        self.log_console.setReadOnly(True)
        self.layout.addWidget(self.log_console)

        self.pasta = ""
        self.worker = None 

    def selecionar_pasta(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if pasta:
            self.pasta = pasta
            self.lbl_pasta.setText(f"<b>Pasta:</b> {pasta}")

    def atualizar_status(self, mensagem):
        self.log_console.appendPlainText(mensagem)
        self.log_console.verticalScrollBar().setValue(self.log_console.verticalScrollBar().maximum())

    def iniciar_execucao_thread(self):
        if not self.pasta:
            QMessageBox.critical(self, "Erro", "Por favor, selecione a pasta de destino antes de continuar.")
            return
        
        transacoes = [t for t, cb in self.checkboxes.items() if cb.isChecked()]
        if not transacoes:
            QMessageBox.warning(self, "Aviso", "Selecione ao menos uma transação para executar.")
            return
        
        # Pega as datas selecionadas na interface
        inicio = self.data_inicio.date().toString("dd.MM.yyyy")
        fim = self.data_fim.date().toString("dd.MM.yyyy")
        
        centros = ["L001", "T001", "T002", "T003", "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011"]

        self.btn_executar.setEnabled(False)
        self.btn_executar.setText("Executando...")
        self.log_console.clear()

        # Passa as datas para a thread
        self.worker = WorkerThread(self.pasta, transacoes, centros, inicio, fim)
        self.worker.log_signal.connect(self.atualizar_status)
        self.worker.finished_signal.connect(self.finalizar_execucao)
        self.worker.start()

    def finalizar_execucao(self, sucesso, erro_msg):
        self.btn_executar.setEnabled(True)
        self.btn_executar.setText("▶️ Executar Consultas SAP")
        if not sucesso:
            QMessageBox.critical(self, "Erro de Execução", f"A automação encontrou um erro:\n\n{erro_msg}")

    def exportar_excel(self):
        if not self.pasta:
            QMessageBox.critical(self, "Erro", "Selecione a pasta onde os arquivos TXT foram salvos.")
            return
        
        caminho_excel, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo Excel", self.pasta, "Excel Files (*.xlsx)")
        if not caminho_excel:
            return
        
        try:
            self.atualizar_status("\n🔄 Exportando dados para o Excel...")
            exportar_para_excel(self.pasta, caminho_excel, self.atualizar_status)
            self.atualizar_status(f"✅ Exportação concluída com sucesso em:\n{caminho_excel}")
            QMessageBox.information(self, "Sucesso", "O arquivo Excel foi gerado com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao exportar para o Excel:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = SAPGui()
    janela.show()
    sys.exit(app.exec())
