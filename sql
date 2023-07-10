import sys
from PyQt6.QtWidgets import (
    QTextEdit, QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QLineEdit, QPushButton, QLabel, QDialog, QDialogButtonBox, QFormLayout,
    QMessageBox, QTabWidget, QRadioButton 
)
import win32print
import tempfile
import os
import datetime
import mysql.connector
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtCore import Qt
from PyQt6.QtPrintSupport import QPrintDialog



class ConnectionDialog(QDialog):
    def __init__(self, parent=None, host="", port=3306, database="", user="", password=""):
        super().__init__(parent)
        self.setWindowTitle("Alterar Conexão")
        self.setModal(True)

        self.host_edit = QLineEdit()
        self.host_edit.setText(host)
        self.port_edit = QLineEdit()
        self.port_edit.setText(str(port))
        self.database_edit = QLineEdit()
        self.database_edit.setText(database)
        self.user_edit = QLineEdit()
        self.user_edit.setText(user)
        self.password_edit = QLineEdit()
        self.password_edit.setText(password)
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)

        form_layout = QFormLayout()
        form_layout.addRow("Host:", self.host_edit)
        form_layout.addRow("Porta:", self.port_edit)
        form_layout.addRow("Banco de Dados:", self.database_edit)
        form_layout.addRow("Usuário:", self.user_edit)
        form_layout.addRow("Senha:", self.password_edit)
  
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        td_layout = QVBoxLayout()
        td_layout.addLayout(form_layout)
        td_layout.addWidget(button_box)

        self.setLayout(td_layout)

    def get_connection_details(self):
        host = self.host_edit.text()
        port = int(self.port_edit.text())
        database = self.database_edit.text()
        user = self.user_edit.text()
        password = self.password_edit.text()
        charset = "utf8"
        return host, port, database, user, password, charset

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Busca informações")
        self.resize(800, 600)
    
        # self.host = "Servidor"
        # self.port = 11111111
        # self.database = "Banco"
        # self.user = "User"
        # self.password = "Pass"
        # self.charset = "999"

        self.connection_details = {
            "host": self.host,
            "port": self.port,
            "database": self.database,
            "user": self.user,
            "password": self.password,
            "charset": self.charset
        }

        self.setup_ui()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Enter or event.key() == Qt.Key.Key_Return:
            self.execute_query()
        else:
            super().keyPressEvent(event)

    def setup_ui(self):
        # Tab Widget
        tab_widget = QTabWidget()

        # Aba "Certidão de Buscas TD"
        td_tab = QWidget()
        td_layout = QVBoxLayout()

        # Título
        title_label = QLabel("Certidão de Buscas Gerais")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont("Arial Narrow", 18, QFont.Weight.Bold)
        title_label.setFont(title_font)
        td_layout.addWidget(title_label)


        # Campos para o intervalo de datas
        date_layout = QHBoxLayout()
        data_inicio_label = QLabel("Data inicial:")
        self.data_inicio_edit = QLineEdit("01/01/2022")
        self.data_inicio_edit.setInputMask("99/99/9999")
        data_final_label = QLabel("Data final:")
        self.end_date_edit = QLineEdit("31/12/2022")
        self.end_date_edit.setInputMask("99/99/9999")
        date_layout.addWidget(data_inicio_label)
        date_layout.addWidget(self.data_inicio_edit)
        date_layout.addWidget(data_final_label)
        date_layout.addWidget(self.end_date_edit)
        td_layout.addLayout(date_layout)

        # Campo para o CPF
        cpf_layout = QHBoxLayout()
        cpf_label = QLabel("CPF/CNPJ:")
        self.cpf_edit = QLineEdit()

        cpf_layout.addWidget(cpf_label)
        cpf_layout.addWidget(self.cpf_edit)
        td_layout.addLayout(cpf_layout)

        # Botão para executar a consulta
        self.execute_button = QPushButton("Realizar Consulta")
        self.execute_button.clicked.connect(self.execute_query)
        td_layout.addWidget(self.execute_button)

        self.result_count_label = QLabel("Resultados encontrados:")
        td_layout.addWidget(self.result_count_label)
        
        # Campo de texto para exibir os resultados formatados
        self.result_text_edit = QTextEdit()
        td_layout.addWidget(self.result_text_edit)

        self.quantidade_paginas_label = QLabel("Quantidade de Páginas")
        td_layout.addWidget(self.quantidade_paginas_label)

        self.valor_certidao_label = QLabel("Valor da certidão:")
        td_layout.addWidget(self.valor_certidao_label)

        self.custo_total_label = QLabel("Custo total:")
        td_layout.addWidget(self.custo_total_label)

        # Botão para copiar resultados
        self.copy_button = QPushButton("Copiar Resultados")
        self.copy_button.setEnabled(False)
        self.copy_button.clicked.connect(self.copy_results)
        td_layout.addWidget(self.copy_button)

        # Botao para limpar a tela
        self.clear_button = QPushButton("Limpar")
        self.clear_button.clicked.connect(self.clear_fields)
        td_layout.addWidget(self.clear_button)

        # Widget principal
        widget = QWidget()
        widget.setLayout(td_layout)
        self.setCentralWidget(widget)

        # Menu superior
        menu_bar = self.menuBar()

        # Menu "Conexão"
        connection_menu = menu_bar.addMenu("Conexão")

        # Ação "Alterar Conexão"
        change_connection_action = QAction("Alterar Conexão", self,)
        change_connection_action.triggered.connect(self.change_connection)
        connection_menu.addAction(change_connection_action)

        td_tab.setLayout(td_layout)
        tab_widget.addTab(td_tab, "Certidão de Buscas TD")

        # Aba "Indicador Pessoal de TD"
        ips_tab = QWidget()
        ips_layout = QVBoxLayout()

        # Título
        title_label = QLabel("Indicador Pessoal de TD")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont("Arial Narrow", 18, QFont.Weight.Bold)
        title_label.setFont(title_font)
        ips_layout.addWidget(title_label)

        self.cpf_radio_button = QRadioButton("CPF")
        self.cnpj_radio_button = QRadioButton("CNPJ")

        # Layout para os radio buttons
        type_layout = QVBoxLayout()
        type_layout.addWidget(self.cpf_radio_button)
        type_layout.addWidget(self.cnpj_radio_button)
        ips_layout.addLayout(type_layout)

        self.cpf_radio_button.setChecked(False)
        self.cpf_radio_button.toggled.connect(self.toggle_cpf_cnpj)
        self.cnpj_radio_button.toggled.connect(self.toggle_cpf_cnpj)

        # Campo para o CPF
        cpf_layout = QHBoxLayout()
        cpf_label = QLabel("CPF:")
        self.cpf_edit_ips = QLineEdit()
        self.cpf_edit_ips.setInputMask("999.999.999-99")
        cpf_layout.addWidget(cpf_label)
        cpf_layout.addWidget(self.cpf_edit_ips)
        ips_layout.addLayout(cpf_layout)

        # Campo para o CNPJ
        cnpj_layout = QHBoxLayout()
        cnpj_label = QLabel("CNPJ:")
        self.cnpj_edit_ips = QLineEdit()
        self.cnpj_edit_ips.setInputMask("99.999.999/9999-99")
        cnpj_layout.addWidget(cnpj_label)
        cnpj_layout.addWidget(self.cnpj_edit_ips)
        ips_layout.addLayout(cnpj_layout)

        # Campos para o período de data
        date_layout = QHBoxLayout()
        data_inicio_label = QLabel("Data Inicial:")
        self.data_inicio_edit_ips = QLineEdit()
        self.data_inicio_edit_ips.setInputMask("99/99/9999")
        data_final_label = QLabel("Data Final:")
        self.data_final_edit_ips = QLineEdit()
        self.data_final_edit_ips.setInputMask("99/99/9999")
        date_layout.addWidget(data_inicio_label)
        date_layout.addWidget(self.data_inicio_edit_ips)
        date_layout.addWidget(data_final_label)
        date_layout.addWidget(self.data_final_edit_ips)
        ips_layout.addLayout(date_layout)

        # Botão para pesquisar
        self.search_button_ips = QPushButton("Pesquisar")
        self.search_button_ips.clicked.connect(self.search_ips)
        ips_layout.addWidget(self.search_button_ips)

        # Resultado da pesquisa
        self.result_text_edit_ips = QTextEdit()
        ips_layout.addWidget(self.result_text_edit_ips)
        self.result_text_edit_ips.setReadOnly(True)

        ips_tab.setLayout(ips_layout)
        tab_widget.addTab(ips_tab, "Indicador Pessoal de TD")

        # Botão para imprimir o resultado
        self.print_result_button = QPushButton("Imprimir Resultado")
        self.print_result_button.clicked.connect(self.print_result)
        ips_layout.addWidget(self.print_result_button)

        # Botao para limpar a tela
        self.clear_button_ips = QPushButton("Limpar")
        self.clear_button_ips.clicked.connect(self.clear_fields_ips)
        ips_layout.addWidget(self.clear_button_ips)


        self.setCentralWidget(tab_widget)

    def print_result(self):
        # Obter o texto do resultado
        result_text = self.result_text_edit_ips.toPlainText()

        try:
            # Criar um arquivo temporário para armazenar o texto
            temp_file = tempfile.mktemp(".txt")
            with open(temp_file, "w") as file:
                file.write(result_text)

            # Obter o nome do arquivo temporário
            file_name = os.path.abspath(temp_file)

            # Abrir a impressora padrão
            printer_handle = win32print.OpenPrinter(win32print.GetDefaultPrinter())

            try:
                # Iniciar um trabalho de impressão
                win32print.StartDocPrinter(printer_handle, 1, (file_name, None, "RAW"))

                try:
                    # Iniciar uma página de impressão
                    win32print.StartPagePrinter(printer_handle)

                    # Imprimir o conteúdo do arquivo
                    with open(file_name, "rb") as file:
                        data = file.read()
                        win32print.WritePrinter(printer_handle, data)

                    # Finalizar a página de impressão
                    win32print.EndPagePrinter(printer_handle)

                finally:
                    # Finalizar o trabalho de impressão
                    win32print.EndDocPrinter(printer_handle)

            finally:
                # Fechar a impressora
                win32print.ClosePrinter(printer_handle)

        finally:
            # Remover o arquivo temporário
            os.remove(temp_file)

    def clear_fields(self):
        # Exibir o diálogo de confirmação
        confirm_dialog = QMessageBox(self)
        confirm_dialog.setIcon(QMessageBox.Icon.Question)
        confirm_dialog.setText("Deseja mesmo limpar a tela?")
        confirm_dialog.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        confirm_dialog.setDefaultButton(QMessageBox.StandardButton.No)

        #traducao dos botoes.
        yes_button = confirm_dialog.button(QMessageBox.StandardButton.Yes)
        yes_button.setText("Sim")

        no_button = confirm_dialog.button(QMessageBox.StandardButton.No)
        no_button.setText("Não")

        # Verificar a resposta do usuário
        response = confirm_dialog.exec()

        if response == QMessageBox.StandardButton.Yes:
            # Limpar os campos
            self.data_inicio_edit.clear()
            self.result_count_label.clear()
            self.quantidade_paginas_label.clear()
            self.end_date_edit.clear()
            self.cpf_edit.clear()
            self.valor_certidao_label.clear()
            self.custo_total_label.clear()
            self.quantidade_paginas_label
            self.result_text_edit.clear()
            # self.copy_button.setEnabled(False)
            # self.clear_button.setEnabled(False)

    def clear_fields_ips(self):
        # Exibir o diálogo de confirmação
        confirm_dialog = QMessageBox(self)
        confirm_dialog.setIcon(QMessageBox.Icon.Question)
        confirm_dialog.setText("Deseja mesmo limpar a tela?")
        confirm_dialog.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        confirm_dialog.setDefaultButton(QMessageBox.StandardButton.No)

        #traducao dos botoes.
        yes_button = confirm_dialog.button(QMessageBox.StandardButton.Yes)
        yes_button.setText("Sim")

        no_button = confirm_dialog.button(QMessageBox.StandardButton.No)
        no_button.setText("Não")

        # Verificar a resposta do usuário
        response = confirm_dialog.exec()

        if response == QMessageBox.StandardButton.Yes:
            # Limpar os campos
            self.data_inicio_edit_ips.clear()
            self.data_final_edit_ips.clear()
            self.cpf_edit_ips.clear()
            self.cnpj_edit_ips.clear()
            self.result_text_edit_ips.clear()



    def toggle_cpf_cnpj(self):
        if self.cpf_radio_button.isChecked():
            self.cpf_edit_ips.setEnabled(True)
            self.cnpj_edit_ips.setEnabled(False)
        else:
            self.cpf_edit_ips.setEnabled(False)
            self.cnpj_edit_ips.setEnabled(True)

    def execute_ips_query(self, query, params):
        cpf_or_cnpj = params['cpf_or_cnpj']
        data_inicio = params['data_inicio']
        data_final = params['data_final']
        try:
            # Estabelecer conexão com o banco de dados
            connection = mysql.connector.connect(
                host=self.connection_details["host"],
                port=self.connection_details["port"],
                database=self.connection_details["database"],
                user=self.connection_details["user"],
                password=self.connection_details["password"],
                charset=self.connection_details["charset"]
            )

            # Criar um cursor para executar a consulta
            cursor = connection.cursor()

            # Executar a consulta
            cursor.execute(query, {'cpf_or_cnpj': cpf_or_cnpj, 'data_inicio': data_inicio, 'data_final': data_final})

            # Obter os resultados da consulta
            results = cursor.fetchall()

            # Fechar o cursor e a conexão com o banco de dados
            cursor.close()
            connection.close()

            return results

        except mysql.connector.Error as error:
            print(f"Erro ao executar a consulta: {error}")

        return None


    def convert_date_format(date_string):
        # Converter a data de dd/MM/aaaa para aaaa-mm-dd
        date_object = datetime.strptime(date_string, '%d/%m/%Y')
        formatted_date = date_object.strftime('%Y-%m-%d')
        return formatted_date

    def search_ips(self):
        if self.cpf_radio_button.isChecked():
            cpf_or_cnpj = self.cpf_edit_ips.text()
        else:
            cpf_or_cnpj = self.cnpj_edit_ips.text()

        data_inicio = self.data_inicio_edit_ips.text()
        data_final = self.data_final_edit_ips.text()
        data_inicio = self.convert_date_format(data_inicio)
        data_final = self.convert_date_format(data_final)

        # Consulta SQL
        query = """
        SELECT
            titulo.protocolo AS Protocolo,
            titulo.pdata AS Data_Protocolo,
            IF(titulo.registro <> '', titulo.registro, titulo.registropri) AS 'Registro/Averbacao',
            IF(titulo.registro <> '', titulo.rdata, titulo.dataaverb) AS Data,
            titulo.nomenat AS Natureza,
            CONCAT(pessoa.Nome, ' CPF/CNPJ: ', pessoa.CpfCgc) AS Nome,
            denominacaocontratante.den_descricao AS Qualificacao
        FROM
            document.titulo
            INNER JOIN document.contratante ON contratante.sequencia = titulo.sequencia
            INNER JOIN document.pessoa ON pessoa.CodPes = contratante.CodPes
            LEFT JOIN document.denominacaocontratante ON denominacaocontratante.den_id = contratante.Denominacao
        WHERE
            denominacaocontratante.den_descricao NOT LIKE '%Procurador%' 
            AND denominacaocontratante.den_descricao NOT LIKE '%Representante%' 
            AND (
                (titulo.rdata BETWEEN %(data_inicio)s AND %(data_final)s)
                OR (
                    titulo.DataAverb BETWEEN %(data_inicio)s AND %(data_final)s
                    AND IF(titulo.dataaverb IS NOT NULL AND titulo.registropri <> '00000000', 1, 0) = 1
                )
            )
            AND (pessoa.CpfCgc = %(cpf_or_cnpj)s)
        GROUP BY titulo.sequencia
        ORDER BY IF(titulo.registro <> '', titulo.registro, titulo.registropri)
        """

        # Executar a consulta
        params = {'cpf_or_cnpj': cpf_or_cnpj, 'data_inicio': data_inicio, 'data_final': data_final}
        results = self.execute_ips_query(query, params)

        if results:
            result_text = ""

            # Cabeçalho
            result_text += "***********************************\n"
            result_text += "** INDICADOR PESSOAL DE TD **\n"
            result_text += "***********************************\n"
            result_text += "\n"
            result_text += "NOME: {}\n".format(results[0][5])  # Índice 5
            result_text += "CPF / CNPJ: {}\n".format(cpf_or_cnpj)
            result_text += "------------------------------------------------------------------------------------------------------------------------------------\n"
            result_text += "Protocolo\tData\tRegistro\tData\tQualificação\n"
            result_text += "------------------------------------------------------------------------------------------------------------------------------------\n"

            # Resultados por protocolo
            protocolo_anterior = ""
            for row, result in enumerate(results):
                protocolo = result[0]  # Índice 0
                data = result[1].strftime("%d/%m/%Y")  # Índice 1
                registro = result[2]  # Índice 2
                data_registro = result[3].strftime("%d/%m/%Y")  # Índice 3
                qualificacao = result[6]  # Índice 6
                natureza = result[4]  # Índice 4

                if protocolo != protocolo_anterior:
                    #result_text += "{}\n".format(protocolo)
                    result_text += "{}\t{}\t{}\t{}\t{}\n".format(protocolo, data, registro, data_registro, qualificacao)
                    result_text += "Natureza: {}\n".format(natureza)
                    result_text += "------------------------------------------------------------------------------------------------------------------------------------\n"
                else:
                    result_text += "{}\t{}\t{}\t{}\t{}\n".format("", data, registro, data_registro, qualificacao)

                protocolo_anterior = protocolo

            # Exibir o texto na caixa de texto
            self.result_text_edit_ips.setText(result_text)
        else:
            self.result_text_edit_ips.setText("Não foram encontrados resultados para a pesquisa.")


    def convert_date_format(self, date):
    # Verifica se a data está no formato "ddmmyyyy"
        if len(date) == 8:
            day = date[0:2]
            month = date[2:4]
            year = date[4:8]
            return f"{year}-{month}-{day}"
        
        elif len(date) == 10:
            day, month, year = date.split('/')
            return f"{year}-{month}-{day}"
        return date

    def change_connection(self):
        dialog = ConnectionDialog(self, self.host, self.port, self.database, self.user, self.password)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.host, self.port, self.database, self.user, self.password, self.charset = dialog.get_connection_details()
            self.connection_details = {
                "host": self.host,
                "port": self.port,
                "database": self.database,
                "user": self.user,
                "password": self.password,
                "charset": self.charset
            }

    def execute_query(self):
        # Obter o intervalo de datas
        start_date = self.convert_date_format(self.data_inicio_edit.text())
        end_date = self.convert_date_format(self.end_date_edit.text())

        if start_date > end_date:
            QMessageBox.warning(self, "Alerta", "A data inicial não pode ser maior que a data final.")
            return
        # Obter o CPF
        cpf = self.cpf_edit.text()

        try:
            # Conectar ao banco de dados MySQL
            connection = mysql.connector.connect(
                host=self.connection_details["host"],
                port=self.connection_details["port"],
                database=self.connection_details["database"],
                user=self.connection_details["user"],
                password=self.connection_details["password"],
                charset=self.connection_details["charset"]
            )

            # Montar a consulta SQL com o intervalo de datas e CPF
            query = """
Omitido por seguranca!
""".format(start_date, end_date, start_date, end_date, start_date, end_date, start_date, end_date, cpf)

            # Executar a consulta
            cursor = connection.cursor()
            cursor.execute(query)
            results = cursor.fetchall()
            
            # Calcular o custo com base no número de páginas
            
            registro_count = len(results)
            self.result_count_label.setText(f"Resultados encontrados: {registro_count}")

            # Calcular o custo total
            custo_certidao = 87.16
            custo_pagina_adicional = 8.74
            numero_paginas = registro_count // 30  # Divisão inteira para obter o número de páginas

            custo_total = custo_certidao + custo_pagina_adicional * max(0, numero_paginas - 2)

            # Exibir o custo total para o usuário
            self.custo_total_label.setText(f"Custo total: R$ {custo_total:.2f}")
            self.valor_certidao_label.setText(f"Valor da certidão: R$ {custo_certidao:.2f}")
            self.quantidade_paginas_label.setText(f"Quantidade de Páginas: {numero_paginas}")
           
            # Formatar os resultados como texto
            grouped_results = {}
            registro_count = 0

            for row_data in results:
                registro_averbacao = row_data[0]
                natureza = row_data[1]
                texto = row_data[2]
                cpf_cnpj = row_data[3]
                descricao = row_data[4]

                if registro_averbacao not in grouped_results:
                    grouped_results[registro_averbacao] = {
                        'natureza': natureza,
                        'cpf_cnpj': set(),
                        'texto': texto,
                        'descricao': descricao
                    }
                    registro_count += 1

                grouped_results[registro_averbacao]['cpf_cnpj'].add(cpf_cnpj)

            formatted_results = ""
            
            for registro_averbacao, data in grouped_results.items():
                natureza = data['natureza']
                cpf_cnpj = ", ".join(data['cpf_cnpj'])
                texto = data['texto']
                descricao = data['descricao']

                formatted_results += f"Reg.: {registro_averbacao}\n"
                formatted_results += f"{texto}\n"
                #formatted_results += f"NATUREZA: {natureza}\n"
                formatted_results += f"CONTRATANTE: {cpf_cnpj}"
                formatted_results += f"{descricao}"
                formatted_results += "\n\n"

            formatted_results = formatted_results.strip()  # Remove leading/trailing whitespace
            

            self.result_text_edit.setPlainText(formatted_results)

            font = QFont("Arial Narrow", 10)
            self.result_text_edit.setFont(font)
            self.result_text_edit.setPlainText(formatted_results)
            self.result_text_edit.setReadOnly(True)
            
            # Ativar o botão de copiar resultados
            self.copy_button.setEnabled(True)

            # Fechar a conexão
            cursor.close()
            connection.close()

        except mysql.connector.Error as error:
            print(f"Erro na conexão ao banco de dados: {error}")

        if not results:
            self.result_text_edit.setPlainText("Não há dados para o CPF informado")
            self.copy_button.setEnabled(False)
            self.result_count_label.setText("Resultados encontrados: 0")
            return

        # Contar o número de resultados encontrados
        self.result_count_label.setText(f"Resultados encontrados: {registro_count}")
    
    def copy_results(self):
        results_text = self.result_text_edit.toPlainText()
        clipboard = QApplication.clipboard()
        clipboard.setText(results_text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
