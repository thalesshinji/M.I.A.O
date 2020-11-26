import sys
from programa_gui import Ui_MainWindow
from PySide2 import QtCore, QtGui, QtWidgets
from PySide2.QtWidgets import QMessageBox, QInputDialog
import SysTrayIcon
import xlrd
import xlwt
import datetime
import time
from xlutils.copy import copy
import openpyxl

#Cabeçalho para herdar a gui
class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, codb = ''):
        super(MainWindow, self).__init__()
        self.setupUi(self)
        self.codb = codb

        # Inicio do script do programa de Controle de Produção.
        #Função que pesquisa a OF.
        def busca():
            coda = self.OF_principal.text()
            codc = self.OF_FaltaOF.text()
            self.OF_FaltaOF.clear()
            wb = xlrd.open_workbook('C:/Files Controle de Produção/FS17001_2019.xlsx', on_demand=True)
            sh = wb.sheet_by_index(0)
            tem = ''
            for x in range(sh.nrows):
                celula = sh.cell(x, 0)
                if coda != '':
                    if celula.value == coda:
                        tem = 'tem'
                        if tem == 'tem':
                            self.BackgroundVermelho.raise_()
                            self.cod_parada.setFocus()
                            self.label_produto.setText(str(sh.cell_value(rowx=x, colx=1)))
                            self.label_medida.setText(str(sh.cell_value(rowx=x, colx=2)))
                            self.label_lote.setText(str(sh.cell_value(rowx=x, colx=0)))
                            self.label_mat_pr.setText(str(sh.cell_value(rowx=x, colx=8)))
                            self.label_diam.setText(str(sh.cell_value(rowx=x, colx=9)))
                            self.label_maquina.setText(str(sh.cell_value(rowx=x, colx=6)))
                            self.label_data.setText('')
                if codc != '':
                    if celula.value == codc:
                        tem = 'tem'
                        if tem == 'tem':
                            self.BackgroundVermelho.raise_()
                            self.cod_parada.setFocus()
                            self.label_produto.setText(str(sh.cell_value(rowx=x, colx=1)))
                            self.label_medida.setText(str(sh.cell_value(rowx=x, colx=2)))
                            self.label_lote.setText(str(sh.cell_value(rowx=x, colx=0)))
                            self.label_mat_pr.setText(str(sh.cell_value(rowx=x, colx=8)))
                            self.label_diam.setText(str(sh.cell_value(rowx=x, colx=9)))
                            self.label_maquina.setText(str(sh.cell_value(rowx=x, colx=6)))
                            self.label_data.setText('')
                            self.OF_FaltaOF.returnPressed.disconnect()
            if tem != 'tem':
                QMessageBox.warning(self, 'ALERTA', f'{coda or codc:^} essa OF. não existe, digite novamente.'.upper())


        def gravar():
            wbg = openpyxl.load_workbook('C:/Files Controle de Produção/Dados_Controle_de_Produção_C12I_2020 - Copia.xlsx')
            shg = wbg.active
            shg['K4'] = '0:30:00'
            shg['J4'] = '0:30:00'
            wbg.save('C:/Files Controle de Produção/Dados_Controle_de_Produção_C12I_2020 - Copia.xlsx')


        #Alternando entre janelas principal/ paradas/ sem OF.
        def a():
            coda = self.OF_principal.text()
            if coda != '00':
                busca()
            if coda == '00':
                self.BackgroundFaltaOF.raise_()
                self.OF_FaltaOF.setFocus()
                def c():
                    busca()
                self.OF_FaltaOF.returnPressed.connect(c)

        self.OF_principal.returnPressed.connect(a)
        self.OF_principal.returnPressed.connect(self.OF_principal.clear)


        #Aguardando código da parada
        self.label_status_parada.setText('CÓDIGO DA PARADA?')

        #Alternando entre janelas de paradas
        def b():
            data_formatada = time.strftime("%d/%m/%Y", time.localtime())
            self.label_data.setText(data_formatada)
            self.codb = self.cod_parada.text()
            if self.codb == '01':
                self.label_status_parada.setText('PRODUÇÃO')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(0, 170, 0)")
                self.watch_pause_liberacao()
                self.watch_pause_parada()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_producao()
            if self.codb == '02':
                self.label_status_parada.setText('PREPARAÇÃO DE MÁQ. (P1)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
                gravar()



            if self.codb == '03':
                self.label_status_parada.setText('MÁQUINA EM STAND BY (M2)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '04':
                self.label_status_parada.setText('AFIAÇÃO DE FERRAMENTA (F3)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '05':
                self.label_status_parada.setText('FALTA DE FERRAMENTA (F4)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '06':
                self.label_status_parada.setText('FALTA DE PROGRAMADOR (F5)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '07':
                self.label_status_parada.setText('FALTA DE OPERADOR (F6)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '08':
                self.label_status_parada.setText('LIMPEZA (L7)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '09':
                self.label_status_parada.setText('MANUTENÇÃO (F4)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '10':
                self.label_status_parada.setText('AGUARDANDO LIBERAÇÃO(A9)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255,255, 0)")
                self.watch_pause_producao()
                self.watch_pause_parada()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_liberacao()
            if self.codb == '11':
                self.label_status_parada.setText('AJUSTE DE MÁQUINA (A10)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '12':
                self.label_status_parada.setText('FALTA DE MATÉRIA PRIMA (F11)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '13':
                self.label_status_parada.setText('FALTA DE ENERGIA')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '14':
                self.label_status_parada.setText('TRY OUT (T13)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 255)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_parada()
                self.watch_pause_intervalo()
                self.watch_pause_treinamento()
                self.start_watch_try_out()
            if self.codb == '15':
                self.label_status_parada.setText('ALIMENTAÇÃO (T14)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '16':
                self.label_status_parada.setText('TREINAMENTO (T15)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(0, 170, 255)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_parada()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.start_watch_treinamento()
            if self.codb == '17':
                self.label_status_parada.setText('INTERVALO (I16)')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(125, 125, 125)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_parada()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_intervalo()
            if self.codb == '18':
                self.label_status_parada.setText('FALHA DO SOFTWARE')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                self.start_watch_parada()
            if self.codb == '19':
                self.label_status_parada.setText('ENCERRAMENTO TURNO')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.watch_pause_producao()
                self.watch_pause_liberacao()
                self.watch_pause_parada()
                self.watch_pause_intervalo()
                self.watch_pause_try_out()
                self.watch_pause_treinamento()
                operador = QInputDialog.getInt(self, ("DADOS"), "OPERADOR")
                quantidade = QInputDialog.getInt(self, ("DADOS"), "QUANTIDADE PRODUZIDA")
                refugo = QInputDialog.getInt(self, ("DADOS"), "REFUGO")
            if self.codb == '20':
                self.label_status_parada.setText('CÓDIGO DA PARADA?')
                self.frame_Status_Parada.setStyleSheet(u"image: none;background-color: rgb(255, 0, 0)")
                self.stop_watch()
                self.BackgroundPrincipal.raise_()
                self.OF_principal.setFocus()
            if self.codb not in ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20'):
                QMessageBox.warning(self, 'ALERTA', 'CÓDIGO INVÁLIDO, TENTE NOVAMENTE.')

        self.cod_parada.returnPressed.connect(b)
        self.cod_parada.returnPressed.connect(self.cod_parada.clear)

        # Timer paradas
        self.timer_producao = QtCore.QTimer(self)
        self.timer_producao.timeout.connect(self.run_watch_producao)
        self.timer_liberacao = QtCore.QTimer(self)
        self.timer_liberacao.timeout.connect(self.run_watch_liberacao)
        self.timer_parada = QtCore.QTimer(self)
        self.timer_parada.timeout.connect(self.run_watch_parada)
        self.timer_intervalo = QtCore.QTimer(self)
        self.timer_intervalo.timeout.connect(self.run_watch_intervalo)
        self.timer_try_out = QtCore.QTimer(self)
        self.timer_try_out.timeout.connect(self.run_watch_try_out)
        self.timer_treinamento = QtCore.QTimer(self)
        self.timer_treinamento.timeout.connect(self.run_watch_treinamento)

        self.timer_producao.setInterval(1)
        self.timer_liberacao.setInterval(1)
        self.timer_parada.setInterval(1)
        self.timer_intervalo.setInterval(1)
        self.timer_try_out.setInterval(1)
        self.timer_treinamento.setInterval(1)

        self.mscounter_producao = 0
        self.mscounter_liberacao = 0
        self.mscounter_parada = 0
        self.mscounter_parada02 = 0
        self.mscounter_parada03 = 0
        self.mscounter_parada04 = 0
        self.mscounter_parada05 = 0
        self.mscounter_parada06 = 0
        self.mscounter_parada07 = 0
        self.mscounter_parada08 = 0
        self.mscounter_parada09 = 0
        self.mscounter_parada11 = 0
        self.mscounter_parada12 = 0
        self.mscounter_parada13 = 0
        self.mscounter_parada15 = 0
        self.mscounter_parada18 = 0
        self.mscounter_intervalo = 0
        self.mscounter_try_out = 0
        self.mscounter_treinamento = 0
        self.producao = True
        self.liberacao = True
        self.parada = True
        self.intervalo = True
        self.try_out = True
        self.treinamento = True
        self.showLCD()


    def showLCD(self):
        self.mscounter_parada = self.mscounter_parada02 + self.mscounter_parada03 + self.mscounter_parada04 + self.mscounter_parada05 + self.mscounter_parada06 + self.mscounter_parada07 + self.mscounter_parada08 + self.mscounter_parada09 + self.mscounter_parada11 + self.mscounter_parada12 + self.mscounter_parada13 + self.mscounter_parada15 + self.mscounter_parada18
        texto_producao = datetime.timedelta(milliseconds=self.mscounter_producao)
        texto_liberacao = datetime.timedelta(milliseconds=self.mscounter_liberacao)
        texto_parada = datetime.timedelta(milliseconds=self.mscounter_parada)
        texto_intervalo = datetime.timedelta(milliseconds=self.mscounter_intervalo)
        texto_try_out = datetime.timedelta(milliseconds=self.mscounter_try_out)
        texto_treinamento = datetime.timedelta(milliseconds=self.mscounter_treinamento)
        tempo_total = texto_producao + texto_liberacao + texto_parada + texto_intervalo + texto_try_out + texto_treinamento
        if self.producao == False:
            self.label_historico_producao_tempo.setText(f'{str(texto_producao)[:-7]}')
            porcentagem_producao = (texto_producao / tempo_total) * 100
            self.label_historico_producao_porcentagem.setText(f'{int(porcentagem_producao)}%')
        if self.liberacao == False:
            self.label_historico_liberacao_tempo.setText(f'{str(texto_liberacao)[:-7]}')
            porcentagem_liberacao = (texto_liberacao / tempo_total) * 100
            self.label_historico_liberacao_porcentagem.setText(f'{int(porcentagem_liberacao)}%')
        if self.parada == False:
            self.label_historico_parada_tempo.setText(f'{str(texto_parada)[:-7]}')
            porcentagem_parada = (texto_parada / tempo_total) * 100
            self.label_historico_parada_porcentagem.setText(f'{int(porcentagem_parada)}%')
            tempo_parada02 = str(datetime.timedelta(milliseconds=self.mscounter_parada02))[:-7]
            tempo_parada03 = str(datetime.timedelta(milliseconds=self.mscounter_parada03))[:-7]
            tempo_parada04 = str(datetime.timedelta(milliseconds=self.mscounter_parada04))[:-7]
            tempo_parada05 = str(datetime.timedelta(milliseconds=self.mscounter_parada05))[:-7]
            tempo_parada06 = str(datetime.timedelta(milliseconds=self.mscounter_parada06))[:-7]
            tempo_parada07 = str(datetime.timedelta(milliseconds=self.mscounter_parada07))[:-7]
            tempo_parada08 = str(datetime.timedelta(milliseconds=self.mscounter_parada08))[:-7]
            tempo_parada09 = str(datetime.timedelta(milliseconds=self.mscounter_parada09))[:-7]
            tempo_parada11 = str(datetime.timedelta(milliseconds=self.mscounter_parada11))[:-7]
            tempo_parada12 = str(datetime.timedelta(milliseconds=self.mscounter_parada12))[:-7]
            tempo_parada13 = str(datetime.timedelta(milliseconds=self.mscounter_parada13))[:-7]
            tempo_parada15 = str(datetime.timedelta(milliseconds=self.mscounter_parada15))[:-7]
            tempo_parada18 = str(datetime.timedelta(milliseconds=self.mscounter_parada18))[:-7]
        if self.intervalo == False:
            self.label_historico_intervalo_tempo.setText(f'{str(texto_intervalo)[:-7]}')
            porcentagem_intervalo = (texto_intervalo / tempo_total) * 100
            self.label_historico_intervalo_porcentagem.setText(f'{int(porcentagem_intervalo)}%')
        if self.try_out == False:
            self.label_historico_try_out_tempo.setText(f'{str(texto_try_out)[:-7]}')
            porcentagem_try_out = (texto_try_out / tempo_total) * 100
            self.label_historico_try_out_porcentagem.setText(f'{int(porcentagem_try_out)}%')
        if self.treinamento == False:
            self.label_historico_treinamento_tempo.setText(f'{str(texto_treinamento)[:-7]}')
            porcentagem_treinamento = (texto_treinamento / tempo_total) * 100
            self.label_historico_treinamento_porcentagem.setText(f'{int(porcentagem_treinamento)}%')
        if self.producao == True:
            self.label_historico_producao_tempo.setText('0:00:00')
            self.label_historico_producao_porcentagem.setText('0%')
        if self.liberacao == True:
            self.label_historico_liberacao_tempo.setText('0:00:00')
            self.label_historico_liberacao_porcentagem.setText('0%')
        if self.parada == True:
            self.label_historico_parada_tempo.setText('0:00:00')
            self.label_historico_parada_porcentagem.setText('0%')
        if self.intervalo == True:
            self.label_historico_intervalo_tempo.setText('0:00:00')
            self.label_historico_intervalo_porcentagem.setText('0%')
        if self.try_out == True:
            self.label_historico_try_out_tempo.setText('0:00:00')
            self.label_historico_try_out_porcentagem.setText('0%')
        if self.treinamento == True:
            self.label_historico_treinamento_tempo.setText('0:00:00')
            self.label_historico_treinamento_porcentagem.setText('0%')

    def run_watch_producao(self):
        self.mscounter_producao += 1.021
        self.showLCD()
    def run_watch_liberacao(self):
        self.mscounter_liberacao += 1.021
        self.showLCD()

    def run_watch_parada(self):
        if self.codb == '02':
            self.mscounter_parada02 += 1.021
        if self.codb == '03':
            self.mscounter_parada03 += 1.021
        if self.codb == '04':
            self.mscounter_parada04 += 1.021
        if self.codb == '05':
            self.mscounter_parada05 += 1.021
        if self.codb == '06':
            self.mscounter_parada06 += 1.021
        if self.codb == '07':
            self.mscounter_parada07 += 1.021
        if self.codb == '08':
            self.mscounter_parada08 += 1.021
        if self.codb == '09':
            self.mscounter_parada09 += 1.021
        if self.codb == '11':
            self.mscounter_parada11 += 1.021
        if self.codb == '12':
            self.mscounter_parada12 += 1.021
        if self.codb == '13':
            self.mscounter_parada13 += 1.021
        if self.codb == '15':
            self.mscounter_parada15 += 1.021
        if self.codb == '18':
            self.mscounter_parada18 += 1.021
        self.showLCD()
    def run_watch_intervalo(self):
        self.mscounter_intervalo += 1.021
        self.showLCD()
    def run_watch_try_out(self):
        self.mscounter_try_out += 1.021
        self.showLCD()
    def run_watch_treinamento(self):
        self.mscounter_treinamento += 1.021
        self.showLCD()

    def start_watch_producao(self):
        self.timer_producao.start()
        self.producao = False
    def start_watch_liberacao(self):
        self.timer_liberacao.start()
        self.liberacao = False
    def start_watch_parada(self):
        self.timer_parada.start()
        self.parada = False
    def start_watch_intervalo(self):
        self.timer_intervalo.start()
        self.intervalo = False
    def start_watch_try_out(self):
        self.timer_try_out.start()
        self.try_out = False
    def start_watch_treinamento(self):
        self.timer_treinamento.start()
        self.treinamento = False

    def watch_pause_producao(self):
        self.timer_producao.stop()
    def watch_pause_liberacao(self):
        self.timer_liberacao.stop()
    def watch_pause_parada(self):
        self.timer_parada.stop()
    def watch_pause_intervalo(self):
        self.timer_intervalo.stop()
    def watch_pause_try_out(self):
        self.timer_try_out.stop()
    def watch_pause_treinamento(self):
        self.timer_treinamento.stop()

    def stop_watch(self):
        self.timer_producao.stop()
        self.timer_liberacao.stop()
        self.timer_parada.stop()
        self.timer_intervalo.stop()
        self.timer_try_out.stop()
        self.timer_treinamento.stop()
        self.mscounter_producao = 0
        self.mscounter_liberacao = 0
        self.mscounter_parada = 0
        self.mscounter_parada02 = 0
        self.mscounter_parada03 = 0
        self.mscounter_parada04 = 0
        self.mscounter_parada05 = 0
        self.mscounter_parada06 = 0
        self.mscounter_parada07 = 0
        self.mscounter_parada08 = 0
        self.mscounter_parada09 = 0
        self.mscounter_parada11 = 0
        self.mscounter_parada12 = 0
        self.mscounter_parada13 = 0
        self.mscounter_parada15 = 0
        self.mscounter_parada18 = 0
        self.mscounter_intervalo = 0
        self.mscounter_try_out = 0
        self.mscounter_treinamento = 0
        self.producao = True
        self.liberacao = True
        self.parada = True
        self.intervalo = True
        self.try_out = True
        self.treinamento = True
        self.showLCD()

#Função para o System Tray Icon

def trayWaiting():
    while SysTrayIcon.SysTrayIcon.WINDOW == True:
        time.sleep(3)
    else:
        SysTrayIcon.SysTrayIcon.WINDOW = True
        janela.show()
        sys.exit(app.exec_())

#Código para gerar gui
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.aboutToQuit.connect(trayWaiting) #Executa o trayWaiting quando a Janela for fechada
    janela = MainWindow()
    janela.show()
    sys.exit(app.exec_())
