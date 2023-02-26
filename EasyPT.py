import datetime
import sys
import webbrowser
import json
from xml.etree.ElementTree import tostring
import PyQt5
from PyQt5 import QtCore, QtGui, QtWidgets
from designEasyPT import *
from designEasyTBTCategoriasDePerigos import *
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtWidgets import QMainWindow, QApplication, QInputDialog, QMessageBox, QCompleter
from docxtpl import DocxTemplate
import ctypes
from pkg_resources import resource_filename
import os
import win32print
import win32api
import time
import shutil


if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    
class SecondaryWindow(QMainWindow, Ui_SecondaryWindow):
    def __init__(self, parent=None, fields=None):
        super().__init__(parent)
        super().setupUi(self)
        self.fields = fields
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.oldPos = self.pos() #poder mover a tela clicando e arrastando em qualquer lugar
        self.setWindowTitle('Categorias de Perigo')
        self.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'fflaranjaSemFundo.ico'))))
        self.btnEscape.clicked.connect(self.fecha)
        self.btnConfig.clicked.connect(self.minimiza)
        self.btnLogo.clicked.connect(self.logo)
        self.btnOk.clicked.connect(self.send_information) 
        
        self.checkBoxBioTBT.setIcon(QtGui.QIcon('./icones/Biologico-removebg-preview.png'))    
        self.checkBoxCalorTBT.setIcon(QtGui.QIcon('./icones/calo-removebg-preview.png'))    
        self.checkBoxClimaTBT.setIcon(QtGui.QIcon('./icones/condicoes-removebg-preview.png'))    
        self.checkBoxDerramamentoTBT.setIcon(QtGui.QIcon('./icones/derramame-removebg-preview.png'))    
        self.checkBoxEleTBT.setIcon(QtGui.QIcon('./icones/eletrico-removebg-preview.png'))    
        self.checkBoxEntrSaiTBT.setIcon(QtGui.QIcon('./icones/entradaSaida-removebg-preview.png'))    
        self.checkBoxEquipPeriTBT.setIcon(QtGui.QIcon('./icones/equipPerigoso-removebg-preview.png'))    
        self.checkBoxEscorregoesTBT.setIcon(QtGui.QIcon('./icones/escorrego-removebg-preview.png'))    
        self.checkBoxEquipDeIcaTBT.setIcon(QtGui.QIcon('./icones/EquipDeIcamento-removebg-preview.png'))    
        self.checkBoxIluTBT.setIcon(QtGui.QIcon('./icones/ilumi-removebg-preview.png'))    
        self.checkBoxLocElevTBT.setIcon(QtGui.QIcon('./icones/trabalhoLocaisElevados-removebg-preview.png'))    
        self.checkBoxManuseioTBT.setIcon(QtGui.QIcon('./icones/manuseio-removebg-preview.png'))    
        self.checkBoxPessoaFamiTBT.setIcon(QtGui.QIcon('./icones/PessoalNFamili-removebg-preview.png'))    
        self.checkBoxOutraTBT.setIcon(QtGui.QIcon('./icones/outra-removebg-preview.png'))    
        self.checkBoxFrioTBT.setIcon(QtGui.QIcon('./icones/frio-removebg-preview.png'))    
        self.checkBoxRadiacaoTBT.setIcon(QtGui.QIcon('./icones/radiacao-removebg-preview.png'))    
        self.checkBoxRuidosTBT.setIcon(QtGui.QIcon('./icones/ruidos-removebg-preview.png'))    
        self.checkBoxQuedaTBT.setIcon(QtGui.QIcon('./icones/queda-removebg-preview.png'))    
        self.checkBoxPressaoTBT.setIcon(QtGui.QIcon('./icones/Pressao-removebg-preview.png'))    
        self.checkBoxSimopsTBT.setIcon(QtGui.QIcon('./icones/simops-removebg-preview.png'))    
        self.checkBoxVibracaoTBT.setIcon(QtGui.QIcon('./icones/vibracao-removebg-preview.png'))    
        self.checkBoxMovObjTBT.setIcon(QtGui.QIcon('./icones/MovObj-removebg-preview.png')) 
        self.checkBoxSubsPeriTBT.setIcon(QtGui.QIcon('./icones/substPerigosas-removebg-preview.png')) 
     
        
    ###Funções Gerais da tela
        
    def logo(self):
        webbrowser.open('https://github.com/Soulbope')  


    def fecha(self):        
        # escolha = ('Salvar', 'Apenas Fechar')
        # item, ok = QInputDialog.getItem(self, "Salvar", "Deseja salvar anter de fechar?", escolha, 0, False)
        
        # if ok and (item=='Salvar'):
        #     self.salvarInputGroup()
        # elif ok and (item=='Apenas Fechar'):
        #     self.close()
    
        self.close()                
        
    def minimiza(self):
        self.showMinimized()                  
        
    # Funções para mover a tela segurando em qualquer lugar    
    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = QPoint (event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()
    
    #Função para salvar com o Enter    
    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Return or e.key() == Qt.Key_Enter:
            self.salvarInputGroup()
            
    ###Funções específicas
    
    def receive_information(self, fields):        
        self.checkBoxBioTBT.setChecked(fields["checkBoxBioTBT"])    
        self.checkBoxCalorTBT.setChecked(fields["checkBoxCalorTBT"])    
        self.checkBoxClimaTBT.setChecked(fields["checkBoxClimaTBT"])    
        self.checkBoxDerramamentoTBT.setChecked(fields["checkBoxDerramamentoTBT"])    
        self.checkBoxEleTBT.setChecked(fields["checkBoxEleTBT"])    
        self.checkBoxEntrSaiTBT.setChecked(fields["checkBoxEntrSaiTBT"])    
        self.checkBoxEquipPeriTBT.setChecked(fields["checkBoxEquipPeriTBT"])    
        self.checkBoxEscorregoesTBT.setChecked(fields["checkBoxEscorregoesTBT"])    
        self.checkBoxEquipDeIcaTBT.setChecked(fields["checkBoxEquipDeIcaTBT"])    
        self.checkBoxIluTBT.setChecked(fields["checkBoxIluTBT"])    
        self.checkBoxLocElevTBT.setChecked(fields["checkBoxLocElevTBT"])    
        self.checkBoxManuseioTBT.setChecked(fields["checkBoxManuseioTBT"])    
        self.checkBoxPessoaFamiTBT.setChecked(fields["checkBoxPessoaFamiTBT"])    
        self.checkBoxOutraTBT.setChecked(fields["checkBoxOutraTBT"])    
        self.checkBoxFrioTBT.setChecked(fields["checkBoxFrioTBT"])    
        self.checkBoxRadiacaoTBT.setChecked(fields["checkBoxRadiacaoTBT"])    
        self.checkBoxRuidosTBT.setChecked(fields["checkBoxRuidosTBT"])    
        self.checkBoxQuedaTBT.setChecked(fields["checkBoxQuedaTBT"])    
        self.checkBoxPressaoTBT.setChecked(fields["checkBoxPressaoTBT"])    
        self.checkBoxSimopsTBT.setChecked(fields["checkBoxSimopsTBT"])    
        self.checkBoxVibracaoTBT.setChecked(fields["checkBoxVibracaoTBT"])    
        self.checkBoxMovObjTBT.setChecked(fields["checkBoxMovObjTBT"]) 
        self.checkBoxSubsPeriTBT.setChecked(fields["checkBoxSubsPeriTBT"])
    
    def send_information(self):
        try:
            information = {"checkBoxClimaTBT": self.checkBoxClimaTBT.isChecked(),
                        "checkBoxDerramamentoTBT": self.checkBoxDerramamentoTBT.isChecked(),
                        "checkBoxEleTBT": self.checkBoxEleTBT.isChecked(),
                        "checkBoxEntrSaiTBT": self.checkBoxEntrSaiTBT.isChecked(),
                        "checkBoxEquipPeriTBT": self.checkBoxEquipPeriTBT.isChecked(),
                        "checkBoxEscorregoesTBT": self.checkBoxEscorregoesTBT.isChecked(),
                        "checkBoxEquipDeIcaTBT": self.checkBoxEquipDeIcaTBT.isChecked(),
                        "checkBoxIluTBT": self.checkBoxIluTBT.isChecked(),
                        "checkBoxLocElevTBT": self.checkBoxLocElevTBT.isChecked(),
                        "checkBoxManuseioTBT": self.checkBoxManuseioTBT.isChecked(),
                        "checkBoxPessoaFamiTBT": self.checkBoxPessoaFamiTBT.isChecked(),
                        "checkBoxFrioTBT": self.checkBoxFrioTBT.isChecked(),
                        "checkBoxRadiacaoTBT": self.checkBoxRadiacaoTBT.isChecked(),
                        "checkBoxRuidosTBT": self.checkBoxRuidosTBT.isChecked(),
                        "checkBoxPressaoTBT": self.checkBoxPressaoTBT.isChecked(),
                        "checkBoxSimopsTBT": self.checkBoxSimopsTBT.isChecked(),
                        "checkBoxVibracaoTBT": self.checkBoxVibracaoTBT.isChecked(),
                        "checkBoxMovObjTBT": self.checkBoxMovObjTBT.isChecked(),
                        "checkBoxSubsPeriTBT": self.checkBoxSubsPeriTBT.isChecked(),
                        "checkBoxBioTBT": self.checkBoxBioTBT.isChecked(),
                        "checkBoxOutraTBT": self.checkBoxOutraTBT.isChecked(),
                        "checkBoxQuedaTBT": self.checkBoxQuedaTBT.isChecked(),
                        "checkBoxCalorTBT": self.checkBoxCalorTBT.isChecked(),
                        }
            self.parent().receive_information(information)
            self.close()
        except Exception as e:
            print(f'O erro é: {e}') 
    


class Novo(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.oldPos = self.pos() #poder mover a tela clicando e arrastando em qualquer lugar
        self.combobox() #preencher combobox
        self.setWindowTitle('EasyPT')
        self.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
        self.btnEscape.clicked.connect(self.fecha)    
        self.btnSalvar.clicked.connect(self.salvarInputGroup)
        self.btnConfig.clicked.connect(self.minimiza)
        self.btnChoosePrinter.clicked.connect(self.escolherImpressora)
        self.btnImprimir.clicked.connect(self.saveAndPrint)
        self.btnAbrir.clicked.connect(self.escolherInputGroup)
        self.btnLogo.clicked.connect(self.logo)
        self.btnVersion.clicked.connect(self.versionButton)
        self.btnCatPerigos.clicked.connect(self.open_secondary_window)
        
        self.inputs = {}
        self.file_name = 'inputs.json'
        self.btnVersion.setText('Feito por WAGNER OLIVEIRA - Ver. 1.3.3')#26/02/2023        
        
        #popup de IMPRESSÂO
        self.imsg = QMessageBox()
        self.imsg.setWindowTitle("Imprimindo")
        self.imsg.setText("AGUARDE IMPRESSÃO...")
        self.imsg.setIcon(QMessageBox.Information)
        self.imsg.setStandardButtons(QMessageBox.Ok)
        self.imsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
        
        #popup de Preparando arquivos
        self.pmsg = QMessageBox()
        self.pmsg.setWindowTitle("Preparando")
        self.pmsg.setText("AGUARDE PREPARAÇÃO...")
        self.pmsg.setIcon(QMessageBox.Information)
        self.pmsg.setStandardButtons(QMessageBox.Ok)
        self.pmsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
                
        #popup de ERRO
        self.emsg = QMessageBox()
        self.emsg.setWindowTitle("Erro")
        self.emsg.setText("Ocorreu um erro de execução")
        self.emsg.setIcon(QMessageBox.Critical)
        self.emsg.setStandardButtons(QMessageBox.Ok)
        self.emsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
        
        #popup de Update
        self.umsg = QMessageBox()
        self.umsg.setWindowTitle("Foi mal!")
        self.umsg.setText("Entrei de férias antes de conseguir adicionar esse botão, mas na volta eu faço!")
        self.umsg.setIcon(QMessageBox.Information)
        self.umsg.setStandardButtons(QMessageBox.Ok)
        self.umsg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
              
        #popup de IMPRESSÂO
        self.msg2 = QMessageBox()
        self.msg2.setWindowTitle("Atenção")
        self.msg2.setText("Arquivo de mesmo nome, favor alterar!")
        self.msg2.setIcon(QMessageBox.Warning)
        self.msg2.setStandardButtons(QMessageBox.Ok)
        self.msg2.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
        
        # Algumas coisas já padrão
        self.dateEditHoje.setText(datetime.date.today().strftime("%d/%m/%Y"))
        self.timeEditInicio.setText(datetime.datetime.now().strftime("%H:%M"))
        self.inputDepartamento.setText('PIPELAY')
        self.inputLuvas.setText('Vaqueta / Maxflex')
        self.inputPrecaucoes.setText('Área limpa, organizada e inspecionada')
        self.inputHoraInspecao.setText('4H')
        self.inputUHF.setText('4')
        self.inputChecklist.setText('DBR-TE-CK-007')
        self.checkBoxChecklist.setChecked(True)
        self.checkBoxBioTBT = False
        self.checkBoxCalorTBT = False
        self.checkBoxClimaTBT = False
        self.checkBoxDerramamentoTBT = False
        self.checkBoxEleTBT = False
        self.checkBoxEntrSaiTBT = False
        self.checkBoxEquipPeriTBT = False
        self.checkBoxEscorregoesTBT = False
        self.checkBoxEquipDeIcaTBT = False
        self.checkBoxIluTBT = False
        self.checkBoxLocElevTBT = False
        self.checkBoxManuseioTBT = False
        self.checkBoxPessoaFamiTBT = False
        self.checkBoxOutraTBT = False
        self.checkBoxFrioTBT = False
        self.checkBoxRadiacaoTBT = False
        self.checkBoxRuidosTBT = False
        self.checkBoxQuedaTBT = False
        self.checkBoxPressaoTBT = False
        self.checkBoxSimopsTBT = False
        self.checkBoxVibracaoTBT = False
        self.checkBoxMovObjTBT = False
        self.checkBoxSubsPeriTBT = False
        
        Namecompleter = QCompleter(['ALEXANDRO LIMA DOS SANTOS', 'ANDRE AUGUSTO DOS SANTOS NASCIMENTO', 'BERNARDO BEHNKEN PIMENTA', 'BRUNO PORTO DA SILVA', 'CARLOS HENRIQUE NASCIMENTO DA COSTA', 'CARLOS ROBERTO REIS JUNIOR', 'CHRISTIAN HENRI JOSEPH HERTAY', 'DAMIANA FERREIRA DE LIMA', 'DIEGO SOUZA JULIO', 'DIOGO HENRIQUE ANDRADE CORREA', 'DIONI CANTELLI', 'FABIO HENRIQUE SIMAS ABREU', 'FABIO PAULINO ALVES SOARES', 'FREDERIC PHILIPPE BOUDOUX', 'GIOVANI FERREIRA DA CONCEICAO', 'GLEISON LUIZ DA SILVA', 'HENRIQUE FERNANDES DE ARAUJO', 'JONAS DE ALMEIDA SANTOS', 'LEANDRO DE LIMA DA SILVA', 'LUCAS EDUARDO DE CAIO CARVALHO', 'MARCEL NOGUEIRA PINHO', 'LUCAS LOPES DO NASCIMENTO', 'MARCELO MENDONÇA DOS SANTOS', 'MARCOS ADRIANI SANTOS DE OLIVEIRA', 'MARCOS AURELIO MENDES PIMENTA', 'MARCOS MOISES PEREIRA DOS SANTOS', 'MATHEUS GONCALVES LEITE', 'NORBERT HUGUES ZUNINO', 'RAFAEL GOMES MARTINS', 'ROBERTO ALEXANDRE FERREIRA', 'ROBERTO ALVES JUNIOR', 'RODRIGO JOSE DOS SANTOS RIBEIRO', 'ROGERIO VIEIRA GUIMARAES', 'ROMULO FERREIRA DE ALVARENGA', 'RUAN KAIQUE ANTUNES ANDRE', 'SHIRLEI DA SILVA DE CASTRO', 'THIAGO BRAGA DA PAIXAO', 'WAGNER LENI DE OLIVEIRA JUNIOR', 'WAGNER VIANA VIEIRA', 'WANDERSON DE ASSIS MARTINS', 'WESLEY LOROZA CORREA', 'WILLIAN PIMENTEL BATISTA'])
        Namecompleter.setCaseSensitivity(Qt.CaseInsensitive)
        
        for input in [self.inputExecutante1, self.inputExecutante2, self.inputExecutante3, self.inputExecutante4, self.inputExecutante5,
                      self.inputResgatista1, self.inputResgatista2, self.inputResgatista3, self.inputResgatista4, self.inputObservadorVigia,
                      self.inputAutoridade
                      ]:
            input.setCompleter(Namecompleter)
        
    def open_secondary_window(self):
        try:
            fields = {
                'checkBoxBioTBT' : self.checkBoxBioTBT,
                'checkBoxCalorTBT' : self.checkBoxCalorTBT,
                'checkBoxClimaTBT' : self.checkBoxClimaTBT,
                'checkBoxDerramamentoTBT' : self.checkBoxDerramamentoTBT,
                'checkBoxEleTBT' : self.checkBoxEleTBT,
                'checkBoxEntrSaiTBT' : self.checkBoxEntrSaiTBT,
                'checkBoxEquipPeriTBT' : self.checkBoxEquipPeriTBT,
                'checkBoxEscorregoesTBT' : self.checkBoxEscorregoesTBT,
                'checkBoxEquipDeIcaTBT' : self.checkBoxEquipDeIcaTBT,
                'checkBoxIluTBT' : self.checkBoxIluTBT,
                'checkBoxLocElevTBT' : self.checkBoxLocElevTBT,
                'checkBoxManuseioTBT' : self.checkBoxManuseioTBT,
                'checkBoxPessoaFamiTBT' : self.checkBoxPessoaFamiTBT,
                'checkBoxOutraTBT' : self.checkBoxOutraTBT,
                'checkBoxFrioTBT' : self.checkBoxFrioTBT,
                'checkBoxRadiacaoTBT' : self.checkBoxRadiacaoTBT,
                'checkBoxRuidosTBT' : self.checkBoxRuidosTBT,
                'checkBoxQuedaTBT' : self.checkBoxQuedaTBT,
                'checkBoxPressaoTBT' : self.checkBoxPressaoTBT,
                'checkBoxSimopsTBT' : self.checkBoxSimopsTBT,
                'checkBoxVibracaoTBT' : self.checkBoxVibracaoTBT,
                'checkBoxMovObjTBT' : self.checkBoxMovObjTBT,
                'checkBoxSubsPeriTBT' : self.checkBoxSubsPeriTBT,
            }
            self.secondary_window = SecondaryWindow(parent=self, fields=fields)
            self.secondary_window.receive_information(fields=fields)
            self.secondary_window.show()
        except Exception as e:
            print(f'o erro foi: {e}')
            
#recebe a informação da segunda tela                 
    def receive_information(self, information):        
        self.checkBoxBioTBT = (information["checkBoxBioTBT"])    
        self.checkBoxCalorTBT = (information["checkBoxCalorTBT"])    
        self.checkBoxClimaTBT = (information["checkBoxClimaTBT"])    
        self.checkBoxDerramamentoTBT = (information["checkBoxDerramamentoTBT"])    
        self.checkBoxEleTBT = (information["checkBoxEleTBT"])    
        self.checkBoxEntrSaiTBT = (information["checkBoxEntrSaiTBT"])    
        self.checkBoxEquipPeriTBT = (information["checkBoxEquipPeriTBT"])    
        self.checkBoxEscorregoesTBT = (information["checkBoxEscorregoesTBT"])    
        self.checkBoxEquipDeIcaTBT = (information["checkBoxEquipDeIcaTBT"])    
        self.checkBoxIluTBT = (information["checkBoxIluTBT"])    
        self.checkBoxLocElevTBT = (information["checkBoxLocElevTBT"])    
        self.checkBoxManuseioTBT = (information["checkBoxManuseioTBT"])    
        self.checkBoxPessoaFamiTBT = (information["checkBoxPessoaFamiTBT"])    
        self.checkBoxOutraTBT = (information["checkBoxOutraTBT"])    
        self.checkBoxFrioTBT = (information["checkBoxFrioTBT"])    
        self.checkBoxRadiacaoTBT = (information["checkBoxRadiacaoTBT"])    
        self.checkBoxRuidosTBT = (information["checkBoxRuidosTBT"])    
        self.checkBoxQuedaTBT = (information["checkBoxQuedaTBT"])    
        self.checkBoxPressaoTBT = (information["checkBoxPressaoTBT"])    
        self.checkBoxSimopsTBT = (information["checkBoxSimopsTBT"])    
        self.checkBoxVibracaoTBT = (information["checkBoxVibracaoTBT"])    
        self.checkBoxMovObjTBT = (information["checkBoxMovObjTBT"]) 
        self.checkBoxSubsPeriTBT = (information["checkBoxSubsPeriTBT"])
   
    def logo(self):
        webbrowser.open('https://github.com/Soulbope')  
        
    def versionButton(self):
        webbrowser.open('https://github.com/Soulbope/EasyPT/tree/main#readme')  

    def fecha(self):        
        escolha = ('Salvar', 'Apenas Fechar')
        item, ok = QInputDialog.getItem(self, "Salvar", "Deseja salvar anter de fechar?", escolha, 0, False)
        
        if ok and (item=='Salvar'):
            self.salvarInputGroup()
        elif ok and (item=='Apenas Fechar'):
            self.close()
    
        self.close()                
        
    def minimiza(self):
        self.showMinimized()                  
        
    # Funções para mover a tela segurando em qualquer lugar    
    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = QPoint (event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()
    
    #Função para salvar com o Enter    
    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Return or e.key() == Qt.Key_Enter:
            self.salvar()                 
        
    
    #Função para add as opções do Combobox
    def combobox(self):
        jobType = ['Trab. a quente','Trab. em Sist. de hidrocarbonetos','Isolamento','Operação de poço','Trab. em altura','Espaço confinado','Substâncias perigosas','Outros']
        for i in jobType:
            self.comboBoxJobType.addItem(i)
        
    def saveAndPrint(self):
        try:
            #Específico Para a TBT            
            inputRiscos = self.inputRiscos.text()#            
            inputGatilhos = self.inputGatilhos.text()# 
            checkBoxErgonomia = "Sim" if self.checkBoxErgonomia.isChecked() else "Não"
            checkBoxColuna = "Sim" if self.checkBoxColuna.isChecked() else "Não"     
            
            #Para a PT
            jobType = self.comboBoxJobType.currentText()#
            inputDetentor = self.inputDetentor.text()#
            inputLocSenso = self.inputLocSenso.text()#
            inputEpiEspecifico = self.inputEpiEspecifico.text()#            
            inputUHF = self.inputUHF.text()#
            inputFerramentas = self.inputFerramentas.text()#
            inputJRA = self.inputJRA.text()#
            textEditJobDescription = self.textEditJobDescription.toPlainText()#
            inputArea = self.inputArea.text()#                
            inputDeck = self.inputDeck.text()# 
            inputAnexos = self.inputAnexos.text()#
            inputChecklist = self.inputChecklist.text()#
            inputHoraInspecao = self.inputHoraInspecao.text()#
            inputResponsavelInspecao = self.inputResponsavelInspecao.text()#
            inputIsolamentoEletrico = self.inputIsolamentoEletrico.text()#
            inputPrecaucoes = self.inputPrecaucoes.text()#
            inputLuvas = self.inputLuvas.text()#
            inputResgatista1 = self.inputResgatista1.text()#
            inputResgatista2 = self.inputResgatista2.text()#
            inputResgatista3 = self.inputResgatista3.text()#
            inputResgatista4 = self.inputResgatista4.text()#
            inputObservadorVigia = self.inputObservadorVigia.text()#
            inputAutoridade = self.inputAutoridade.text()#
            inputTecnico = self.inputTecnico.text()#
            dateEditHoje = self.dateEditHoje.text()#
            timeEditInicio = self.timeEditInicio.text()#
            timeEditFinal = self.timeEditFinal.text()#
            checkBoxDetectorGas = self.checkBoxDetectorGas.isChecked()#
            checkBoxIsoMec = self.checkBoxIsoMec.isChecked()#
            checkBoxIsoEle = self.checkBoxIsoEle.isChecked()#
            checkBoxExtintor = self.checkBoxExtintor.isChecked()#
            checkBoxMaquinaSolda = self.checkBoxMaquinaSolda.isChecked()#
            checkBoxComunicacao = self.checkBoxComunicacao.isChecked()#
            checkBoxDrenos = self.checkBoxDrenos.isChecked()#
            checkBoxBarreira = self.checkBoxBarreira.isChecked()#
            checkBoxCooperar = self.checkBoxCooperar.isChecked()#
            checkBoxTrabalhoEmAltura = self.checkBoxTrabalhoEmAltura.isChecked()#
            checkBoxFispq = self.checkBoxFispq.isChecked()#
            checkBoxChecklist = self.checkBoxChecklist.isChecked()#
            checkBoxIcamento = self.checkBoxIcamento.isChecked()#
            checkBoxConfinado = self.checkBoxConfinado.isChecked()#
            checkBoxEpiEspecial = self.checkBoxEpiEspecial.isChecked()#
            checkBoxOutros = self.checkBoxOutros.isChecked()
            inputExecutante1 = self.inputExecutante1.text()#
            inputExecutante2 = self.inputExecutante2.text()#
            inputExecutante3 = self.inputExecutante3.text()#
            inputExecutante4 = self.inputExecutante4.text()#
            inputExecutante5 = self.inputExecutante5.text()#
            inputDepartamento = self.inputDepartamento.text()#
            
            #secondary Window            
            checkBoxBioTBT = "x" if self.checkBoxBioTBT else ""
            checkBoxCalorTBT = "x" if self.checkBoxCalorTBT else ""  
            checkBoxClimaTBT = "x" if self.checkBoxClimaTBT else ""  
            checkBoxDerramamentoTBT = "x" if self.checkBoxDerramamentoTBT else ""
            checkBoxEleTBT = "x" if self.checkBoxEleTBT else ""
            checkBoxEntrSaiTBT = "x" if self.checkBoxEntrSaiTBT else ""    
            checkBoxEquipPeriTBT = "x" if self.checkBoxEquipPeriTBT else ""    
            checkBoxEscorregoesTBT = "x" if self.checkBoxEscorregoesTBT else ""
            checkBoxEquipDeIcaTBT = "x" if self.checkBoxEquipDeIcaTBT else ""  
            checkBoxIluTBT = "x" if self.checkBoxIluTBT else ""
            checkBoxLocElevTBT = "x" if self.checkBoxLocElevTBT else ""    
            checkBoxManuseioTBT = "x" if self.checkBoxManuseioTBT else ""  
            checkBoxPessoaFamiTBT = "x" if self.checkBoxPessoaFamiTBT else ""
            checkBoxOutraTBT = "x" if self.checkBoxOutraTBT else ""  
            checkBoxFrioTBT = "x" if self.checkBoxFrioTBT else "" 
            checkBoxRadiacaoTBT = "x" if self.checkBoxRadiacaoTBT else ""   
            checkBoxRuidosTBT = "x" if self.checkBoxRuidosTBT else ""   
            checkBoxQuedaTBT = "x" if self.checkBoxQuedaTBT else ""  
            checkBoxPressaoTBT = "x" if self.checkBoxPressaoTBT else ""    
            checkBoxSimopsTBT = "x" if self.checkBoxSimopsTBT else ""   
            checkBoxVibracaoTBT = "x" if self.checkBoxVibracaoTBT else ""    
            checkBoxMovObjTBT = "x" if self.checkBoxMovObjTBT else ""
            checkBoxSubsPeriTBT = "x" if self.checkBoxSubsPeriTBT else ""
            
            
            #Para colocar a opção do Combobox com 'x' no Doc    
            if jobType == 'Trab. a quente':
                type = 'jtTrabalhoQuente'
            elif jobType == 'Trab. em Sist. de hidrocarbonetos':
                type = 'jtHidro'
            elif jobType == 'Isolamento':
                type = 'jtIsolamento'
            elif jobType == 'Operação de poço':
                type = 'jtPoco'
            elif jobType == 'Trab. em altura':
                type = 'jtAltura'
            elif jobType == 'Espaço confinado':
                type = 'jtConfinado'
            elif jobType == 'Substâncias perigosas':
                type = 'jtSubsPerigosas'
            elif jobType == 'Outros':
                type = 'jtOutros'
                
                    
            doc = DocxTemplate(os.path.join(resource_filename(__name__, 'pttemplat1.docx')))     
            context = {
                'inputDetentor' : inputDetentor,
                'inputEpiEspecifico' : inputEpiEspecifico,
                'inputLocSenso' : inputLocSenso,
                'inputUHF' : inputUHF,
                'inputFerramentas' : inputFerramentas,
                'inputJRA' : inputJRA,
                'textEditJobDescription' : textEditJobDescription,
                'inputArea' : inputArea,                
                'inputDeck' : inputDeck,
                'inputAnexos' : inputAnexos,
                'inputHoraInspecao' : inputHoraInspecao,
                'inputResponsavelInspecao' : inputResponsavelInspecao,
                'inputIsolamentoEletrico' : inputIsolamentoEletrico,
                'inputPrecaucoes' : inputPrecaucoes,
                'inputLuvas' : inputLuvas,
                'inputResgatista1' : inputResgatista1,
                'inputResgatista2' : inputResgatista2,
                'inputResgatista3' : inputResgatista3,
                'inputResgatista4' : inputResgatista4,
                'inputObservadorVigia' : inputObservadorVigia,
                'inputChecklist' : inputChecklist,
                'inputAutoridade' : inputAutoridade,
                'inputTecnico' : inputTecnico,
                'dateEditHoje' : dateEditHoje,
                'timeEditInicio' : timeEditInicio,
                'timeEditFinal' : timeEditFinal,
                'inputExecutante1' : inputExecutante1,
                'inputExecutante2' : inputExecutante2,
                'inputExecutante3' : inputExecutante3,
                'inputExecutante4' : inputExecutante4,
                'inputExecutante5' : inputExecutante5,
                'inputDepartamento' : inputDepartamento,
                type : 'x',
                'checkBoxDetectorGas' : 'x' if checkBoxDetectorGas == True else '',
                'checkBoxIsoMec' : 'x' if checkBoxIsoMec == True else '',
                'checkBoxIsoEle' : 'x' if checkBoxIsoEle == True else '',
                'checkBoxIsoEle2' : 'x' if checkBoxIsoEle == True else '',
                'checkBoxExtintor' : 'x' if checkBoxExtintor == True else '',
                'checkBoxMaquinaSolda' : 'x' if checkBoxMaquinaSolda == True else '',
                'checkBoxComunicacao' : 'x' if checkBoxComunicacao == True else '',
                'checkBoxDrenos' : 'x' if checkBoxDrenos == True else '',
                'checkBoxBarreira' : 'x' if checkBoxBarreira == True else '',
                'checkBoxCooperar' : 'x' if checkBoxCooperar == True else '',
                'checkBoxTrabalhoEmAltura' : 'x' if checkBoxTrabalhoEmAltura == True else '',
                'checkBoxFispq' : 'x' if checkBoxFispq == True else '',
                'checkBoxChecklist' : 'x' if checkBoxChecklist == True else '',
                'checkBoxIcamento' : 'x' if checkBoxIcamento == True else '',
                'checkBoxConfinado' : 'x' if checkBoxConfinado == True else '',
                'checkBoxEpiEspecial' : 'x' if checkBoxEpiEspecial == True else '',
                'checkBoxOutros' : 'x' if checkBoxOutros == True else ''
            }
            doc.render(context)
            
            doctbt = DocxTemplate(os.path.join(resource_filename(__name__, 'tbttemplate.docx')))     
            context = {
                'inputDetentor' : inputDetentor,
                'inputRiscos' : inputRiscos,
                'inputGatilhos' : inputGatilhos,
                'checkBoxColuna' : checkBoxColuna,
                'checkBoxErgonomia' : checkBoxErgonomia,
                'textEditJobDescription' : textEditJobDescription,
                'inputArea' : inputArea,
                'inputResgatista1' : inputResgatista1,
                'inputResgatista2' : inputResgatista2,
                'inputResgatista3' : inputResgatista3,
                'inputResgatista4' : inputResgatista4,
                'inputObservadorVigia' : inputObservadorVigia,
                'inputAutoridade' : inputAutoridade,
                'dateEditHoje' : dateEditHoje,
                'inputExecutante1' : inputExecutante1,
                'inputExecutante2' : inputExecutante2,
                'inputExecutante3' : inputExecutante3,
                'inputExecutante4' : inputExecutante4,
                'inputExecutante5' : inputExecutante5,
                'checkBoxBioTBT' : checkBoxBioTBT,
                'checkBoxCalorTBT' : checkBoxCalorTBT,
                'checkBoxClimaTBT' : checkBoxClimaTBT,
                'checkBoxDerramamentoTBT' : checkBoxDerramamentoTBT,
                'checkBoxEleTBT' : checkBoxEleTBT,
                'checkBoxEntrSaiTBT' : checkBoxEntrSaiTBT,
                'checkBoxEquipPeriTBT' : checkBoxEquipPeriTBT,
                'checkBoxEscorregoesTBT' : checkBoxEscorregoesTBT,
                'checkBoxEquipDeIcaTBT' : checkBoxEquipDeIcaTBT,
                'checkBoxIluTBT' : checkBoxIluTBT,
                'checkBoxLocElevTBT' : checkBoxLocElevTBT,
                'checkBoxManuseioTBT' : checkBoxManuseioTBT,
                'checkBoxPessoaFamiTBT' : checkBoxPessoaFamiTBT,
                'checkBoxOutraTBT' : checkBoxOutraTBT,
                'checkBoxFrioTBT' : checkBoxFrioTBT,
                'checkBoxRadiacaoTBT' : checkBoxRadiacaoTBT,
                'checkBoxRuidosTBT' : checkBoxRuidosTBT,
                'checkBoxQuedaTBT' : checkBoxQuedaTBT,
                'checkBoxPressaoTBT' : checkBoxPressaoTBT,
                'checkBoxSimopsTBT' : checkBoxSimopsTBT,
                'checkBoxVibracaoTBT' : checkBoxVibracaoTBT,
                'checkBoxMovObjTBT' : checkBoxMovObjTBT,
                'checkBoxSubsPeriTBT' : checkBoxSubsPeriTBT,
            }
            doctbt.render(context)
            
            
            doccheck = DocxTemplate(os.path.join(resource_filename(__name__, 'checktemplate.docx')))     
            context = {
                'textEditJobDescription' : textEditJobDescription,    
                'dateEditHoje' : dateEditHoje,
                'timeEditInicio' : timeEditInicio,
                'inputAutoridade' : inputAutoridade,
                'inputDetentor' : inputDetentor,
            }
            doccheck.render(context)
            
            self.pmsg.show()
            QApplication.processEvents()
            
            #cria pasta temporária, salva os arquivos
            pastaTempAtual = os.path.join(resource_filename(__name__, 'temp'))
            os.mkdir(pastaTempAtual) #cria a pasta
            doctbt.save(f'{pastaTempAtual}/tbt.docx')
            doccheck.save(f'{pastaTempAtual}/check.docx')
            try:
                for i in range(3):
                    doc.save(f'{pastaTempAtual}/{i}.docx')
            except Exception as e:
                print(f'O erro ao criar a pasta foi: {e}')
                pass                    
            
            time.sleep(5)
            
            self.pmsg.close()             
            
            printer_handle = win32print.OpenPrinter(win32print.GetDefaultPrinter())
            status = win32print.GetPrinter(printer_handle, 2)['Status']
            while status == win32print.PRINTER_STATUS_BUSY:
                time.sleep(1)
                status = win32print.GetPrinter(printer_handle, 2)['Status']
            win32print.ClosePrinter(printer_handle)     
                
            #imprime todos os arquivos da pasta
            listaPts = os.listdir(pastaTempAtual)
            try:
                self.imsg.show()
                QApplication.processEvents()
                for arquivo in listaPts:
                    time.sleep(5)
                    #win32api.ShellExecute(0, "print", arquivo , None, pastaTempAtual , 0)
                 
            except Exception as e:
                x = self.emsg.setInformativeText(str(e))
                x = self.emsg.exec_()
                print(f'O erro ao inprimir foi: {e}')
                pass
                            
            time.sleep(15)
            
            self.imsg.close()
            QApplication.processEvents()     
            
            time.sleep(15)
            
            
            try:
                shutil.rmtree(pastaTempAtual) #apaga a pasta
            except Exception as e:
                print(f'O erro é: {e}')
        
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            print(f'o erro é: {e}')
            pass 
        
        
    def escolherImpressora(self):
        try:
            lista_impressoras = win32print.EnumPrinters(2)
            
            impressoras = []       
                    
            for infoTotalImpressoras in lista_impressoras:
                impressoras.append(infoTotalImpressoras[2])   
                
            item, ok = QInputDialog.getItem(self, "IMPRESSORA", "Selecionar impressora", impressoras, 0, False)
                    
            if ok and item:
                for impressorass in lista_impressoras:
                    if impressorass[2] == item:
                        impressoraAtual = impressorass[2]
                        
            win32print.SetDefaultPrinter(impressoraAtual)
            
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            print(f'o erro é: {e}')
            pass 
        
        
        
    #Preciso colocar um botão pasa salvar o template que o usuário acabou de preencher (deseja salvar essas informações?)    
    def salvarInputGroup(self):
        try:
            #específico da TBT
            inputRiscos = self.inputRiscos.text()#            
            inputGatilhos = self.inputGatilhos.text()# 
            checkBoxColuna = self.checkBoxColuna.isChecked()#           
            checkBoxErgonomia = self.checkBoxErgonomia.isChecked()#
            
            checkBoxBioTBT = self.checkBoxBioTBT
            checkBoxCalorTBT = self.checkBoxCalorTBT
            checkBoxClimaTBT = self.checkBoxClimaTBT
            checkBoxDerramamentoTBT = self.checkBoxDerramamentoTBT
            checkBoxEleTBT = self.checkBoxEleTBT
            checkBoxEntrSaiTBT = self.checkBoxEntrSaiTBT
            checkBoxEquipPeriTBT = self.checkBoxEquipPeriTBT
            checkBoxEscorregoesTBT = self.checkBoxEscorregoesTBT
            checkBoxEquipDeIcaTBT = self.checkBoxEquipDeIcaTBT
            checkBoxIluTBT = self.checkBoxIluTBT
            checkBoxLocElevTBT = self.checkBoxLocElevTBT
            checkBoxManuseioTBT = self.checkBoxManuseioTBT
            checkBoxPessoaFamiTBT = self.checkBoxPessoaFamiTBT
            checkBoxOutraTBT = self.checkBoxOutraTBT
            checkBoxFrioTBT = self.checkBoxFrioTBT
            checkBoxRadiacaoTBT = self.checkBoxRadiacaoTBT
            checkBoxRuidosTBT = self.checkBoxRuidosTBT
            checkBoxQuedaTBT = self.checkBoxQuedaTBT
            checkBoxPressaoTBT = self.checkBoxPressaoTBT
            checkBoxSimopsTBT = self.checkBoxSimopsTBT
            checkBoxVibracaoTBT = self.checkBoxVibracaoTBT
            checkBoxMovObjTBT = self.checkBoxMovObjTBT
            checkBoxSubsPeriTBT = self.checkBoxSubsPeriTBT
            
            #Para a PT
            jobType = self.comboBoxJobType.currentText()
            inputDetentor = self.inputDetentor.text()#
            inputEpiEspecifico = self.inputEpiEspecifico.text()#
            inputLocSenso = self.inputLocSenso.text()#
            inputUHF = self.inputUHF.text()#
            inputFerramentas = self.inputFerramentas.text()#
            inputJRA = self.inputJRA.text()#
            textEditJobDescription = self.textEditJobDescription.toPlainText()#
            inputArea = self.inputArea.text()#                
            inputDeck = self.inputDeck.text()# 
            inputAnexos = self.inputAnexos.text()#
            inputHoraInspecao = self.inputHoraInspecao.text()#
            inputResponsavelInspecao = self.inputResponsavelInspecao.text()#
            inputIsolamentoEletrico = self.inputIsolamentoEletrico.text()#
            inputPrecaucoes = self.inputPrecaucoes.text()#
            inputLuvas = self.inputLuvas.text()#
            inputResgatista1 = self.inputResgatista1.text()#
            inputResgatista2 = self.inputResgatista2.text()#
            inputResgatista3 = self.inputResgatista3.text()#
            inputResgatista4 = self.inputResgatista4.text()#
            inputObservadorVigia = self.inputObservadorVigia.text()#
            inputAutoridade = self.inputAutoridade.text()#
            inputTecnico = self.inputTecnico.text()#
            dateEditHoje = self.dateEditHoje.text()#
            timeEditInicio = self.timeEditInicio.text()#
            timeEditFinal = self.timeEditFinal.text()#
            checkBoxDetectorGas = self.checkBoxDetectorGas.isChecked()#
            checkBoxIsoMec = self.checkBoxIsoMec.isChecked()#
            checkBoxIsoEle = self.checkBoxIsoEle.isChecked()#
            checkBoxExtintor = self.checkBoxExtintor.isChecked()#
            checkBoxMaquinaSolda = self.checkBoxMaquinaSolda.isChecked()#
            checkBoxComunicacao = self.checkBoxComunicacao.isChecked()#
            checkBoxDrenos = self.checkBoxDrenos.isChecked()#
            checkBoxBarreira = self.checkBoxBarreira.isChecked()#
            checkBoxCooperar = self.checkBoxCooperar.isChecked()#
            checkBoxTrabalhoEmAltura = self.checkBoxTrabalhoEmAltura.isChecked()#
            checkBoxFispq = self.checkBoxFispq.isChecked()#
            checkBoxChecklist = self.checkBoxChecklist.isChecked()#
            checkBoxIcamento = self.checkBoxIcamento.isChecked()#
            checkBoxConfinado = self.checkBoxConfinado.isChecked()#
            checkBoxEpiEspecial = self.checkBoxEpiEspecial.isChecked()#
            checkBoxOutros = self.checkBoxOutros.isChecked()
            inputExecutante1 = self.inputExecutante1.text()#
            inputExecutante2 = self.inputExecutante2.text()#
            inputExecutante3 = self.inputExecutante3.text()#
            inputExecutante4 = self.inputExecutante4.text()#
            inputExecutante5 = self.inputExecutante5.text()#
            inputDepartamento = self.inputDepartamento.text()#
            
                     
            name, _ = QtWidgets.QInputDialog.getText(self, 'Salvar', 'Dê um nome para esse template:')

            if name:
                self.inputs[name] = {
                'jobType' : jobType,
                'inputRiscos' : inputRiscos,
                'inputGatilhos' : inputGatilhos,
                'checkBoxColuna' : checkBoxColuna,
                'checkBoxErgonomia' : checkBoxErgonomia,
                'inputDetentor' : inputDetentor,
                'inputEpiEspecifico' : inputEpiEspecifico,
                'inputLocSenso' : inputLocSenso,
                'inputUHF' : inputUHF,
                'inputFerramentas' : inputFerramentas,
                'inputJRA' : inputJRA,
                'textEditJobDescription' : textEditJobDescription,
                'inputArea' : inputArea,                
                'inputDeck' : inputDeck,
                'inputAnexos' : inputAnexos,
                'inputHoraInspecao' : inputHoraInspecao,
                'inputResponsavelInspecao' : inputResponsavelInspecao,
                'inputIsolamentoEletrico' : inputIsolamentoEletrico,
                'inputPrecaucoes' : inputPrecaucoes,
                'inputLuvas' : inputLuvas,
                'inputResgatista1' : inputResgatista1,
                'inputResgatista2' : inputResgatista2,
                'inputResgatista3' : inputResgatista3,
                'inputResgatista4' : inputResgatista4,
                'inputObservadorVigia' : inputObservadorVigia,
                'inputAutoridade' : inputAutoridade,
                'inputTecnico' : inputTecnico,
                'dateEditHoje' : dateEditHoje,
                'timeEditInicio' : timeEditInicio,
                'timeEditFinal' : timeEditFinal,
                'inputExecutante1' : inputExecutante1,
                'inputExecutante2' : inputExecutante2,
                'inputExecutante3' : inputExecutante3,
                'inputExecutante4' : inputExecutante4,
                'inputExecutante5' : inputExecutante5,
                'inputDepartamento' : inputDepartamento,
                'checkBoxDetectorGas' : checkBoxDetectorGas,
                'checkBoxIsoMec' : checkBoxIsoMec,
                'checkBoxIsoEle' : checkBoxIsoEle,
                'checkBoxIsoEle2' : checkBoxIsoEle,
                'checkBoxExtintor' : checkBoxExtintor,
                'checkBoxMaquinaSolda' : checkBoxMaquinaSolda,
                'checkBoxComunicacao' : checkBoxComunicacao,
                'checkBoxDrenos' : checkBoxDrenos,
                'checkBoxBarreira' : checkBoxBarreira,
                'checkBoxCooperar' : checkBoxCooperar,
                'checkBoxTrabalhoEmAltura' : checkBoxTrabalhoEmAltura,
                'checkBoxFispq' : checkBoxFispq,
                'checkBoxChecklist' : checkBoxChecklist,
                'checkBoxIcamento' : checkBoxIcamento,
                'checkBoxConfinado' : checkBoxConfinado,
                'checkBoxEpiEspecial' : checkBoxEpiEspecial,
                'checkBoxOutros' : checkBoxOutros,
                'checkBoxBioTBT' : checkBoxBioTBT,
                'checkBoxCalorTBT' : checkBoxCalorTBT,
                'checkBoxClimaTBT' : checkBoxClimaTBT,
                'checkBoxDerramamentoTBT' : checkBoxDerramamentoTBT,
                'checkBoxEleTBT' : checkBoxEleTBT,
                'checkBoxEntrSaiTBT' : checkBoxEntrSaiTBT,
                'checkBoxEquipPeriTBT' : checkBoxEquipPeriTBT,
                'checkBoxEscorregoesTBT' : checkBoxEscorregoesTBT,
                'checkBoxEquipDeIcaTBT' : checkBoxEquipDeIcaTBT,
                'checkBoxIluTBT' : checkBoxIluTBT,
                'checkBoxLocElevTBT' : checkBoxLocElevTBT,
                'checkBoxManuseioTBT' : checkBoxManuseioTBT,
                'checkBoxPessoaFamiTBT' : checkBoxPessoaFamiTBT,
                'checkBoxOutraTBT' : checkBoxOutraTBT,
                'checkBoxFrioTBT' : checkBoxFrioTBT,
                'checkBoxRadiacaoTBT' : checkBoxRadiacaoTBT,
                'checkBoxRuidosTBT' : checkBoxRuidosTBT,
                'checkBoxQuedaTBT' : checkBoxQuedaTBT,
                'checkBoxPressaoTBT' : checkBoxPressaoTBT,
                'checkBoxSimopsTBT' : checkBoxSimopsTBT,
                'checkBoxVibracaoTBT' : checkBoxVibracaoTBT,
                'checkBoxMovObjTBT' : checkBoxMovObjTBT,
                'checkBoxSubsPeriTBT' : checkBoxSubsPeriTBT,
            }         

                with open(os.path.join(resource_filename(__name__, 'inputs.json')), "w") as f:
                    json.dump(self.inputs, f)
                    
            
        except Exception as e:
                x = self.emsg.setInformativeText(str(e))
                x = self.emsg.exec_()
                print(f'o erro ao tentar salvar: {e}')
                pass
    
    #Seria o "Abrir", só que como seria mais complicado, abrir de fato um arquivo, preferi deixar o usuário salvar um template.         
    def escolherInputGroup(self):
        try:             
            if os.path.exists(self.file_name):
                with open(self.file_name, 'r') as f:
                    self.inputs = json.load(f)  
                    
                
            name, _ = QInputDialog.getItem(self, 'Carregar', 'Favor selecionar seu template:', list(self.inputs.keys()), 0, False)

            if name:
                inputs = self.inputs[name]
                
                #específico TBT
                self.inputRiscos.setText(inputs['inputRiscos'])
                self.inputGatilhos.setText(inputs['inputGatilhos'])
                self.checkBoxColuna.setChecked(inputs['checkBoxColuna'])
                self.checkBoxErgonomia.setChecked(inputs['checkBoxErgonomia'])
                
                #secondary_window
                self.checkBoxBioTBT = inputs['checkBoxBioTBT']
                self.checkBoxCalorTBT = inputs['checkBoxCalorTBT']
                self.checkBoxClimaTBT = inputs['checkBoxClimaTBT']
                self.checkBoxDerramamentoTBT = inputs['checkBoxDerramamentoTBT']
                self.checkBoxEleTBT = inputs['checkBoxEleTBT']
                self.checkBoxEntrSaiTBT = inputs['checkBoxEntrSaiTBT']
                self.checkBoxEquipPeriTBT = inputs['checkBoxEquipPeriTBT']
                self.checkBoxEscorregoesTBT = inputs['checkBoxEscorregoesTBT']
                self.checkBoxEquipDeIcaTBT = inputs['checkBoxEquipDeIcaTBT']
                self.checkBoxIluTBT = inputs['checkBoxIluTBT']
                self.checkBoxLocElevTBT = inputs['checkBoxLocElevTBT']
                self.checkBoxManuseioTBT = inputs['checkBoxManuseioTBT']
                self.checkBoxPessoaFamiTBT = inputs['checkBoxPessoaFamiTBT']
                self.checkBoxOutraTBT = inputs['checkBoxOutraTBT']
                self.checkBoxFrioTBT = inputs['checkBoxFrioTBT']
                self.checkBoxRadiacaoTBT = inputs['checkBoxRadiacaoTBT']
                self.checkBoxRuidosTBT = inputs['checkBoxRuidosTBT']
                self.checkBoxQuedaTBT = inputs['checkBoxQuedaTBT']
                self.checkBoxPressaoTBT = inputs['checkBoxPressaoTBT']
                self.checkBoxSimopsTBT = inputs['checkBoxSimopsTBT']
                self.checkBoxVibracaoTBT = inputs['checkBoxVibracaoTBT']
                self.checkBoxMovObjTBT = inputs['checkBoxMovObjTBT']
                self.checkBoxSubsPeriTBT = inputs['checkBoxSubsPeriTBT'] 
                
                #PT
                self.comboBoxJobType.setCurrentText(inputs['jobType'])
                self.inputDetentor.setText(inputs['inputDetentor'])
                self.inputEpiEspecifico.setText(inputs['inputEpiEspecifico'])
                self.inputLocSenso.setText(inputs['inputLocSenso'])
                self.inputUHF.setText(inputs['inputUHF'])
                self.inputFerramentas.setText(inputs['inputFerramentas'])
                self.inputJRA.setText(inputs['inputJRA'])
                self.textEditJobDescription.setText(inputs['textEditJobDescription'])
                self.inputArea.setText(inputs['inputArea'])
                self.inputDeck.setText(inputs['inputDeck'])
                self.inputAnexos.setText(inputs['inputAnexos'])
                self.inputHoraInspecao.setText(inputs['inputHoraInspecao'])
                self.inputResponsavelInspecao.setText(inputs['inputResponsavelInspecao'])
                self.inputIsolamentoEletrico.setText(inputs['inputIsolamentoEletrico'])
                self.inputPrecaucoes.setText(inputs['inputPrecaucoes'])
                self.inputLuvas.setText(inputs['inputLuvas'])
                self.inputResgatista1.setText(inputs['inputResgatista1'])
                self.inputResgatista2.setText(inputs['inputResgatista2'])
                self.inputResgatista3.setText(inputs['inputResgatista3'])
                self.inputResgatista4.setText(inputs['inputResgatista4'])
                self.inputObservadorVigia.setText(inputs['inputObservadorVigia'])
                self.inputAutoridade.setText(inputs['inputAutoridade']) 
                self.inputTecnico.setText(inputs['inputTecnico']) 
                self.dateEditHoje.setText(inputs['dateEditHoje']) 
                self.timeEditInicio.setText(inputs['timeEditInicio']) 
                self.timeEditFinal.setText(inputs['timeEditFinal']) 
                self.inputExecutante1.setText(inputs['inputExecutante1']) 
                self.inputExecutante2.setText(inputs['inputExecutante2']) 
                self.inputExecutante3.setText(inputs['inputExecutante3']) 
                self.inputExecutante1.setText(inputs['inputExecutante1']) 
                self.inputExecutante4.setText(inputs['inputExecutante4']) 
                self.inputExecutante5.setText(inputs['inputExecutante5']) 
                self.inputDepartamento.setText(inputs['inputDepartamento']) 
                self.checkBoxDetectorGas.setChecked(inputs['checkBoxDetectorGas'])
                self.checkBoxIsoMec.setChecked(inputs['checkBoxIsoMec'])
                self.checkBoxIsoEle.setChecked(inputs['checkBoxIsoEle'])
                self.checkBoxExtintor.setChecked(inputs['checkBoxExtintor'])
                self.checkBoxMaquinaSolda.setChecked(inputs['checkBoxMaquinaSolda'])
                self.checkBoxComunicacao.setChecked(inputs['checkBoxComunicacao'])
                self.checkBoxComunicacao.setChecked(inputs['checkBoxComunicacao'])
                self.checkBoxDrenos.setChecked(inputs['checkBoxDrenos'])
                self.checkBoxBarreira.setChecked(inputs['checkBoxBarreira'])
                self.checkBoxCooperar.setChecked(inputs['checkBoxCooperar'])
                self.checkBoxTrabalhoEmAltura.setChecked(inputs['checkBoxTrabalhoEmAltura'])
                self.checkBoxFispq.setChecked(inputs['checkBoxFispq'])
                self.checkBoxChecklist.setChecked(inputs['checkBoxChecklist'])
                self.checkBoxIcamento.setChecked(inputs['checkBoxIcamento'])
                self.checkBoxConfinado.setChecked(inputs['checkBoxConfinado'])
                self.checkBoxEpiEspecial.setChecked(inputs['checkBoxEpiEspecial'])
                self.checkBoxOutros.setChecked(inputs['checkBoxOutros'])
            
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            pass 
        
        
if __name__ == '__main__':
    qt = QApplication(sys.argv)
    novo = Novo()
    novo.show()    
    
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(os.path.join(resource_filename(__name__, 'EasyPT.py')))
    
    with open(os.path.join(resource_filename(__name__, 'style.qss')), "r") as f:
        _style = f.read()
        qt.setStyleSheet(_style)
        
    os.system("ie4uinit.exe -show") #reseta os icones do sistema pro nosso aparecer
        
    qt.exec_()
