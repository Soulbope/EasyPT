import datetime
import sys
import webbrowser
import json
from xml.etree.ElementTree import tostring
import PyQt5
from PyQt5 import QtCore, QtGui, uic, QtWidgets
from PyQt5.QtWidgets import QWidget
from designEasyPT import *
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QInputDialog, QMessageBox, QWidget, QLineEdit
from docxtpl import DocxTemplate
import ctypes
from pkg_resources import resource_filename
import os
import win32print
import win32api

if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)


class Novo(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.oldPos = self.pos() #poder mover a tela clicando e arrastando em qualquer lugar
        self.combobox() #preencher combobox
        self.lerpastaDestino()
        self.setWindowTitle('EasyPT')
        self.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
        self.btnEscape.clicked.connect(self.fecha)    
        self.btnSalvar.clicked.connect(self.salvar)
        self.btnConfig.clicked.connect(self.minimiza)
        self.btnEscolherPasta.clicked.connect(self.escolher_pasta)
        self.btnChoosePrinter.clicked.connect(self.escolherImpressora)
        self.btnImprimir.clicked.connect(self.imprimir)
        self.btnSaveTemplate.clicked.connect(self.salvarInputGroup)
        self.btnAbrir.clicked.connect(self.escolherInputGroup)
        self.btnLogo.clicked.connect(self.logo)
        self.inputs = {}
        self.file_name = 'inputs.json'        
        
        #popup de IMPRESSÂO
        self.msg = QMessageBox()
        self.msg.setWindowTitle("Imprimindo")
        self.msg.setText("PRONTO! Imprimindo PT!")
        self.msg.setIcon(QMessageBox.Information)
        self.msg.setStandardButtons(QMessageBox.Ok)
        self.msg.setWindowIcon(QtGui.QIcon(os.path.join(resource_filename(__name__, 'ffVermelhoSemFundo2.ico'))))
        
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
   
    def logo(self):
        webbrowser.open('https://github.com/Soulbope')  


    def fecha(self):
        self.nomeArquivoDestino = f'/PT - { self.inputDetentor.text() } - { self.dateEditHoje.text().replace("/", "_") }.docx'
        path = f'{self.inputSelectFolder.text()}{self.nomeArquivoDestino}'
        
        isFile = os.path.isfile(path)
        
        if not isFile:
            escolha = ('Salvar', 'Apenas Fechar')
            item, ok = QInputDialog.getItem(self, "Salvar", "Deseja salvar anter de fechar?", escolha, 0, False)
            
            if ok and (item=='Salvar'):
                self.salvar()
            elif ok and (item=='Apenas Fechar'):
                self.close()
        else:
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
            
    def escolher_pasta(self):
        pasta_escolhida = QFileDialog.getExistingDirectory(
            self.centralwidget,
            'Escolher pasta',
            r'C:',
            #options=QFileDialog.DontUseNativeDialog
        )
        try:
            self.inputSelectFolder.setText(pasta_escolhida.replace('\\', '/'))
        
            with open(os.path.join(resource_filename(__name__, 'pastaDestino.txt')), "w") as f:
                f.write(pasta_escolhida.replace('\\', '/'))
                
            return pasta_escolhida
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            pass
        
    def lerpastaDestino(self):
        with open(os.path.join(resource_filename(__name__, 'pastaDestino.txt')), "r") as f:
            endereco = f.read()
            self.inputSelectFolder.setText(endereco)
        
    def salvar(self):
        try:
            jobType = self.comboBoxJobType.currentText()#
            inputDetentor = self.inputDetentor.text()#
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
            inputObservador = self.inputObservador.text()#
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
                
            dataHoje = self.dateEditHoje.text().replace('/', '_')
                    
            doc = DocxTemplate(os.path.join(resource_filename(__name__, 'pttemplat1.docx')))     
            context = {
                'inputDetentor' : inputDetentor,
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
                'inputObservador' : inputObservador,
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
            nomeArquivoDestino = f'/PT - { inputDetentor } - { dataHoje }.docx'
            
            
            
            with open(os.path.join(resource_filename(__name__, 'pastaDestino.txt')), "r") as f:                
                endereco = f.read()
            
            pasta = endereco
            path = f'{pasta}{nomeArquivoDestino}'
            
            isFile = os.path.isfile(path)
            
            if isFile:
                escolha = ('Sobrescrever', 'Cancelar')
                
                item, ok = QInputDialog.getItem(self, "Existente", "Arquivo já existente, deseja sobrescrever?", escolha, 0, False)
                
                if ok and (item == 'Sobrescrever'):
                    doc.save(f'{pasta}{nomeArquivoDestino}')
                elif ok and (item == 'Cancelar'):
                    pass
            else:
                doc.save(f'{pasta}{nomeArquivoDestino}')
                
             
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
            jobType = self.comboBoxJobType.currentText()
            inputDetentor = self.inputDetentor.text()#
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
            inputObservador = self.inputObservador.text()#
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
            inputSelectFolder = self.inputSelectFolder.text()
                     
            name, _ = QtWidgets.QInputDialog.getText(self, 'Save Inputs', 'Enter a name for the inputs:')

            if name:
                self.inputs[name] = {
                'jobType' : jobType,
                'inputDetentor' : inputDetentor,
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
                'inputObservador' : inputObservador,
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
                'inputSelectFolder' : inputSelectFolder,
            }         

                with open(os.path.join(resource_filename(__name__, 'inputs.json')), "w") as f:
                    json.dump(self.inputs, f)
                    
            
        except Exception as e:
                x = self.emsg.setInformativeText(str(e))
                x = self.emsg.exec_()
                print(f'o erro é: {e}')
                pass
            
            
            
            
    
    ####################################################################################################################################################
    
    
    
    #Seria o "Abrir", só que como seria mais complicado, abrir de fato um arquivo, preferi deixar o usuário salvar um template.         
    def escolherInputGroup(self):
        try:             
            if os.path.exists(self.file_name):
                with open(self.file_name, 'r') as f:
                    self.inputs = json.load(f)  
                
            name, _ = QInputDialog.getItem(self, 'Load Inputs', 'Select a group of inputs:', list(self.inputs.keys()), 0, False)

            if name:
                inputs = self.inputs[name]
                
                self.comboBoxJobType.setCurrentText(inputs['jobType'])
                self.inputDetentor.setText(inputs['inputDetentor'])
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
                self.inputObservador.setText(inputs['inputObservador'])
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
                self.inputSelectFolder.setText(inputs['inputSelectFolder'])
            
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            pass 
        
        
        
    def imprimir(self):
        try:
            self.salvar()            
            dataHoje = self.dateEditHoje.text().replace('/', '_')
            self.nomeArquivoDestino = f'PT - { self.inputDetentor.text() } - { dataHoje }.docx'           
            win32api.ShellExecute(0, "print", self.nomeArquivoDestino , None, self.inputSelectFolder.text() , 0)
            x = self.msg.exec_()
        except Exception as e:
            x = self.emsg.setInformativeText(str(e))
            x = self.emsg.exec_()
            print(f'o erro é: {e}')
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
