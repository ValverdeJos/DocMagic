# coding:utf-8
import os
import sys
from tkinter import filedialog
import qtawesome as qta
import ExporterPub as ConverterPDF
import ExporterExecl as ConverterExcel
import ExporterWord as ConverterWord

from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QIcon, QPainter, QImage, QBrush, QColor, QFont,QPixmap
from PyQt5.QtWidgets import QApplication, QFrame, QStackedWidget, QHBoxLayout, QLabel,QPushButton,QCheckBox
from qfluentwidgets import (NavigationInterface, NavigationItemPosition, NavigationWidget, MessageBox, isDarkTheme, setTheme, Theme)
from qfluentwidgets import FluentIcon as FIF
from qframelesswindow import FramelessWindow, StandardTitleBar

extention = [".xls",".xlsm",".xlsx",".xlt",".xlsb",".xltx",".xltm"]
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

class Widget(QFrame):

    def __init__(self, text: str, parent=None):
        super().__init__(parent=parent)
        
        self.label = QLabel(text, self)
        self.label.setAlignment(Qt.AlignCenter)
        self.hBoxLayout = QHBoxLayout(self)
        self.hBoxLayout.addWidget(self.label, 1, Qt.AlignCenter)
        self.setObjectName(text.replace(' ', '-'))

class AvatarWidget(NavigationWidget):
    """ Avatar widget """

    def __init__(self, parent=None):
        super().__init__(isSelectable=False, parent=parent)
        self.avatar = QImage(f'{desktop_path}\\DocMagic-main\\resource\\shoko.png').scaled(
            24, 24, Qt.KeepAspectRatio, Qt.SmoothTransformation)

    def paintEvent(self, e):
        painter = QPainter(self)
        painter.setRenderHints(
            QPainter.SmoothPixmapTransform | QPainter.Antialiasing)

        painter.setPen(Qt.NoPen)

        if self.isPressed:
            painter.setOpacity(0.7)

        # draw background
        if self.isEnter:
            c = 255 if isDarkTheme() else 0
            painter.setBrush(QColor(c, c, c, 10))
            painter.drawRoundedRect(self.rect(), 5, 5)

        # draw avatar
        painter.setBrush(QBrush(self.avatar))
        painter.translate(8, 6)
        painter.drawEllipse(0, 0, 24, 24)
        painter.translate(-8, -6)

        if not self.isCompacted:
            painter.setPen(Qt.white if isDarkTheme() else Qt.black)
            font = QFont('Segoe UI')
            font.setPixelSize(14)
            painter.setFont(font)
            painter.drawText(QRect(44, 0, 255, 36), Qt.AlignVCenter, 'User')

class Window(FramelessWindow):

    def __init__(self):
        super().__init__()
        self.setTitleBar(StandardTitleBar(self))

        setTheme(Theme.DARK)


        self.hBoxLayout = QHBoxLayout(self)
        self.navigationInterface = NavigationInterface(self, showMenuButton=True)
        self.stackWidget = QStackedWidget(self)
        self.iconWord = qta.icon('fa.file-word-o', color='white',scale_factor=1.1)
        self.iconPDF = qta.icon('mdi.file-pdf', color='white',scale_factor=1.3)
        self.iconExcel = qta.icon('mdi.file-excel', color='white',scale_factor=1.3)


        # Ajouter un QLabel pour l'image
        self.imageLabel = QLabel(self)
        self.imageLabel.setPixmap(QPixmap(f"{desktop_path}\\DocMagic-main\\resource\\assets\\Xls.png").scaled(200, 200))
        self.imageLabel.move(350, 170)
        self.imageLabel.resize(300, 300)

        self.upload = QPushButton("Escolha o seu Ficheiro", self)
        self.upload.move(230, 580)
        self.upload.resize(160, 45)
        self.upload.setObjectName("upload")
        self.upload.clicked.connect(self.UploadFiles)

        self.confirm = QPushButton("Confirma", self)
        self.confirm.move(570, 580)
        self.confirm.resize(160, 45)
        self.confirm.setObjectName("confirm")
        self.confirm.clicked.connect(self.run_code)

        self.checkBoxSplit = QCheckBox("Separa ficheiro ",self)
        self.checkBoxSplit.move(245, 545)
        self.checkBoxSplit.resize(120, 30)
        self.checkBoxSplit.setObjectName("checkBoxSplit")
        self.checkBoxSplit.stateChanged.connect(self.CheckText)


        with open(f"{desktop_path}\\DocMagic-main\\resource\\style.css", "r") as f:
            style_sheet = f.read()
        self.setStyleSheet(style_sheet)

        # create sub interface
        self.pubInterface = Widget('File Publisher', self)
        self.excelInterface = Widget('File Excel', self)
        self.pdfInterface = Widget('File PDF', self)
        self.wordInterface = Widget('File Word', self)
        self.folderInterface = Widget('Folder Interface', self)
        self.settingInterface = Widget('Setting Interface', self)
        

        # initialize layout
        self.initLayout()

        # add items to navigation interface
        self.initNavigation()

        self.initWindow()

    def initLayout(self):
        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(0, self.titleBar.height(), 0, 0)
        self.hBoxLayout.addWidget(self.navigationInterface)
        self.hBoxLayout.addWidget(self.stackWidget)
        self.hBoxLayout.setStretchFactor(self.stackWidget, 1)

    def initNavigation(self):
        
        self.addSubInterface(self.pubInterface, FIF.DOCUMENT, 'Publisher')
        self.addSubInterface(self.excelInterface, self.iconExcel, 'Excel')
        self.addSubInterface(self.pdfInterface, self.iconPDF , 'PDF')
        self.addSubInterface(self.wordInterface, self.iconWord, 'Word')


        self.navigationInterface.addSeparator()

        # add navigation items to scroll area
        self.addSubInterface(self.folderInterface, FIF.FOLDER, 'Folder library', NavigationItemPosition.SCROLL)
        # for i in range(1, 21):
        #     self.navigationInterface.addItem(
        #         f'folder{i}',
        #         FIF.FOLDER,
        #         f'Folder {i}',
        #         lambda: print('Folder clicked'),
        #         position=NavigationItemPosition.SCROLL
        #     )

        # add custom widget to bottom
        self.navigationInterface.addWidget(
            routeKey='avatar',
            widget=AvatarWidget(),
            onClick=self.showMessageBox,
            position=NavigationItemPosition.BOTTOM
        )

        self.addSubInterface(self.settingInterface, FIF.SETTING, 'Settings', NavigationItemPosition.BOTTOM)

        #!IMPORTANT: don't forget to set the default route key if you enable the return button
        # self.navigationInterface.setDefaultRouteKey(self.musicInterface.objectName())

        # set the maximum width
        # self.navigationInterface.setExpandWidth(300)

        self.stackWidget.currentChanged.connect(self.onCurrentInterfaceChanged)
        self.stackWidget.setCurrentIndex(1)

    def initWindow(self):
        
        self.resize(900, 700)
        self.setWindowIcon(QIcon('f"{desktop_path}\\DocMagic-main\\resource\\LogoPng.png'))
        self.setWindowTitle('DocMagic')
        self.titleBar.setAttribute(Qt.WA_StyledBackground)

        desktop = QApplication.desktop().availableGeometry()
        w, h = desktop.width(), desktop.height()
        self.move(w//2 - self.width()//2, h//2 - self.height()//2)

        self.setQss()

        self.pubInterface.setHidden(True)
        self.excelInterface.setHidden(True)
        self.pdfInterface.setHidden(True)
        self.wordInterface.setHidden(True)
        self.folderInterface.setHidden(True)
        self.settingInterface.setHidden(True)

    def addSubInterface(self, interface, icon, text: str, position=NavigationItemPosition.TOP):
        """ add sub interface """
        self.stackWidget.addWidget(interface)
        self.navigationInterface.addItem(
            routeKey=interface.objectName(),
            icon=icon,
            text=text,
            onClick=lambda: self.switchTo(interface),
            position=position,
            tooltip=text
        )

    def setQss(self):

        color = 'dark' if isDarkTheme() else 'light'
        with open(f'{desktop_path}\\DocMagic-main\\resource\\{color}\\demo.qss', encoding='utf-8') as f:
            self.setStyleSheet(f.read())

    def switchTo(self, widget):
        self.stackWidget.setCurrentWidget(widget)

    def onCurrentInterfaceChanged(self, index):
        widget = self.stackWidget.widget(index)
        arrayName = widget.objectName()
        Name = arrayName.split('-')
        pathImages={
            "Excel": f"{desktop_path}\\DocMagic-main\\assets\\Xls.png",
            "Publisher": f"{desktop_path}\\DocMagic-main\\assets\\Pub.png",
            "PDF": f"{desktop_path}\\DocMagic-main\\assets\\Pdf.png",
            "Word": f"{desktop_path}\\DocMagic-main\\assets\\Word.png"
            }


        self.upload.pyqtConfigure(text=f"Escolha o seu {Name[1]}")
        #print(index)

        if Name[1] in pathImages:
            pathImage = pathImages[Name[1]]
            self.imageLabel.setPixmap(QPixmap(pathImage).scaled(200,200)) 
            self.textBreve = QLabel("Funcionalidade Brevemente disponivel", self) 
            self.textBreve.resize(300, 30)
            self.textBreve.move(350,500)
            self.textBreve.setObjectName("Feature")
        else:
            self.imageLabel.clear()

        if index == 4:
            
            self.upload.setHidden(True)
            self.confirm.setHidden(True)
            self.checkBoxSplit.setHidden(True)
            self.folderInterface.setHidden(False)
            self.textBreve.setHidden(True)
            
        elif index == 5:
            self.settingInterface.setHidden(False)
            self.upload.setHidden(True)
            self.confirm.setHidden(True)
            self.textBreve.setHidden(True)
            self.checkBoxSplit.setHidden(True)


        else:
            self.upload.setHidden(False)
            self.confirm.setHidden(False)
            self.checkBoxSplit.setHidden(False)
            self.pubInterface.setHidden(True)
            self.excelInterface.setHidden(True)
            self.pdfInterface.setHidden(True)
            self.wordInterface.setHidden(True)
            self.textBreve.setHidden(True)
              
    def showMessageBox(self):
        w = MessageBox(
            'New feature ',
            'Isto é uma feature que vais ser desenvolvida em breve 😉',
            self
        )
        w.exec()

    def UploadFiles(self):
            global file_path
            file_path = filedialog.askopenfilename()

    def CheckText(self):
        if self.checkBoxSplit.isChecked():
            self.checkBoxSplit.pyqtConfigure(text="Separa ficheiro On")
        else:
            self.checkBoxSplit.pyqtConfigure(text="Separa ficheiro Off")

    def run_code(self):
        self.MessageSplit = MessageBox(
            'Tarefa Concluida',
            'A Tarefa foi concluida com sucesso Pdfs Separados 😉',
            self
        )
        self.MessageNoSplit = MessageBox(
            'Tarefa Concluida',
            'A Tarefa foi concluida com sucesso Pdfs não separados 😉',
            self
        )
        self.MessageError = MessageBox(
            'Error',
            'Error: file_path esta vazio',
            self
        )
        self.MessageErrorExecl = MessageBox(
            'Error',
            'Não é um fichiero Excel.',
            self
        )
        
        _, file_extension = os.path.splitext(file_path)

        if file_extension == ".pub":
            if self.checkBoxSplit.isChecked():

                try:
                    ConverterPDF.get_pdf_pages_from_pub(pub_location=file_path)
                    #messagebox.showinfo("Sucesso","Converção em Pdfs Separados foi um sucesso")
                    self.MessageSplit.exec()
                except:
                    self.MessageError.exec()
            else:

                try:
                    ConverterPDF.get_pdf_SimplePage_from_pub(pub_location=file_path)
                    #messagebox.showinfo("Sucesso","Converção em Pdf foi um sucesso")
                    self.MessageNoSplit.exec()
                except:
                    self.MessageError.exec()
        elif file_extension == ".pdf":

            try:
                ConverterPDF.cut_pdf_to_pages(path_PDF=file_path)
                self.MessageNoSplit.exec()
            except:
                self.MessageError.exec()

        elif file_extension in extention :

            if self.checkBoxSplit.isChecked():

                if file_path:
                    try:
                        ConverterExcel.get_pdf_pages_from_Excel(pub_location=file_path)
                        self.MessageSplit.exec()
                    except:
                        self.MessageError.exec()
                else:
                    self.MessageErrorExecl.exec()

            else:
                if file_path:
                    try:
                        ConverterExcel.get_pdf_SimplePage_from_Excel()
                        self.MessageNoSplit.exec()
                    except:
                        self.MessageError.exec()
                else:
                    self.MessageErrorExecl.exec()
        elif file_extension == ".docx":
            if self.checkBoxSplit.isChecked():

                if file_path:
                    try:
                        ConverterWord.get_pdf_pages_from_Word(pub_location=file_path)
                        self.MessageSplit.exec()
                    except:
                        self.MessageError.exec()
                else:
                    self.MessageErrorExecl.exec()
            else:
                if file_path:
                    try:
                        ConverterWord.get_pdf_SimplePage_from_Word(pub_location=file_path)
                        self.MessageSplit.exec()
                    except:
                        self.MessageError.exec()
                else:
                    self.MessageErrorExecl.exec()
        
if __name__ == '__main__':
    
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

    app = QApplication(sys.argv)
    w = Window()
    w.show()
    app.exec_()