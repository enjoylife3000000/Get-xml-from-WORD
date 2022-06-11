import myTreeView1
import combobox1
from PySide2.QtWidgets import QApplication, QHeaderView, QAbstractItemView, QFrame, QMainWindow, QWidget
from PySide2.QtGui import QIcon, QFont, QColor
from PySide2.QtCore import Qt
import sys

class UseForm(QMainWindow):

    def __init__(self, getdata, searchdata, clickontree, clickontable, expanddata, filterselect):
        self.ui = myTreeView1.Ui_MainWindow()
        self.ui.setupUi(self.ui)
        self.ui.setWindowTitle('读取docx的document xml工具')
        self.ui.setWindowIcon(QIcon('../image/1.ico'))
        self.ui.pushButton.setText('读取数据')
        self.ui.pushButton_2.setText('搜索数据')
        self.ui.pushButton_3.setText('展开数据')
        self.ui.pushButton_4.setText('搜索设置')
        self.tree = self.ui.treeWidget
        self.table = self.ui.tableWidget
        self._myformat = formFormat()
        self.bt3 = self.ui.pushButton_3
        self.text = self.ui.textEdit
        self._myformat.main_form_format(self.ui)
        self._myformat.tree_widegt_format(self.tree)
        self._myformat.table_widget_format(self.table)
        self._myformat.text_edit_format(self.text)
        self.ui.pushButton.clicked.connect(getdata)
        self.ui.pushButton_2.clicked.connect(searchdata)
        self.ui.pushButton_3.clicked.connect(expanddata)
        self.ui.treeWidget.clicked.connect(clickontree)
        self.ui.tableWidget.clicked.connect(clickontable)
        self.ui.pushButton_4.clicked.connect(filterselect)
        self.ui.label.setText('')
        self.ui.label.setEnabled(False)

class ComboboxForm(QWidget):
    def __init__(self, clickok):
        self.ui = combobox1.Ui_Form()
        self.ui.setProperty('name', 'widget')
        self.ui.setupUi(self.ui)
        self.ui.setWindowFlags(Qt.WindowCloseButtonHint)
        self.ui.setWindowTitle(' ')
        self.ui.setWindowIcon(QIcon('../image/1.ico'))
        self.ui.pushButton.setText('确认')
        self._myformat = formFormat()
        self._myformat.main_form_format(self.ui)
        self.ui.setStyleSheet(
            "QPushButton{font-size :7.5pt;background-color:rgb(230,240,210)}"
            "QCombobox{font-color: rgb(140, 140, 140)}"
            "QWidget[name = 'widget'] {color: rgb(0,40,80); background-color: rgb(230,240,210);font-size :6.5pt;}")
        self.ui.comboBox.setEditable(False)
        self.ui.comboBox.addItems(['', '内容', '标签', '属性'])
        self.ui.pushButton.clicked.connect(clickok)

class formFormat(object):
    def table_widget_format(self, tablewidget):
        hfont = QFont()
        hfont.setFamily("Arial")
        hfont.setPointSize(6)
        hfont.setItalic(True)
        vfont = QFont()
        vfont.setFamily("Arial")
        vfont.setPointSize(5)
        vfont.setItalic(True)
        tablewidget.horizontalHeader().setFont(hfont)
        tablewidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tablewidget.verticalHeader().setFont(vfont)
        tablewidget.horizontalHeader().setStyleSheet(
            "QHeaderView::section{background-color:rgb(193,210,240)};]")
        tablewidget.verticalHeader().setStyleSheet(
            "QHeaderView::section{background-color:rgb(193,210,240)};]")
        tablewidget.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        tablewidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tablewidget.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical{width:10px;background:transparent;background-color:rgb(248, 248, 248);margin:0px,0px,0px,0px;padding-top:10px;padding-bottom:10px;}"
            "QScrollBar::handle:vertical{width:10px;background:lightgray ;border-radius:5px;min-height:20px;}"
            "QScrollBar::handle:vertical:hover{width:10px;background:gray;border-radius:5px;min-height:20px;}"
            "QScrollBar::add-line:vertical{height:10px;width:10px;border-image:url(:/button/images/button/down.png);subcontrol-position:bottom;}"
            "QScrollBar::sub-line:vertical{height:10px;width:10px;border-image:url(:/button/images/button/up.png);subcontrol-position:top;}"
            "QScrollBar::add-line:vertical:hover{height:10px;width:10px;border-image:url(:/button/images/button/down_mouseDown.png);subcontrol-position:bottom;}"
            "QScrollBar::sub-line:vertical:hover{height:10px;width:10px;border-image:url(:/button/images/button/up_mouseDown.png);subcontrol-position:top;}"
            "QScrollBar::add-page:vertical,QScrollBar::sub-page:vertical{background:transparent;border-radius:5px;}")
        tablewidget.horizontalScrollBar().setStyleSheet(
            "QScrollBar:horizontal{height:10px;background:transparent;background-color:rgb(248, 248, 248);margin:0px,0px,0px,0px;padding-left:10px;padding-right:10px;}"
            "QScrollBar::handle:horizontal{height:10px;background:lightgray;border-radius:5px;/*min-height:20;*/}"
            "QScrollBar::handle:horizontal:hover{height:10px;background:gray;border-radius:5px;/*min-height:20;*/}"
            "QScrollBar::add-line:horizontal{/*height:10px;width:10px;*/border-image:url(:/button/images/button/right.png);/*subcontrol-position:right;*/}"
            "QScrollBar::sub-line:horizontal{/*height:10px;width:10px;*/border-image:url(:/button/images/button/left.png);/*subcontrol-position:left;*/}"
            "QScrollBar::add-line:horizontal:hover{/*height:10px;width:10px;*/border-image:url(:/button/images/button/right_mouseDown.png);/*subcontrol-position:right;*/}"
            "QScrollBar::sub-line:horizontal:hover{/*height:10px;width:10px;*/border-image:url(:/button/images/button/left_mouseDown.png);/*subcontrol-position:left;*/}"
            "QScrollBar::add-page:horizontal,QScrollBar::sub-page:horizontal{background:transparent;border-radius:5px;}")
    def tree_widegt_format(self,treewidegt):
        treewidegt.header().setStyleSheet("QHeaderView::section{background-color:rgb(193,210,240)};]")
        treewidegt.header().setSectionResizeMode(QHeaderView.Stretch)
        treewidegt.header().setDefaultAlignment(Qt.AlignCenter)
        treewidegt.setEditTriggers(QAbstractItemView.NoEditTriggers)
        treewidegt.verticalScrollBar().setStyleSheet(
            "QScrollBar:vertical{width:10px;background:transparent;background-color:rgb(248, 248, 248);margin:0px,0px,0px,0px;padding-top:10px;padding-bottom:10px;}"
            "QScrollBar::handle:vertical{width:10px;background:lightgray ;border-radius:5px;min-height:20px;}"
            "QScrollBar::handle:vertical:hover{width:10px;background:gray;border-radius:5px;min-height:20px;}"
            "QScrollBar::add-line:vertical{height:10px;width:10px;border-image:url(:/button/images/button/down.png);subcontrol-position:bottom;}"
            "QScrollBar::sub-line:vertical{height:10px;width:10px;border-image:url(:/button/images/button/up.png);subcontrol-position:top;}"
            "QScrollBar::add-line:vertical:hover{height:10px;width:10px;border-image:url(:/button/images/button/down_mouseDown.png);subcontrol-position:bottom;}"
            "QScrollBar::sub-line:vertical:hover{height:10px;width:10px;border-image:url(:/button/images/button/up_mouseDown.png);subcontrol-position:top;}"
            "QScrollBar::add-page:vertical,QScrollBar::sub-page:vertical{background:transparent;border-radius:5px;}")
        treewidegt.horizontalScrollBar().setStyleSheet(
            "QScrollBar:horizontal{height:10px;background:transparent;background-color:rgb(248, 248, 248);margin:0px,0px,0px,0px;padding-left:10px;padding-right:10px;}"
            "QScrollBar::handle:horizontal{height:10px;background:lightgray;border-radius:5px;/*min-height:20;*/}"
            "QScrollBar::handle:horizontal:hover{height:10px;background:gray;border-radius:5px;/*min-height:20;*/}"
            "QScrollBar::add-line:horizontal{/*height:10px;width:10px;*/border-image:url(:/button/images/button/right.png);/*subcontrol-position:right;*/}"
            "QScrollBar::sub-line:horizontal{/*height:10px;width:10px;*/border-image:url(:/button/images/button/left.png);/*subcontrol-position:left;*/}"
            "QScrollBar::add-line:horizontal:hover{/*height:10px;width:10px;*/border-image:url(:/button/images/button/right_mouseDown.png);/*subcontrol-position:right;*/}"
            "QScrollBar::sub-line:horizontal:hover{/*height:10px;width:10px;*/border-image:url(:/button/images/button/left_mouseDown.png);/*subcontrol-position:left;*/}"
            "QScrollBar::add-page:horizontal,QScrollBar::sub-page:horizontal{background:transparent;border-radius:5px;}")
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(6.5)
        font.setItalic(True)
        treewidegt.header().setFont(font)
        treewidegt.setFrameShape(QFrame.StyledPanel)
        treewidegt.setStyleSheet("QTreeWidget {color: rgb(0,40,80); background-color: rgb(249,255,229);font-size :6.5pt;}")
        # treewidegt.setSelectionMode(QAbstractItemView.MultiSelection)
    def text_edit_format(self,textedit):
        textedit.setTextColor(QColor(35, 35, 35))
        textedit.setTextBackgroundColor(QColor(248, 248, 248))
        font = QFont()
        font.setPointSize(9)
        textedit.setFont(font)
        # textedit.setStyleSheet(
        #     "QTextEdit {color: rgb(247,55,20); background-color: rgb(248,248,248);font-size :6.5pt;}")
    def main_form_format(self,mainform):
        mainform.setFont(QFont('Arial', 7))




