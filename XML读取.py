#！/user/bing/evn python
#! conding:utf-8

import re
from PySide2.QtWidgets import QFileDialog, QTreeWidgetItem, QTableWidgetItem, QHeaderView, QApplication
from PySide2.QtCore import Qt
from PySide2.QtGui import QFont, QColor
import random
import openForm
import zipfile
import sys
import shutil
import os
try:
    import xml.etree.cElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET

class WordXMLParse(object):
    def __init__(self):
        self._path = ''
        self.xml_tag = []
        self.xml_tag_attribute = []
        self.xml_tag_value = []
        self.xml_child = []
        self.xml_parent = []
        self._root = []
        self.id = []
        self._for_tag_count = []
        self.count = 0
        self.exist_document_xml = False
        self.expand = False
        self.my_form = openForm.UseForm(self._get_xml, self.searchdata, self.clickontree, self.clickontable, self.collapse_expand, self.filterselect)
        self.my_form.tree.setColumnCount(4)
        self.my_form.tree.setHeaderLabels(['标签', '属性', '内容', 'p', 'me'])
        self.my_form.tree.setColumnHidden(3, True)
        self.my_form.tree.setColumnHidden(4, True)
        # self.my_form.tree.header().setSectionResizeMode(0, QHeaderView.Interactive)
        # self.my_form.tree.header().setSectionResizeMode(1, QHeaderView.Interactive)
        # self.my_form.tree.header().setSectionResizeMode(2, QHeaderView.Interactive)
        self.my_form.table.setColumnCount(2)
        self.my_form.table.setHorizontalHeaderLabels(['标签名', '数量'])
        self.my_form.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.cb = openForm.ComboboxForm(self.clickok)
        self.cb_value = ''
    def _recursion(self, arr, p, c):
        for i in range(len(arr)):
            if i > len(arr):
                return 1
            c = self._update_child(c)
            # print(
            #     re.compile(r'{.*}').sub('', arr[i].tag) + ' ' +
            #     re.compile(r'\'{.*?}').sub('', str(arr[i].attrib)) + ' ' +
            #     str(p) + ' ' +
            #     str(c) + ' ' +
            #     str(self.count)
            # )
            # print(re.compile(r'{.*}').sub('',arr[i].tag))
            # print(re.compile(r'\'{.*?}').sub('', str(arr[i].attrib)))
            # print(str(p))
            # print(str(c))
            self.xml_tag.append(re.compile(r'{.*}').sub('', arr[i].tag))
            if self._for_tag_count.count(re.compile(r'{.*}').sub('', arr[i].tag)) == 0:
                self._for_tag_count.append(re.compile(r'{.*}').sub('', arr[i].tag))
            self.xml_tag_attribute.append(re.compile(r'\'{.*?}').sub('', str(arr[i].attrib)))
            if arr[i].text is None:
                self.xml_tag_value.append('')
            else:
                self.xml_tag_value.append(str(arr[i].text))
            self.xml_parent.append(str(p))
            self.xml_child.append(str(c))
            self.count += 1
            self._recursion(arr[i], c, self._update_child(c))
    def _update_child(self, c):
        child = random.randrange(10002, 99999)
        if self.id.count(child) == 0:
            self.id.append(child)
            return child
        return self._update_child(child)
    def takeSecond(self, elem):
        return elem[1]
    def _get_xml(self):
        # self._path = QFileDialog.getOpenFileName(None, 'WORD XML展示工具', './', 'XML文件(*.xml *.XML)')
        # self._path = QFileDialog.getOpenFileName(None, 'WORD XML展示工具', './', 'XML文件(*.zip *.ZIP)')
        self._path = QFileDialog.getOpenFileName(None, 'WORD XML展示工具', './', 'XML文件(*.docx *.DOCX)')
        if self._path[0] == '':
            pass
        else:
            self.my_form.tree.clear()
            store_path = self._path[0]
            store_path = store_path[0:store_path.rfind('/')]
            store_path = store_path.replace('/', '\\')
            # shutil.copy(self._path[0], sys.path[0] + r'\temp.docx')  # 复制docx文件
            shutil.copy(self._path[0], store_path + r'\temp.docx')
            # os.rename(sys.path[0] + r'\temp.docx', sys.path[0] + r'\temp.zip')  # 重命名docx至zip
            os.rename(store_path + r'\temp.docx', store_path + r'\temp.zip')
            # zp = zipfile.ZipFile(self._path[0], mode='r')  # 获取zip压缩包中所有的文件
            # 此处需要判断zp压缩包中是否为空，如果压缩包为空,退出函数，return返回0
            try:
                # zp = zipfile.ZipFile(sys.path[0] + r'\temp.zip', mode='r')
                zp = zipfile.ZipFile(store_path + r'\temp.zip', mode='r')
            except:
                # os.remove(sys.path[0] + r'\temp.zip')
                os.remove(store_path + r'\temp.zip')
                return 0
            for zip_file in zp.namelist():
                if zip_file == 'word/document.xml':
                    # xml_path = sys.path[0] + '/word/document.xml'
                    xml_path = store_path + '/word/document.xml'
                    xml_path = xml_path.replace('/', '\\')
                    # zp.extract(r'word/document.xml', path=sys.path[0], pwd=None)  # 提取document.xml
                    zp.extract(r'word/document.xml', path=store_path, pwd=None)
                    # del_folder = sys.path[0] + '/word'
                    del_folder = store_path + '/word'
                    del_folder = del_folder.replace('/', '\\')
                    self.exist_document_xml = True
                    break
            if self.exist_document_xml == False:
                # del_folder = sys.path[0] + '/word'
                del_folder = store_path + '/word'
                del_folder = del_folder.replace('/', '\\')
                zp.close()
                shutil.rmtree(del_folder)
                # os.remove(sys.path[0] + r'\temp.zip')
                os.remove(store_path + r'\temp.zip')
                return 0
            # self._root = ET.parse(self._path[0]).getroot()
            self._root = ET.parse(xml_path).getroot()
            self._recursion(self._root, 10000, 10001)
            i = 0
            for xml_p in self.xml_parent:
                search = self.my_form.tree.findItems(xml_p, Qt.MatchRecursive, 4)
                if search == []:
                    node = QTreeWidgetItem(self.my_form.tree)
                else:
                    node = QTreeWidgetItem(search[0])
                node.setText(0, self.xml_tag[i])
                node.setText(1, self.xml_tag_attribute[i])
                node.setText(2, self.xml_tag_value[i])
                node.setText(3, self.xml_parent[i])
                node.setText(4, self.xml_child[i])
                i += 1
            print(str(self.count))
            print(self.xml_child.count('None'))
            print(self.xml_parent.count('None'))
            i = 0
            itmp = []
            for tmp in self._for_tag_count:
                itmp.append([tmp, self.xml_tag.count(tmp)])
            itmp.sort(key=lambda elem: (-elem[1], elem[0]), reverse= False)  # -与reverse可以控制二次排序....
            self.my_form.table.setRowCount(len(itmp))
            i = 0
            for tmp in itmp:
                self.my_form.table.setItem(i, 0, QTableWidgetItem(tmp[0]))
                self.my_form.table.setItem(i, 1, QTableWidgetItem(str(tmp[1])))
                self.my_form.table.item(i, 1).setTextAlignment(Qt.AlignCenter)
                i += 1
            self.xml_child.clear()
            self.xml_parent.clear()
            self.xml_tag_attribute.clear()
            self.xml_tag.clear()
            self._for_tag_count.clear()
            self.my_form.tree.expandAll()
            # self.my_form.tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
            self.my_form.tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            self.my_form.tree.header().setSectionResizeMode(1, QHeaderView.Interactive)
            self.my_form.tree.header().setSectionResizeMode(2, QHeaderView.Interactive)
            zp.close()
            shutil.rmtree(del_folder)
            # os.remove(sys.path[0] + r'\temp.zip')
            os.remove(store_path + r'\temp.zip')
            self.expand = True
            self.my_form.bt3.setText('折叠数据')
    def searchdata(self):
        # 防止粘贴带入样式，需要使用监听事件。。。。 以后实现
        # self.my_form.text.setTextColor(QColor(35, 35, 35))
        # self.my_form.text.setTextBackgroundColor(QColor(248, 248, 248))
        # font = QFont()
        # font.setPointSize(9)
        # # self.my_form.text.setFont(font)
        # self.my_form.text.setCurrentFont(font)
        # self.my_form.text.setText(self.my_form.text.toPlainText())

        # 遍历所有节点
        if self.cb_value =='标签':
            column = 0
        elif self.cb_value =='属性':
            column = 1
        else:
            column = 2
        root_count = self.my_form.tree.topLevelItemCount()
        i = 0
        while i < root_count:
            root = self.my_form.tree.topLevelItem(i)
            self.test_digui(root, root.childCount(), column)
            i += 1
    def test_digui(self, item, index, column):
        if item.childCount() == 0:
            return 1
        for i in range(item.childCount()):
            if self.my_form.text.toPlainText() != '':
                if item.child(i).text(column).upper().find(self.my_form.text.toPlainText().upper()) >= 0:
                    item.child(i).setTextColor(column, QColor(254, 254, 254))
                    item.child(i).setBackgroundColor(column, QColor(210, 80, 10))
                    # print(item.child(i).text(2) + ' ' + str(item) + ' ' + str(i))
                    for ii in range(3):
                        if ii != column:
                            item.child(i).setTextColor(ii, QColor(0, 40, 80))
                            item.child(i).setBackgroundColor(ii, QColor(249, 255, 229))
                else:
                    for ii in range(3):
                        if ii != column:
                            item.child(i).setTextColor(ii, QColor(0, 40, 80))
                            item.child(i).setBackgroundColor(ii, QColor(249, 255, 229))
                    item.child(i).setTextColor(column, QColor(0, 40, 80))
                    item.child(i).setBackgroundColor(column, QColor(249, 255, 229))
            else:
                for ii in range(3):
                    item.child(i).setTextColor(ii, QColor(0, 40, 80))
                    item.child(i).setBackgroundColor(ii, QColor(249, 255, 229))
            self.test_digui(item.child(i), item.childCount(), column)
    def clickontree(self):
        # item = self.my_form.tree.currentItem()
        # print(item.parent().indexOfChild(item))
        # aa = (item.text(0))
        # b = (item.text(1))
        # c = (item.text(2))
        # item.setText(0, aa)
        # item.setText(1, b)
        # item.setText(2, c)
        pass
    def clickontable(self):
        pass
        #print('click on table')
    def collapse_expand(self):
        if self.expand == False:
            self.my_form.tree.expandAll()
            self.my_form.bt3.setText('折叠数据')
            self.expand = True
        elif self.expand == True:
            self.my_form.tree.collapseAll()
            self.my_form.bt3.setText('展开数据')
            self.expand = False
    def filterselect(self):
        self.cb.ui.setWindowModality(Qt.ApplicationModal)
        self.cb.ui.show()
    def clickok(self):
        self.cb_value = self.cb.ui.comboBox.currentText()
        self.cb.ui.hide()
app = QApplication(sys.argv)
main = WordXMLParse()
main.my_form.ui.show()
sys.exit(app.exec_())


