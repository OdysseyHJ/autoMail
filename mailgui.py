import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWebEngineWidgets import *

import traceback

from commonlib import write_log
from mail import g_mail_dict
from mail import g_mail_common

class CAutoMail(QWidget):
    def __init__(self, mail_dict):
        super().__init__()
        self.sections = {}
        self.mail_dict = mail_dict
        self.initUI()

    def initUI(self):

        # self.setGeometry(300, 300, 1000, 800)
        self.setWindowTitle('autoMail')
        # self.setWindowIcon(QIcon('AC.jpg'))
        self.resize(620, 600)

        self.table = CMailTable(self.mail_dict)


        self.btn_clear = QPushButton("重置发送状态")
        self.btn_clear.clicked.connect(self.button_event_clear_sended_set)

        self.btn_reset_send = QPushButton("发送选中邮件")
        self.btn_reset_send.clicked.connect(self.button_event_send)

        self.btn_selected_all = QPushButton("全选")
        self.btn_selected_all.clicked.connect(self.button_event_selected_all)

        self.btn_clear_selected = QPushButton("清空选择")
        self.btn_clear_selected.clicked.connect(self.button_event_clear_selected)

        self.spacerItem = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.vbox = QVBoxLayout()
        self.vbox.addWidget(self.btn_clear)
        self.vbox.addWidget(self.btn_reset_send)
        self.vbox.addWidget(self.btn_selected_all)
        self.vbox.addWidget(self.btn_clear_selected)

        self.vbox.addSpacerItem(self.spacerItem)
        self.vbox2 = QVBoxLayout()
        self.vbox2.addWidget(self.table)
        self.hbox = QHBoxLayout()
        self.hbox.addLayout(self.vbox2)
        self.hbox.addLayout(self.vbox)
        self.setLayout(self.hbox)
        self.show()

    def button_event_clear_sended_set(self):
        if g_mail_common.is_selected_joint_sended():
            reply = QMessageBox.question(self, "提示", "确认重置选中邮件的发送状态?", QMessageBox.Yes|QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return
        for mail_id in g_mail_common.selected_set:
            print(mail_id)
            g_mail_common.remove_sended_set(mail_id)
        g_mail_common.dump_sended_set()
        self.table.refresh_table()

    def button_event_send(self):
        if g_mail_common.is_selected_joint_sended():
            reply = QMessageBox.question(self, "提示", "选中邮件中存在已发送的邮件,是否继续发送?", QMessageBox.Yes|QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return
        else:
            reply = QMessageBox.question(self, "提示", "确认发送选中的邮件?", QMessageBox.Yes|QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return

        g_mail_common.send_selected_mail()
        self.table.refresh_table()
        QMessageBox.information(self, "提示", "邮件发送成功", QMessageBox.Yes)

    def button_event_selected_all(self):
        for checkBox in self.table.checkBoxList:
            # print(checkBox.mail_id)
            checkBox.setChecked(True)
        # self.table.show()

    def button_event_clear_selected(self):
        for checkBox in self.table.checkBoxList:
            # print(checkBox.mail_id)
            checkBox.setChecked(False)

    def init_mail_info(self):
        raw_count = 9
        column_count = 9
        self.table.setRowCount(raw_count)
        self.table.setColumnCount(column_count)

        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        for index in range(column_count):
            self.table.horizontalHeader().resizeSection(index, 60)
            self.table.verticalHeader().resizeSection(index, 60)

        # self.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.setFont(QFont('Times', 20, QFont.Black))

        for x_pos in range(column_count):
            for y_pos in range(raw_count):
                unit_item = QTableWidgetItem()
                unit_item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                # unit_item.setFlags(Qt.ItemIsSelectable | Qt.ItemI

class CMailTable(QTableWidget):
    def __init__(self, mail_dict={}):
        super().__init__()

        self.mail_dict = mail_dict

        # 设置行数列数
        rowCnt = len(mail_dict)
        self.setRowCount(rowCnt)
        self.setColumnCount(3)

        # 设置行列标题
        rowNumList = [str(x) for x in range(1, rowCnt+1)]
        self.setVerticalHeaderLabels(rowNumList)
        columnNamelist = ['选择', '邮件ID', '发送状态']
        self.setHorizontalHeaderLabels(columnNamelist)
        self.setColumnWidth(0, 64)
        self.setColumnWidth(0, 128)
        self.setColumnWidth(0, 128)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.setAlternatingRowColors(True)

        self.setTable()

        self.show()

    def setTable(self):
        index = 0
        self.checkBoxList = []
        sorted_mail_id = list(self.mail_dict.keys())
        sorted_mail_id.sort()
        for mail_id in sorted_mail_id:
            mail_obj = self.mail_dict[mail_id]
            checkBox = MailIDCheckBox(mail_obj.id)
            # checkBox.resize(200, 200)
            # toggle 默认勾选
            # cb.toggle()
            checkBox.stateChanged.connect(checkBox.mail_select)
            checkBox.setStyleSheet("QCheckBox::indicator { width: 22px; height: 22px;}")
            hLayout = QHBoxLayout()
            hLayout.addWidget(checkBox)
            hLayout.setAlignment(checkBox, Qt.AlignRight)
            widget = QWidget()
            widget.setLayout(hLayout)
            self.checkBoxList.append(checkBox)

            button = MailIDButton(mail_obj)
            # button.setDown(True)
            # 修改按钮大小
            button.setStyleSheet("QPushButton{margin:4px;font-size:24px};")
            button.clicked.connect(button.showInfo)

            send_status = "未发送"
            if g_mail_common.is_in_sended_set(mail_id):
                send_status = "发送成功"

            # 将按钮添加到单元格
            self.setCellWidget(index, 0, widget)
            self.setCellWidget(index, 1, button)
            self.setItem(index, 2, QTableWidgetItem(send_status))
            index += 1

    def refresh_table(self):
        for index in range(len(self.checkBoxList)):
            send_status = "未发送"
            if g_mail_common.is_in_sended_set(self.checkBoxList[index].mail_id):
                send_status = "发送成功"

            self.setItem(index, 2, QTableWidgetItem(send_status))
            self.checkBoxList[index].setChecked(False)


class MailIDCheckBox(QCheckBox):
    def __init__(self, mail_id, ):
        super().__init__(str())

        self.mail_id = mail_id

    def mail_select(self):
        # print(g_mail_common.selected_set)
        if self.checkState() == Qt.Checked:
            g_mail_common.add_selected_mail(self.mail_id)
        else:
            g_mail_common.remove_selected_mail(self.mail_id)

        # print(g_mail_common.selected_set)

class MailIDButton(QPushButton):
    def __init__(self, mail_obj):
        super().__init__(str(mail_obj.id))

        self.mail_obj = mail_obj

    def showInfo(self):
        self.infoTbl = MailInfoTable(self.mail_obj)

class MailInfoTable(QTableWidget):
    def __init__(self, mail_obj):
        super().__init__()
        self.mail_obj = mail_obj

        self.setWindowTitle("邮件-" + str(self.mail_obj.id))

        # 设置大小
        self.resize(1000, 800)

        v_head_list = ["邮件ID", "发件人", "收件人", "抄送人", "邮件主题", "附件信息", "邮件内容"]
        rowCnt = len(v_head_list)
        columCnt = 1
        self.setRowCount(rowCnt)
        self.setColumnCount(columCnt)
        self.setVerticalHeaderLabels(v_head_list)
        self.horizontalHeader().hide()
        # self.setColumnWidth(0, 800)

        self.verticalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.setAlternatingRowColors(True)

        # 设置表格是否可编辑
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        button = CMailDisplayButton(self.mail_obj.content)

        # 修改按钮大小
        button.setStyleSheet("QPushButton{margin:4px;font-size:24px};")
        button.clicked.connect(button.showInfo)

        self.setItem(0, 0, QTableWidgetItem(self.mail_obj.id))
        self.setItem(1, 0, QTableWidgetItem(self.mail_obj.sender))
        self.setItem(2, 0, QTableWidgetItem(";".join(self.mail_obj.receiver)))
        self.setItem(3, 0, QTableWidgetItem(self.mail_obj.cc))
        self.setItem(4, 0, QTableWidgetItem(self.mail_obj.subject))
        self.setItem(5, 0, QTableWidgetItem(str('\n'.join(self.mail_obj.appendix))))
        # self.setItem(6, 0, QTableWidgetItem(self.mail_obj.content))
        self.setCellWidget(6, 0, button)

        self.show()
class CMailDisplayButton(QPushButton):
    def __init__(self, mail_html):
        super().__init__(str("点击查看邮件正文"))

        self.mail_html = mail_html

    def showInfo(self):
        self.infoTbl = CMailDisplay(self.mail_html)

class CMailDisplay(QMainWindow):
    def __init__(self, mail_html):
        super(CMailDisplay, self).__init__()
        self.setWindowTitle('邮件内容')  #窗口标题
        self.setGeometry(30,30,1500,600)  #窗口的大小和位置设置
        self.browser=QWebEngineView()
        # 加载html代码(这里注意html代码是用三个
        self.browser.setHtml(mail_html)
        self.setCentralWidget(self.browser)
        self.show()

def proc():

    app = QApplication(sys.argv)
    ex = CAutoMail(g_mail_dict)

    write_log("AutoMail GUI init complete!")
    sys.exit(app.exec_())
