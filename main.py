# coding=utf-8
import sys
import os
import configparser
import xlrd
import logging
import datetime
from smtplib import SMTP
from email.header import Header
from email.mime.text import MIMEText
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

log_format = logging.Formatter('%(asctime)s %(levelname)s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

logger = logging.getLogger('main')
logger.setLevel(logging.DEBUG)

init_handler = logging.FileHandler('logs/main.log')
init_handler.setFormatter(log_format)
init_handler.setLevel(logging.INFO)
logger.addHandler(init_handler)

email_host = ''
email_user = ''
email_pass = ''

unit_groups = []


def set_logger(file_name):
    logger.handlers.clear()
    log_handler = logging.FileHandler(file_name)
    log_handler.setFormatter(log_format)
    log_handler.setLevel(logging.DEBUG)
    logger.addHandler(log_handler)


def send_mail(to_addrs, cc_addrs, subject, content, attach):
    try:
        email_client = SMTP(host=email_host)
        email_client.login(email_user, email_pass)
        msg = MIMEText(content, 'plain', 'utf-8')
        msg['from'] = email_user
        msg['to'] = to_addrs
        msg['Cc'] = cc_addrs
        msg['Subject'] = Header(subject, 'utf-8')
        email_client.sendmail(email_user, to_addrs, msg.as_string())
        email_client.quit()
        return True
    except Exception as e:
        print(str(e))
        return False


class Recipient:
    def __init__(self, full_name, unit_code, unit_school, name, cc, address, note):
        super().__init__()
        self.full_name = full_name
        self.unit_code = unit_code
        self.unit_school = unit_school
        self.name = name
        self.cc = cc
        self.address = address
        self.note = note

    def get_simple_str(self):
        if self.cc:
            return self.name + '(抄送,' + self.address + ')'
        return self.name + '(' + self.address + ')'


class Unit:
    def __init__(self, code, name):
        super().__init__()
        self.code = code
        self.name = name
        self.recipients = []

    def add_recipient(self, r):
        self.recipients.append(r)

    def get_recipients(self):
        return self.recipients

    def __str__(self, *args, **kwargs):
        return self.code + ': ' + self.name


class UGroup:
    def __init__(self, code, name):
        super().__init__()
        self.code = code
        self.name = name
        self.units = []

    def add_unit(self, u):
        self.units.append(u)

    def get_unit_str(self):
        unit_str = []
        for unit in self.units:
            unit_str.append(str(unit))
        return ', '.join(unit_str)

    def get_code_set(self):
        code_set = set()
        code_set.add(self.code)
        for u in self.units:
            code_set.add(u.code)
        return code_set

    def get_recipients(self):
        recipients = []
        for u in self.units:
            recipients.extend(u.get_recipients())
        return recipients

    def get_recipient_str(self):
        recipient_str = []
        for r in self.get_recipients():
            recipient_str.append(r.get_simple_str())
        return ', '.join(recipient_str)

    def __str__(self, *args, **kwargs):
        return self.code + '-' + self.name


def read_groups(file_path='conf/unit.xls'):
    group_dict = {}
    unit_dict = {}
    with xlrd.open_workbook(file_path) as data:
        unit_table = data.sheets()[0]
        rows = unit_table.nrows
        for i in range(1, rows):
            g_code = unit_table.cell(i, 0).value
            u_code = unit_table.cell(i, 2).value
            if u_code in unit_dict:
                continue
            # read groups
            unit = Unit(u_code, unit_table.cell(i, 3).value)
            unit_dict[u_code] = unit
            if g_code in group_dict:
                group = group_dict[g_code]
                group.units.append(unit)
            else:
                group = UGroup(g_code, unit_table.cell(i, 1).value)
                group.add_unit(unit)
                unit_groups.append(group)
                group_dict[g_code] = group
        # read recipients
        recipient_table = data.sheets()[1]
        rows = recipient_table.nrows
        for i in range(1, rows):
            full_name = recipient_table.cell(i, 0).value
            unit_code = recipient_table.cell(i, 1).value
            unit_name = recipient_table.cell(i, 2).value
            name = recipient_table.cell(i, 3).value
            cc = recipient_table.cell(i, 4).value == '是'
            address = recipient_table.cell(i, 5).value
            note = recipient_table.cell(i, 6).value
            recipient = Recipient(full_name, unit_code, unit_name, name, cc, address, note)
            unit = unit_dict[recipient.unit_code]
            unit.add_recipient(recipient)
    logger.info('读取院系联系人信息成功：共有%d个院系组，%d个院系，%d个联系人' %(len(group_dict), len(unit_dict), rows-1))
    return True


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.setWindowTitle('邮件助手')
        self.resize(800, 350)
        # 左右分割窗口
        main_splitter = QSplitter(Qt.Horizontal, self)
        left_widget = QWidget(main_splitter)
        right_widget = QWidget(main_splitter)
        # 左侧单位
        read_groups()
        vbox = QVBoxLayout()
        vbox.addWidget(QLabel('所有院系'))
        self.unit_table = QTableWidget(len(unit_groups), 4)
        self.unit_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.unit_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.unit_table.setSelectionMode(QTableWidget.SingleSelection)
        self.unit_table.setAlternatingRowColors(True)
        self.unit_table.setHorizontalHeaderLabels(['代码', '名称', '包含院系', '联系人'])
        i = 0
        for group in unit_groups:
            code_item = QTableWidgetItem(group.code)
            code_item.setTextAlignment(Qt.AlignCenter)
            self.unit_table.setItem(i, 0, code_item)
            self.unit_table.setItem(i, 1, QTableWidgetItem(group.name))
            self.unit_table.setItem(i, 2, QTableWidgetItem(group.get_unit_str()))
            self.unit_table.setItem(i, 3, QTableWidgetItem(group.get_recipient_str()))
            i += 1
        self.unit_table.resizeColumnsToContents()
        self.unit_table.resizeRowsToContents()
        self.unit_table.setColumnWidth(0, 60)
        self.unit_table.setColumnWidth(1, 120)
        self.unit_table.setColumnWidth(2, 140)
        self.unit_table.selectRow(0)
        vbox.addWidget(self.unit_table)
        left_widget.setLayout(vbox)

        grid_layout = QGridLayout()
        # 右侧主题
        grid_layout.addWidget(QLabel('主题：'), 0, 0, 1, 1)
        self.subject_editor = QLineEdit()
        grid_layout.addWidget(self.subject_editor, 0, 1, 1, 1)
        # 右侧正文
        grid_layout.addWidget(QLabel('正文：'), 1, 0, 1, 1)
        self.content_editor = QPlainTextEdit()
        grid_layout.addWidget(self.content_editor, 1, 1, 1, 1)
        # 右侧附件
        attach_widget = QWidget()
        hbox = QHBoxLayout()
        hbox.addWidget(QLabel('附件：'))
        self.attach_line = QLineEdit()
        self.attach_line.setReadOnly(True)
        hbox.addWidget(self.attach_line)
        attach_button = QPushButton("更改")
        attach_button.clicked.connect(self.change_path)
        hbox.addWidget(attach_button)
        attach_widget.setLayout(hbox)
        grid_layout.addWidget(attach_widget, 2, 0, 1, 2)
        send_button = QPushButton('发送')
        grid_layout.addWidget(send_button, 3, 0, 1, 2)
        send_button.clicked.connect(self.send_click)
        right_widget.setLayout(grid_layout)

        main_splitter.setStretchFactor(0, 7)
        self.setCentralWidget(main_splitter)

    def change_path(self):
        path_open = QFileDialog()
        attach_path = path_open.getExistingDirectory()
        self.attach_line.setText(attach_path)

    def send_click(self):
        dialog = DetailDialog(self)
        dialog.setModal(True)
        dialog.show()


class Attach:
    def __init__(self, code, name, p_path):
        super().__init__()
        self.code = code
        self.name = name
        self.p_path = p_path


class DetailDialog(QDialog):
    def __init__(self, parent=None):
        super(DetailDialog, self).__init__(parent)
        self.resize(500, 240)
        # the index and current group
        self.index = -1
        self.group = None
        self.count = 0
        self.recipient = []
        self.all_attach = []
        self.attach = []
        # the weights
        self.check_recipient = []
        self.check_attach = []
        grid_layout = QGridLayout()
        self.label = QLabel()
        grid_layout.addWidget(self.label, 0, 0, 1, 1)
        group_box = QGroupBox()
        self.group_layout = QGridLayout()
        group_box.setLayout(self.group_layout)
        grid_layout.addWidget(group_box, 1, 0, 1, 1)
        button_widget = QWidget()
        hbox = QHBoxLayout()
        hbox.setAlignment(Qt.AlignRight)
        self.next_button = QPushButton(' 确认，下一个 ')
        hbox.addWidget(self.next_button)
        skip_button = QPushButton(' 取消，下一个 ')
        skip_button.clicked.connect(self.skip)
        hbox.addWidget(skip_button)
        quit_button = QPushButton(' 退出 ')
        quit_button.clicked.connect(self.close)
        hbox.addWidget(quit_button)
        button_widget.setLayout(hbox)
        grid_layout.addWidget(button_widget, 2, 0, 1, 1)
        self.setLayout(grid_layout)
        # set logger
        set_logger('logs/%s.txt' % datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S'))
        # read attach
        self._read_attach()
        # update with index=0
        self.update(0)

    def skip(self):
        if self.index == len(unit_groups) - 1:
            msg = QMessageBox()
            msg.setWindowTitle('完成')
            msg.setText('共发送%d个单位的邮件，详细日志请查看%s' % (self.count, self.log_file))
            msg.exec_()
            self.close()
        else:
            logger.warn('跳过单位[%s]' % str(unit_groups[self.index]))
            self.update(self.index + 1)

    def update(self, index):
        self.index = index
        self.group = unit_groups[self.index]
        self.setWindowTitle(self.group.name)
        self._update_recipient()
        self._update_attach()
        self._update_label()
        self._update_check()

    def _update_label(self):
        self.label.setText('请选择单位[' + str(self.group) + ']的收件人和附件')

    def _update_check(self):
        # recipient
        for check_box in self.check_recipient:
            check_box.setParent(None)
        i = 0
        for recipient in self.recipient:
            check_box = QCheckBox(recipient.get_simple_str())
            check_box.setCheckState(Qt.Checked)
            self.check_recipient.append(check_box)
            self.group_layout.addWidget(check_box, i, 0, 1, 1)
            i += 1
        # attach
        for check_box in self.check_attach:
            check_box.setParent(None)
        i = 0
        for attach in self.attach:
            check_box = QCheckBox(attach.name)
            check_box.setCheckState(Qt.Checked)
            self.check_attach.append(check_box)
            self.group_layout.addWidget(check_box, i, 1, 1, 1)
            i += 1

    def _update_recipient(self):
        self.recipient = []
        self.recipient.extend(self.group.get_recipients())

    def _update_attach(self):
        self.attach = []
        g_code_set = self.group.get_code_set()
        g_code_set.add('000')
        for attach in self.all_attach:
            if attach.code in g_code_set:
                self.attach.append(attach)

    def _read_attach(self):
        a_path = self.parent().attach_line.text()
        if len(a_path) == 0:
            return
        logger.debug('读取附件: 开始遍历[%s]的所有文件' % a_path)
        for c_path in os.listdir(a_path):
            full_path = os.path.join(a_path, c_path)
            if os.path.isdir(full_path):
                logger.debug('读取附件: [%s]是目录跳过' % c_path)
                continue
            f_code = c_path[0:3]
            self.all_attach.append(Attach(f_code, c_path, a_path))
            logger.debug('读取附件: [%s]单位代码是[%s]' % (c_path, f_code))
        self.all_attach = sorted(self.all_attach, key= lambda x: x.code)


if __name__ == '__main__':
    cf = configparser.ConfigParser()
    cf.read('conf/server.ini', encoding='utf-8')
    email_host=cf.get('server_163', 'host')
    email_user=cf.get('server_163', 'user')
    email_pass=cf.get('server_163', 'password')
    # send_mail('498248393@qq.com,wangzhics@126.com', '你好', '利用此种方法（绿色代码部分）即可解决相关邮箱的554', [])
    app = QApplication(sys.argv)
    demo = MainWindow()
    demo.show()
    app.exec_()
