from PyQt5 import QtCore, QtWidgets,QtGui
from PyQt5.QtWidgets import *

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QApplication.translate(context, text, disambig)

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName(_fromUtf8("Form"))
        Form.resize(829, 521)
        self.centralwidget = QtWidgets.QWidget(Form)
        self.path_excel_label = QtWidgets.QLabel(Form)
        self.path_excel_label.setGeometry(QtCore.QRect(120, 320, 101, 20))
        self.path_excel_label.setObjectName(_fromUtf8("path_excel_label"))
        self.email_id_label = QtWidgets.QLabel(Form)
        self.email_id_label.setGeometry(QtCore.QRect(120, 80, 66, 17))
        self.email_id_label.setObjectName(_fromUtf8("email_id_label"))
        self.password_label = QtWidgets.QLabel(Form)
        self.password_label.setGeometry(QtCore.QRect(120, 130, 66, 17))
        self.password_label.setObjectName(_fromUtf8("password_label"))
        self.subject_label = QtWidgets.QLabel(Form)
        self.subject_label.setGeometry(QtCore.QRect(120, 180, 66, 17))
        self.subject_label.setObjectName(_fromUtf8("subject_label"))
        self.body_label = QtWidgets.QLabel(Form)
        self.body_label.setGeometry(QtCore.QRect(120, 230, 141, 31))
        self.body_label.setObjectName(_fromUtf8("body_label"))
        self.email_id_text = QtWidgets.QLineEdit(Form)
        self.email_id_text.setGeometry(QtCore.QRect(250, 80, 331, 27))
        self.email_id_text.setObjectName(_fromUtf8("email_id_text"))
        self.password_text = QtWidgets.QLineEdit(Form)
        self.password_text.setEchoMode(QLineEdit.Password)
        self.password_text.setGeometry(QtCore.QRect(250, 130, 331, 27))
        self.password_text.setObjectName(_fromUtf8("password_text"))
        self.subject_text = QtWidgets.QLineEdit(Form)
        self.subject_text.setGeometry(QtCore.QRect(250, 180, 441, 27))
        self.subject_text.setObjectName(_fromUtf8("subject_text"))
        self.main_label = QtWidgets.QLabel(Form)
        self.main_label.setGeometry(QtCore.QRect(370, 20, 120, 25))
        self.main_label.setStyleSheet(_fromUtf8("font: 63 20pt \"Ubuntu\";"))
        self.main_label.setObjectName(_fromUtf8("main_label"))
        self.send_button = QPushButton(Form)
        self.send_button.setGeometry(QtCore.QRect(340, 430, 161, 51))
        self.send_button.setObjectName(_fromUtf8("send_button"))
        self.browse_button_excel = QPushButton(Form)
        self.browse_button_excel.setGeometry(QtCore.QRect(250, 320, 331, 27))
        self.browse_button_excel.setObjectName(_fromUtf8("browse_button_excel"))
        self.browse_button_mail = QPushButton(Form)
        self.browse_button_mail.setGeometry(QtCore.QRect(250, 230, 331, 27))
        self.browse_button_mail.setObjectName(_fromUtf8("browse_button_mail"))
        self.Open_word_button = QPushButton(Form)
        self.Open_word_button.setGeometry(QtCore.QRect(590, 260, 101, 31))
        self.Open_word_button.setObjectName(_fromUtf8("Open_word_button"))
        self.Open_word_button.setDisabled(True)
        self.Open_excel_button = QPushButton(Form)
        self.Open_excel_button.setGeometry(QtCore.QRect(590, 290, 101, 31))
        self.Open_excel_button.setObjectName(_fromUtf8("Open_excel_button"))
        self.Open_excel_button.setDisabled(True)


        # Radio button for Attach Existing attachment
        self.radioButton_existing = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_existing.setGeometry(QtCore.QRect(120, 280, 150, 20))
        self.radioButton_existing.toggled.connect(self.attach_selected)

        # Radio button for Creting attachment
        self.radioButton_create = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_create.setGeometry(QtCore.QRect(400, 280, 150, 20))
        self.radioButton_create.toggled.connect(self.create_selected)

        # self.label = QtWidgets.QLabel(self.centralwidget)
        # self.label.setGeometry(QtCore.QRect(170, 90, 211, 20))
        # MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def create_selected(self, selected):
        if selected:
            self.Open_word_button.setDisabled(False)
            self.Open_excel_button.setDisabled(False)

    def attach_selected(self, selected):
        if selected:
            self.Open_word_button.setDisabled(True)
            self.Open_excel_button.setDisabled(True)

    def retranslateUi(self, Form):
        Form.setWindowTitle(_translate("Form", "Multimail", None))
        self.path_excel_label.setText(_translate("Form", "Attachments", None))
        self.email_id_label.setText(_translate("Form", "Email ID", None))
        self.password_label.setText(_translate("Form", "Password", None))
        self.subject_label.setText(_translate("Form", "Subject", None))
        self.body_label.setText(_translate("Form", "Body", None))
        self.main_label.setText(_translate("Form", "MultiMail", None))
        self.send_button.setText(_translate("Form", "Send Mail", None))
        self.browse_button_excel.setText(_translate("Form", "Attach", None))
        self.browse_button_mail.setText(_translate("Form", "Create Body", None))
        self.Open_word_button.setText(_translate("Form", "Create file", None))
        self.Open_excel_button.setText(_translate("Form", "Resource", None))
        self.radioButton_existing.setText(_translate("Form", "Attach Existing Attachments",None))
        self.radioButton_create.setText(_translate("Form", "Create Attachments",None))
