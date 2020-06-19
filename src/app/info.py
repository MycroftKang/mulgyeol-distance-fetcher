from PyQt5 import QtCore, QtGui, QtWidgets
import json

from app.product import PRODUCT_CONFIG

class App_Info(QtWidgets.QDialog):
    def __init__(self, _fixed_width):
        super().__init__()
        self.fixed_width = _fixed_width
        self.initUI()

    def initUI(self):
        with open("..\\info\\version.json", 'rt') as f:
            info = json.load(f)

        
        # self.commit_version = text[1].split('.')[3]+'.'+text[1].split('.')[4]

        pixmap = QtGui.QPixmap('..\\resources\\app\\Mulgyeol Labs CI 3.0.png')
        pixmap = pixmap.scaledToHeight(50)
        lbl_img = QtWidgets.QLabel()
        lbl_img.setPixmap(pixmap)
        pixmap2 = QtGui.QPixmap('..\\resources\\app\\MDF_Icon.png')
        pixmap2 = pixmap2.scaledToHeight(70)
        lbl_img2 = QtWidgets.QLabel()
        lbl_img2.setPixmap(pixmap2)
        label1 = QtWidgets.QLabel(PRODUCT_CONFIG['PRODUCT_NAME'])
        font = label1.font()
        font.setPointSize(11)
        font.setFamily('Segoe UI')
        label1.setFont(font)
        label3 = QtWidgets.QLabel('버전 {} <a href="https://github.com/MycroftKang/mulgyeol-distance-fetcher/releases">Release Note</a><br>커밋 {}'.format(info['version'], info['commit']))
        label3.setOpenExternalLinks(True)
        font2 = label3.font()
        font2.setFamily('맑은 고딕')
        font2.setPointSize(9)
        label3.setFont(font2)
        label2 = QtWidgets.QLabel('Copyright © 2020 Mulgyeol Labs. All Rights Reserved.')
        label2.setFont(font2)
        hbox1 = QtWidgets.QHBoxLayout()
        hbox1.addWidget(lbl_img)
        hbox1.addStretch(1)
        hbox1.addWidget(lbl_img2)
        vbox = QtWidgets.QVBoxLayout()
        vbox.addSpacing(10)
        vbox.addLayout(hbox1)
        vbox.addSpacing(15)
        vbox.addWidget(label1)
        vbox.addWidget(label3)
        vbox.addWidget(label2)
        vbox.addSpacing(10)
        hbox = QtWidgets.QHBoxLayout()
        hbox.addSpacing(20)
        hbox.addLayout(vbox)
        hbox.addSpacing(20)
        self.setLayout(hbox)
        self.setStyleSheet('background-color: #fff')
        self.setFixedSize(self.fixed_width, self.sizeHint().height())
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        self.setWindowIcon(QtGui.QIcon('..\\resources\\app\\MDF_Icon.png'))
        self.setWindowTitle('소프트웨어 정보')

    def show_window(self, _stay_on_top):
        if _stay_on_top:
            self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.CustomizeWindowHint | QtCore.Qt.WindowCloseButtonHint)
        else:
            self.setWindowFlags(QtCore.Qt.Window)
        self.show()