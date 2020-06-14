import sys, os
from PyQt5 import QtCore, QtGui, QtWidgets
import pickle, win32api
import time, datetime, openpyxl as excel
import sqlite3, json
import time

from app.UpdateDriver import UDriver
from app.setting_window import setting_dialog
from app.mapmecro import Mecro

class Main_Window(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        datalist = []
        with open('..\\data\\bin\\config.bin', 'rb') as (f):
            self.config_list = pickle.load(f)
        with open('..\\data\\bin\\viewconfig.bin', 'rt') as (f):
            viewconfig = f.read()

        con = sqlite3.connect('..\\data\\bin\\data.db')
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RecentFile(address text, sheet text, start_row text, end_row text, input_column text, output_column text, time real);")
        con.close()

        try:
            con = sqlite3.connect('..\\data\\data.db')
            cur = con.cursor()
            cur.execute("SELECT * FROM RecentFile ORDER BY time DESC;")
            datalist = cur.fetchall()
            con.close()
        except:
            pass

        widget = QtWidgets.QWidget()
        self.setCentralWidget(widget)
        loadAction = QtWidgets.QAction('파일 불러오기...', self)
        loadAction.triggered.connect(self.showDialog)
        openAction = QtWidgets.QAction('파일 열기', self)
        openAction.triggered.connect(self.open_file)
        self.loadrencentMenu = QtWidgets.QMenu('최근 항목 불러오기', self)
        for i in range(len(datalist)):
            loadrecentAct = QtWidgets.QAction(datalist[i][0], self)
            loadrecentAct.triggered.connect(lambda checked, item=datalist[i]: self.LoadRecentFile(item))
            self.loadrencentMenu.addAction(loadrecentAct)
        removerecentAct = QtWidgets.QAction('최근에 불러온 항목 지우기', self)
        removerecentAct.triggered.connect(self.remove_recentfile)
        self.loadrencentMenu.addSeparator()
        self.loadrencentMenu.addAction(removerecentAct)
        configAction = QtWidgets.QAction('설정', self)
        configAction.triggered.connect(self.showApp_Setting)
        exitAction = QtWidgets.QAction('끝내기', self)
        exitAction.triggered.connect(QtWidgets.qApp.quit)
        stayontopAction = QtWidgets.QAction('항상 위에 유지', self, checkable=True)
        if viewconfig == '0':
            stayontopAction.setChecked(False)
            self.setWindowFlags(QtCore.Qt.Window)
            self.stay_on_top_flag = False
        else:
            stayontopAction.setChecked(True)
            self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.CustomizeWindowHint | QtCore.Qt.WindowCloseButtonHint)
            self.stay_on_top_flag = True
        stayontopAction.setShortcut('Ctrl+T')
        stayontopAction.triggered.connect(self.stay_on_top)
        oplAction = QtWidgets.QAction('오픈소스 라이선스', self)
        oplAction.triggered.connect(self.show_OSL)
        licenseAction = QtWidgets.QAction('소프트웨어 라이선스', self)
        licenseAction.triggered.connect(self.show_SL)
        helpAction = QtWidgets.QAction('소프트웨어 정보', self)
        helpAction.triggered.connect(self.showApp_Info)
        self.updateAction = QtWidgets.QAction('업데이트 확인...', self)
        self.updateAction.triggered.connect(self.update_notice)
        menubar = self.menuBar()
        menubar.setNativeMenuBar(False)
        filemenu = menubar.addMenu('파일')
        filemenu.addAction(loadAction)
        filemenu.addAction(openAction)
        filemenu.addSeparator()
        filemenu.addMenu(self.loadrencentMenu)
        filemenu.addSeparator()
        filemenu.addAction(configAction)
        filemenu.addSeparator()
        filemenu.addAction(exitAction)
        viewmenu = menubar.addMenu('보기')
        viewmenu.addAction(stayontopAction)
        helpmenu = menubar.addMenu('도움말')

        helpmenu.aboutToShow.connect(self.renew_update)
        helpmenu.addAction(self.updateAction)
        helpmenu.addSeparator()
        helpmenu.addAction(oplAction)
        helpmenu.addSeparator()
        helpmenu.addAction(licenseAction)
        helpmenu.addAction(helpAction)
        
        label1 = QtWidgets.QLabel('파일 위치:')
        font1 = label1.font()
        font1.setFamily('맑은 고딕')
        font1.setPointSize(9)
        label1.setFont(font1)
        self.lineedit1 = QtWidgets.QLineEdit()
        self.lineedit1.setFont(font1)
        
        btn1 = QtWidgets.QPushButton('파일 찾기', self)
        btn1.setFont(font1)
        btn1.clicked.connect(self.showDialog)
        label2 = QtWidgets.QLabel('시트 이름:')
        label2.setFont(font1)
        self.lineedit2 = QtWidgets.QLineEdit()
        self.lineedit2.setFont(font1)

        label3 = QtWidgets.QLabel('입력 열:')
        label3.setFont(font1)
        self.lineedit3 = QtWidgets.QLineEdit()
        self.lineedit3.setFont(font1)
        
        label4 = QtWidgets.QLabel('끝 행:')
        label4.setFont(font1)
        self.lineedit4 = QtWidgets.QLineEdit()
        self.lineedit4.setFont(font1)
        
        label5 = QtWidgets.QLabel('출력 열:')
        label5.setFont(font1)
        self.lineedit5 = QtWidgets.QLineEdit()
        self.lineedit5.setFont(font1)

        label6 = QtWidgets.QLabel('검색 옵션:')
        label6.setFont(font1)
        self.cbox1 = QtWidgets.QComboBox(self)
        self.cbox1.setFont(font1)
        self.cbox1.addItem('최단거리')
        label7 = QtWidgets.QLabel('시작 행:')
        label7.setFont(font1)
        self.lineedit6 = QtWidgets.QLineEdit()
        self.lineedit6.setFont(font1)
        
        label8 = QtWidgets.QLabel('도착 위치:')
        label8.setFont(font1)
        self.lineedit7 = QtWidgets.QLineEdit()
        self.lineedit7.setFont(font1)
        self.lineedit7.setReadOnly(True)
        self.set_lineedit7_text(self.config_list[0])
        btn2 = QtWidgets.QPushButton('시작', self)
        btn2.setFont(font1)
        btn2.clicked.connect(self.savenrun)
        btn3 = QtWidgets.QPushButton('파일 열기', self)
        btn3.setFont(font1)
        btn3.clicked.connect(self.open_file)
        copyright_label = QtWidgets.QLabel('© 2020 Mulgyeol Labs (Developed by 강태혁)')
        font2 = copyright_label.font()
        font2.setFamily('맑은 고딕')
        font2.setPointSize(9)
        copyright_label.setFont(font2)
        grid = QtWidgets.QGridLayout()
        grid.addWidget(label1, 0, 0, 1, 1)
        grid.addWidget(self.lineedit1, 0, 1, 1, 7)
        grid.addWidget(btn1, 0, 8, 1, 1)
        grid.addWidget(label8, 1, 0, 1, 1)
        grid.addWidget(self.lineedit7, 1, 1, 1, 7)
        grid.addWidget(btn3, 1, 8, 1, 1)
        grid.addWidget(label2, 2, 0, 1, 1)
        grid.addWidget(self.lineedit2, 2, 1, 1, 3)
        grid.addWidget(label6, 2, 5, 1, 1)
        grid.addWidget(self.cbox1, 2, 6, 1, 3)
        grid.addWidget(label7, 3, 0, 1, 1)
        grid.addWidget(self.lineedit6, 3, 1, 1, 3)
        grid.addWidget(label3, 3, 5, 1, 1)
        grid.addWidget(self.lineedit3, 3, 6, 1, 3)
        grid.addWidget(label4, 4, 0, 1, 1)
        grid.addWidget(self.lineedit4, 4, 1, 1, 3)
        grid.addWidget(label5, 4, 5, 1, 1)
        grid.addWidget(self.lineedit5, 4, 6, 1, 3)
        grid.addWidget(copyright_label, 5, 0, 1, 7)
        grid.addWidget(btn2, 5, 8, 1, 1)
        blank = QtWidgets.QLabel()
        grid.addWidget(blank, 6, 0)
        grid.addWidget(blank, 6, 1)
        grid.addWidget(blank, 6, 3)
        grid.addWidget(blank, 6, 5)
        grid.addWidget(blank, 6, 6)
        grid.addWidget(blank, 6, 8)
        blank.hide()
        grid.setColumnStretch(0, 15)
        grid.setColumnStretch(1, 30)
        grid.setColumnStretch(3, 30)
        grid.setColumnStretch(5, 15)
        grid.setColumnStretch(6, 30)
        grid.setColumnStretch(8, 30)
        grid.setColumnStretch(2, 1)
        grid.setColumnStretch(4, 1)
        grid.setColumnStretch(7, 1)
        hbox = QtWidgets.QHBoxLayout()
        hbox.addSpacing(20)
        hbox.addLayout(grid)
        hbox.addSpacing(20)
        vbox = QtWidgets.QVBoxLayout()
        vbox.addSpacing(5)
        vbox.addLayout(hbox)
        vbox.addSpacing(5)
        widget.setLayout(vbox)
        self.fixed_width = self.sizeHint().width() * 1.2
        self.setFixedSize(self.fixed_width, self.sizeHint().height())
        self.win_progress = App_Process(self.fixed_width)
        self.info_win = App_Info(self.fixed_width)
        self.setting_win = App_Setting()
        self.setting_win.setting_signal.connect(self.set_lineedit7_text)
        self.win_progress.closesignal.connect(self.show_window, QtCore.Qt.QueuedConnection)
        self.setWindowIcon(QtGui.QIcon('..\\resources\\app\\MDF_Icon.png'))
        self.setWindowTitle('Mulgyeol Distance Fetcher')
        self.show()

    def stay_on_top(self, state):
        if state:
            self.stay_on_top_flag = True
            self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.CustomizeWindowHint | QtCore.Qt.WindowCloseButtonHint)
            with open('..\\data\\bin\\viewconfig.bin', 'wt') as (f):
                f.write('1')
            self.show()
        else:
            self.stay_on_top_flag = False
            self.setWindowFlags(QtCore.Qt.Window)
            with open('..\\data\\bin\\viewconfig.bin', 'wt') as (f):
                f.write('0')
            self.show()

    def open_file(self):
        file_address = self.lineedit1.text()
        if file_address:
            print('ctype', file_address)
            try:
                win32api.ShellExecute(0, 'open', file_address, None, None, 1)
            except:
                QtWidgets.QMessageBox.warning(self, 'Warning', '파일 위치를 확인하십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
        else:
            QtWidgets.QMessageBox.warning(self, 'Warning', '먼저 파일을 불러오십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)

    def start_update_notice(self):
        if UDriver.isready():
            reply = QtWidgets.QMessageBox.question(self, 'Mulgyeol Software Update', '최신버전의 소프트웨어 업데이트를 사용할 수  있습니다.\n업데이트를 시작하시겠습니까?', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.Yes)
            if reply == QtWidgets.QMessageBox.Yes:
                self.close()
                sys.exit(0)

    def update_notice(self):
        if UDriver.isready():
            self.close()
            sys.exit(0)
        else:
            QtWidgets.QMessageBox.information(self, 'Mulgyeol Software Update', '현재 사용할 수 있는 업데이트가 없습니다.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)

    def renew_update(self):
        if UDriver.isnew():
            if UDriver.isready():
                self.updateAction.setText('MDF(을)를 다시 시작하여 업데이트 적용')
                self.updateAction.setEnabled(True)
            else:
                self.updateAction.setText('업데이트 다운로드 중...')
                self.updateAction.setEnabled(False)

    def showDialog(self):
        fname = QtWidgets.QFileDialog.getOpenFileName(self, '파일 찾기', (os.getenv('USERPROFILE') + '\\Desktop'), filter='Excel  파일 (*.xl* *.xlsx *.xlsm, *.xlsb *.xlam *.xltx *.xltm *.xls *.xla *.xlt *.xlm *.xlw);;모든 파일 (*.*)')
        file_address = fname[0]
        if file_address:
            self.lineedit1.setText(file_address)
        try:
            wb = excel.load_workbook(file_address)
            sheet_names = wb.sheetnames[:]
            self.completer = QtWidgets.QCompleter(sheet_names)
            self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
            self.lineedit2.setCompleter(self.completer)
            wb.close()
        except:
            pass

    def LoadRecentFile(self, data):
        self.lineedit1.setText(data[0])
        self.lineedit2.setText(data[1])
        self.lineedit3.setText(data[4])
        self.lineedit4.setText(data[3])
        self.lineedit5.setText(data[5])     
        self.lineedit6.setText(data[2])
        try:
            wb = excel.load_workbook(data[0])
            sheet_names = wb.sheetnames[:]
            self.completer = QtWidgets.QCompleter(sheet_names)
            self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
            self.lineedit2.setCompleter(self.completer)
            wb.close()
        except:
            pass

    def savenrun(self):
        with open('..\\data\\bin\\config.bin', 'rb') as (f):
            self.config_list = pickle.load(f)
        try:
            self.file_address = self.lineedit1.text()
            self.sheet_name = self.lineedit2.text()
            self.input_column = self.lineedit3.text()
            self.end_row = int(self.lineedit4.text())
            self.output_column = self.lineedit5.text()
            self.start_row = int(self.lineedit6.text())
        except:
            QtWidgets.QMessageBox.warning(self, 'Warning', '입력 내용을 확인하십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
            return

        
        data = [self.file_address, self.sheet_name, str(self.start_row), str(self.end_row), self.input_column, self.output_column]
        
        if '' in data:
            QtWidgets.QMessageBox.warning(self, 'Warning', '입력 내용을 확인하십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
            return

        con = sqlite3.connect('..\\data\\bin\\data.db')
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RecentFile(address text, sheet text, start_row text, end_row text, input_column text, output_column text, time real);")
        cur.execute("SELECT * FROM RecentFile WHERE address=?;", (self.file_address, ))
        tdatalist = cur.fetchall()

        if not len(tdatalist) == 0:
            cur.execute("UPDATE RecentFile SET sheet=?, start_row=?, end_row=?, input_column=?, output_column=?, time=? WHERE address=?;", (self.sheet_name, str(self.start_row), str(self.end_row), self.input_column, self.output_column, time.time(), self.file_address))
        else:
            cur.execute("INSERT INTO RecentFile VALUES(?, ?, ?, ?, ?, ?, ?);", (self.file_address, self.sheet_name, str(self.start_row), str(self.end_row), self.input_column, self.output_column, time.time()))
        con.commit()
        con.close()

        self.renew_recentfile()
        try:
            self.win_progress.setValue(self.file_address, self.sheet_name, self.input_column, self.output_column, self.start_row, self.end_row, self.stay_on_top_flag)
        except Exception as e:
            print('[Error]:', e)
            QtWidgets.QMessageBox.warning(self, 'Warning', '입력 내용을 확인하십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
            return

        self.hide()
        self.win_progress.show()

    def show_OSL(self):
        win32api.ShellExecute(0, 'open', 'https://drive.google.com/open?id=1l7m0jminM4ru3k2E1DSFNJv_yFONgrhq', None, None, 1)

    def show_SL(self):
        win32api.ShellExecute(0, 'open', 'https://drive.google.com/open?id=1P3V6Hk3vuM__AIRlCqqco4kKw2HSWuvU', None, None, 1)

    def set_lineedit7_text(self, platform):
        if platform == 0:
            self.lineedit7.setText('화성고등학교정문')
        else:
            self.lineedit7.setText('경기 화성시 향남읍 장짐길 4')

    def show_window(self, condition=False):
        if condition:
            self.show()

    def showApp_Info(self):
        self.info_win.show_window(self.stay_on_top_flag)

    def showApp_Setting(self):
        self.setting_win.show(self.stay_on_top_flag)

    def remove_recentfile(self):
        con = sqlite3.connect('..\\data\\bin\\data.db')
        cur = con.cursor()
        cur.execute("DELETE FROM RecentFile;")
        con.commit()
        con.close()
        self.renew_recentfile()

    def renew_recentfile(self):
        for menu in self.loadrencentMenu.actions():
            self.loadrencentMenu.removeAction(menu)

        try:
            con = sqlite3.connect('..\\data\\bin\\data.db')
            cur = con.cursor()
            cur.execute("SELECT * FROM RecentFile ORDER BY time DESC;")
            datalist = cur.fetchall()
            con.close()
        except:
            pass

        for i in range(len(datalist)):
            loadrecentAct = QtWidgets.QAction(datalist[i][0], self)
            loadrecentAct.triggered.connect(lambda checked, item=datalist[i]: self.LoadRecentFile(item))
            self.loadrencentMenu.addAction(loadrecentAct)
        removerecentAct = QtWidgets.QAction('최근에 불러온 항목 지우기', self)
        removerecentAct.triggered.connect(self.remove_recentfile)
        self.loadrencentMenu.addSeparator()
        self.loadrencentMenu.addAction(removerecentAct)

    def closeEvent(self, event):
        if UDriver.isready():
            UDriver.run_update()
        event.accept()

class App_Setting(setting_dialog):
    setting_signal = QtCore.pyqtSignal(int)

    def __init__(self):
        self.dialog = QtWidgets.QDialog()
        super().setupUi(self.dialog)
        self.dialog.setWindowIcon(QtGui.QIcon('..\\resources\\app\\MDF_Icon.png'))
        with open('..\\data\\bin\\config.bin', 'rb') as (f):
            self.config_list = pickle.load(f)
        self.pushButton.clicked.connect(self.save_setting)
        self.pushButton_2.clicked.connect(self.close)
        self.setDefaultValue()

    def setDefaultValue(self):
        if self.config_list[0] == 0:
            self.radioButton_naver.setChecked(True)
        else:
            self.radioButton_2_kakao.setChecked(True)
        self.waittime_start.setValue(self.config_list[1])
        self.waittime_end.setValue(self.config_list[2])
        self.waittime_item.setValue(self.config_list[3])
        self.waittime_item_sec.setValue(self.config_list[4])
        if self.config_list[5] == 0:
            self.radioButton_finish.setChecked(True)
        else:
            self.radioButton_wait_again.setChecked(True)
        self.wait_again_min.setValue(self.config_list[6])

    def save_setting(self):
        if self.radioButton_naver.isChecked():
            platform = 0
        else:
            platform = 1
        if self.radioButton_finish.isChecked():
            network = 0
        else:
            network = 1
        renew_config_list = [platform, self.waittime_start.value(), self.waittime_end.value(), self.waittime_item.value(), self.waittime_item_sec.value(), network, self.wait_again_min.value()]
        with open('..\\data\\bin\\config.bin', 'wb') as (f):
            pickle.dump(renew_config_list, f)
        self.setting_signal.emit(platform)
        self.close()

    def show(self, _stay_on_top):
        if _stay_on_top:
            self.dialog.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.CustomizeWindowHint | QtCore.Qt.WindowCloseButtonHint)
        else:
            self.dialog.setWindowFlags(QtCore.Qt.Window)
        self.dialog.show()

    def close(self):
        self.dialog.close()


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
        label1 = QtWidgets.QLabel('Mulgyeol Distance Fetcher Education')
        font = label1.font()
        font.setPointSize(11)
        font.setFamily('Segoe UI')
        label1.setFont(font)
        label3 = QtWidgets.QLabel('[교육기관용] '+ info['version'] +' <a href="https://github.com/MycroftKang/mulgyeol-distance-fetcher/releases">Release Note</a>')
        label3.setOpenExternalLinks(True)
        font2 = label3.font()
        font2.setFamily('맑은 고딕')
        font2.setPointSize(9)
        label3.setFont(font2)
        label2 = QtWidgets.QLabel('이 소프트웨어에 대한 저작권은 강태혁에게 있습니다.\nCopyright © 2020 Mulgyeol Labs. All Rights Reserved.')
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


class App_Process(QtWidgets.QMainWindow):
    closesignal = QtCore.pyqtSignal(bool)

    def __init__(self, _fixed_width):
        super().__init__()
        self.fixed_width = _fixed_width
        self.worker = Worker()
        self.thread = QtCore.QThread()
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.start_work)
        self.worker.percent_signal.connect(self.info_system, QtCore.Qt.QueuedConnection)
        self.worker.label_signal.connect(self.setLabel, QtCore.Qt.QueuedConnection)
        self.worker.permission_Signal.connect(self.permission_error, QtCore.Qt.QueuedConnection)
        self.worker.wait_signal.connect(self.wait_info, QtCore.Qt.QueuedConnection)
        self.initUI()

    def setValue(self, file_address, sheet_name, addr_column, distance_column, start_row, end_row, _stay_on_top):
        if _stay_on_top:
            self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.CustomizeWindowHint | QtCore.Qt.WindowCloseButtonHint)
        else:
            self.setWindowFlags(QtCore.Qt.Window)
        self.pbar.setRange(0, 0)
        self.pbar.setValue(0)
        self.emit_data = None
        self.window_resize_fewer()
        self.label1.setText('소프트웨어를 구성하고 있습니다.')
        self.label2.setText('0% 완료')
        self.label3.setText('출발 위치: 불러오는 중...')
        self.label4.setText('남은 시간: 계산 중...')
        self.label5.setText('남은 항목: 계산 중...')
        self.btn1.setEnabled(False)
        self.worker.setValue(file_address, sheet_name, addr_column, distance_column, start_row, end_row)
        self.worker.fetcher.online = True
        self.st = time.time()
        self.thread.start()

    def initUI(self):
        widget = QtWidgets.QWidget()
        self.setCentralWidget(widget)
        self.pbar = QtWidgets.QProgressBar(self)
        self.pbar.setTextVisible(False)
        self.pbar.setRange(0, 0)
        self.label1 = QtWidgets.QLabel('소프트웨어를 구성하고 있습니다.')
        font1 = self.label1.font()
        font1.setFamily('맑은 고딕')
        font1.setPointSize(9)
        self.label1.setFont(font1)
        self.label2 = QtWidgets.QLabel('0% 완료')
        font2 = self.label2.font()
        font2.setFamily('맑은 고딕')
        font2.setPointSize(15)
        self.label2.setFont(font2)
        self.label3 = QtWidgets.QLabel('출발 위치: 불러오는 중...')
        self.label3.setFont(font1)
        self.label4 = QtWidgets.QLabel('남은 시간: 계산 중...')
        self.label4.setFont(font1)
        self.label5 = QtWidgets.QLabel('남은 항목: 계산 중...')
        self.label5.setFont(font1)
        self.btn1 = QtWidgets.QPushButton('완료', self)
        self.btn1.setFont(font1)
        self.btn1.setStyleSheet('background-color: light gray')
        self.btn1.clicked.connect(self.close)
        self.btn2 = QtWidgets.QPushButton('자세히', self)
        self.btn2.setFont(font1)
        self.btn2.setStyleSheet('background-color: light gray')
        self.btn2.clicked.connect(self.window_resize_more)
        hbox = QtWidgets.QHBoxLayout()
        hbox.addStretch(3)
        hbox.addWidget(self.btn2)
        hbox.addWidget(self.btn1)
        vbox = QtWidgets.QVBoxLayout()
        vbox.addSpacing(5)
        vbox.addWidget(self.label1)
        vbox.addSpacing(-6)
        vbox.addWidget(self.label2)
        vbox.addSpacing(10)
        vbox.addWidget(self.pbar)
        vbox.addSpacing(10)
        vbox.addWidget(self.label3)
        vbox.addWidget(self.label4)
        vbox.addWidget(self.label5)
        self.label3.hide()
        self.label4.hide()
        self.label5.hide()
        vbox.addLayout(hbox)
        vbox.addSpacing(5)
        hbox2 = QtWidgets.QHBoxLayout()
        hbox2.addSpacing(20)
        hbox2.addLayout(vbox)
        hbox2.addSpacing(20)
        widget.setLayout(hbox2)
        self.setStyleSheet('background-color: #fff')
        self.setFixedSize(self.fixed_width, self.sizeHint().height())
        self.init_window_size = self.size()
        self.setWindowIcon(QtGui.QIcon('..\\resources\\app\\MDF_Icon.png'))
        self.setWindowTitle('Mulgyeol Distance Fetcher')

    def window_resize_more(self):
        self.btn2.setText('간단히')
        self.btn2.disconnect()
        self.btn2.clicked.connect(self.window_resize_fewer)
        self.label3.show()
        self.label4.show()
        self.label5.show()
        self.setFixedSize(self.fixed_width, self.sizeHint().height())

    def window_resize_fewer(self):
        self.btn2.setText('자세히')
        self.btn2.disconnect()
        self.btn2.clicked.connect(self.window_resize_more)
        self.setFixedSize(self.init_window_size)
        self.label3.hide()
        self.label4.hide()
        self.label5.hide()

    def setLabel(self, data):
        self.emit_data = data
        self.label4_last_time = datetime.timedelta(seconds=(round(data[3])))
        self.label1.setText(str(data[0]) + '개 항목에서 <a style="color:#0078D7;">화성고등학교</a>까지의 거리 정보 수집 중')
        self.label3.setText('수집 항목: ' + str(data[1]) + ' (' + str(data[2]) + 'km)')
        self.label4.setText('남은 시간: 약 ' + str(self.label4_last_time) + ' 남음')
        self.label5.setText('남은 항목: ' + str(data[4]) + '개')
        if data[0] == data[5]:
            self.label1.setText(str(self.emit_data[0]) + '개 항목에서 <a style="color:#0078D7;">화성고등학교</a>까지의  거리 정보 수집이 완료되었습니다.')
            self.label4.setText('소요 시간: ' + str(datetime.timedelta(seconds=(round(self.et - self.st)))) + ' 소요됨')

    def info_system(self, percent):
        if percent == 0:
            self.pbar.setRange(0, 100)
        self.pbar.setValue(percent)
        self.label2.setText(str(percent) + '% 완료')
        if percent == 100:
            self.btn1.setEnabled(True)
            self.et = time.time()
            self.thread.terminate()
            print('Thread 종료가 감지되었습니다.')

    def wait_info(self, time):
        self.label1.setText('네트워크 오류(Abusing 감지)가 발생하여 ' + str(time) + '분 후 다시 시도합니다.')
        self.label3.setText('수집 항목: 대기 중...')
        try:
            self.label4.setText('남은 시간: ' + str(self.label4_last_time + datetime.timedelta(minutes=time)) + ' 남음')
        except:
            pass

    def permission_error(self, check=0):
        if check == 1:
            QtWidgets.QMessageBox.warning(self, 'Warning', '네트워크 오류(Abusing 감지)가 발생했습니다.\n잠시 후 다시 시도하십시오.')
            self.close()
        elif check == 2:
            QtWidgets.QMessageBox.warning(self, 'Warning', 'Excel 파일을 닫고, 다시 시도하십시오.')
            self.close()

    def closeEvent(self, event):
        self.worker.fetcher.online = False
        self.thread.terminate()
        self.closesignal.emit(True)
        event.accept()
        print('Thread 종료가 확인되었습니다.')


class Worker(QtCore.QObject):
    percent_signal = QtCore.pyqtSignal(int)
    label_signal = QtCore.pyqtSignal(list)
    permission_Signal = QtCore.pyqtSignal(int)
    end_signal = QtCore.pyqtSignal()
    wait_signal = QtCore.pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.fetcher = Mecro('..\\data\\log\\MDF.log')
        self.fetcher.percentChanged.connect(self.get_percent_signal, QtCore.Qt.QueuedConnection)
        self.fetcher.completion.connect(self.label_signal.emit, QtCore.Qt.QueuedConnection)
        self.fetcher.permission.connect(self.permission_Signal.emit, QtCore.Qt.QueuedConnection)
        self.fetcher.waitsignal.connect(self.wait_signal.emit, QtCore.Qt.QueuedConnection)

    def setValue(self, file_address, sheet_name, addr_column, distance_column, start_row, end_row):
        self.file_address = file_address
        self.sheet_name = sheet_name
        self.addr_column = addr_column
        self.distance_column = distance_column
        self.start_row = start_row
        self.end_row = end_row

    def start_work(self):
        with open(os.getenv('LOCALAPPDATA') + '\\Mulgyeol\\Mulgyeol Distance Fetcher\\User Data\\bin\\config.bin', 'rb') as (f):
            config_list = pickle.load(f)
        self.fetcher.setValue(config_list, self.file_address, self.sheet_name, self.addr_column, self.distance_column, self.start_row, self.end_row)
        print('pass')
        self.fetcher.fetch()

    def get_percent_signal(self, percent):
        self.percent_signal.emit(percent)

    def driver_update(self):
        self.fetcher.update_driver()
        self.driver_update_signal.emit(2)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    splash_pix = QtGui.QPixmap('..\\resources\\app\\Mulgyeol Labs splash.png')
    splash_pix = splash_pix.scaledToHeight(350)
    splash = QtWidgets.QSplashScreen(splash_pix, QtCore.Qt.WindowStaysOnTopHint)
    splash.show()

    UDriver.check_update()
    ex1 = Main_Window()
    splash.finish(ex1)
    ex1.start_update_notice()
    sys.exit(app.exec_())