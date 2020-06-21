import sys, os
from PyQt5 import QtCore, QtGui, QtWidgets
import pickle, win32.win32api as win32api
import time, datetime, openpyxl as excel
import sqlite3, json
import time

from update.UpdateDriver import UDriver
from app.setting import App_Setting
from app.info import App_Info
from app.process import App_Process

from app.product import PRODUCT_CONFIG
import win32.win32gui as win32gui
import win32.win32process as win32process

class Main_Window(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.auto_run = False
        with open('..\\data\\settings\\settings.json', 'rt', encoding='utf-8') as f:
            self.json_settings = json.load(f)
        self.initUI()

    def initUI(self):
        datalist = []
        with open('..\\data\\bin\\config.bin', 'rb') as (f):
            self.config_list = pickle.load(f)
        with open('..\\data\\bin\\viewconfig.bin', 'rt') as (f):
            viewconfig = f.read()

        con = sqlite3.connect('..\\data\\bin\\data.db')
        cur = con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS RecentFile(address text, sheet text, start_row text, end_row text, option integer, input_column text, output_column text, time real);")
        con.close()

        try:
            con = sqlite3.connect('..\\data\\bin\\data.db')
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

        self.settingsMenu = QtWidgets.QMenu('설정', self)
        configAction = QtWidgets.QAction('기본 설정', self)
        configAction.triggered.connect(self.showApp_Setting)
        devconfigAction = QtWidgets.QAction('구성 변경', self)
        devconfigAction.triggered.connect(self.show_json)
        self.settingsMenu.addAction(configAction)
        self.settingsMenu.addSeparator()
        self.settingsMenu.addAction(devconfigAction)
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
        oplAction = QtWidgets.QAction('법적 공지', self)
        oplAction.triggered.connect(self.show_OSL)
        releaseAction = QtWidgets.QAction('릴리스 정보', self)
        releaseAction.triggered.connect(self.show_release_info)
        reportissueAction = QtWidgets.QAction('문제 보고', self)
        reportissueAction.triggered.connect(self.show_issue)
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
        # filemenu.addAction(configAction)
        filemenu.addMenu(self.settingsMenu)
        filemenu.addSeparator()
        filemenu.addAction(exitAction)
        viewmenu = menubar.addMenu('보기')
        viewmenu.addAction(stayontopAction)
        helpmenu = menubar.addMenu('도움말')

        helpmenu.aboutToShow.connect(self.renew_update)
        helpmenu.addAction(releaseAction)
        helpmenu.addSeparator()
        helpmenu.addAction(reportissueAction)
        helpmenu.addSeparator()
        helpmenu.addAction(licenseAction)
        helpmenu.addAction(oplAction)
        helpmenu.addSeparator()
        helpmenu.addAction(self.updateAction)
        helpmenu.addSeparator()
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
        self.cbox1.addItem('추천경로')
        label7 = QtWidgets.QLabel('시작 행:')
        label7.setFont(font1)
        self.lineedit6 = QtWidgets.QLineEdit()
        self.lineedit6.setFont(font1)
        
        label8 = QtWidgets.QLabel('도착 위치:')
        label8.setFont(font1)
        self.lineedit7 = QtWidgets.QLineEdit()
        self.lineedit7.setFont(font1)
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
        self.setWindowTitle(PRODUCT_CONFIG['PRODUCT_NAME'])
        self.show()

    def show_issue(self):
        win32api.ShellExecute(0, 'open', 'https://github.com/MycroftKang/mulgyeol-distance-fetcher/issues', None, None, 1)

    def show_release_info(self):
        win32api.ShellExecute(0, 'open', 'https://github.com/MycroftKang/mulgyeol-distance-fetcher/releases', None, None, 1)

    def show_json(self):
        win32api.ShellExecute(0, 'open', '..\\data\\settings\\settings.json', None, ".", 1)

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
                self.auto_run = True
                self.close()
                sys.exit(0)

    def update_notice(self):
        if UDriver.isready():
            self.auto_run = True
            self.close()
        else:
            QtWidgets.QMessageBox.information(self, 'Mulgyeol Software Update', '현재 사용할 수 있는 업데이트가 없습니다.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)

    def renew_update(self):
        if UDriver.isnew():
            if UDriver.isready():
                self.updateAction.setText('다시 시작 및 업데이트')
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
        self.lineedit3.setText(data[5])
        self.lineedit4.setText(data[3])
        self.lineedit5.setText(data[6])     
        self.lineedit6.setText(data[2])
        self.cbox1.setCurrentIndex(data[4])
        # try:
        wb = excel.load_workbook(data[0])
        sheet_names = wb.sheetnames[:]
        self.completer = QtWidgets.QCompleter(sheet_names)
        self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.lineedit2.setCompleter(self.completer)
        wb.close()
        # except:
        #     pass

    def savenrun(self):
        with open('..\\data\\bin\\config.bin', 'rb') as (f):
            self.config_list = pickle.load(f)
        try:
            self.file_address = self.lineedit1.text()
            self.target = self.lineedit7.text()
            self.sheet_name = self.lineedit2.text()
            self.input_column = self.lineedit3.text()
            self.end_row = int(self.lineedit4.text())
            self.output_column = self.lineedit5.text()
            self.start_row = int(self.lineedit6.text())
            self.option = self.cbox1.currentIndex()
        except:
            QtWidgets.QMessageBox.warning(self, 'Warning', '입력 내용을 확인하십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
            return
        
        data = [self.file_address, self.target, self.sheet_name, str(self.start_row), str(self.end_row), self.input_column, self.output_column, self.option]
        
        if '' in data:
            QtWidgets.QMessageBox.warning(self, 'Warning', '입력 내용을 확인하십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
            return

        con = sqlite3.connect('..\\data\\bin\\data.db')
        cur = con.cursor()
        cur.execute("SELECT * FROM RecentFile WHERE address=?;", (self.file_address, ))
        tdatalist = cur.fetchall()

        if not len(tdatalist) == 0:
            cur.execute("UPDATE RecentFile SET sheet=?, start_row=?, end_row=?, option=? ,input_column=?, output_column=?, time=? WHERE address=?;", (self.sheet_name, str(self.start_row), str(self.end_row), self.option, self.input_column, self.output_column, time.time(), self.file_address))
        else:
            cur.execute("INSERT INTO RecentFile VALUES(?, ?, ?, ?, ?, ?, ?, ?);", (self.file_address, self.sheet_name, str(self.start_row), str(self.end_row), self.option, self.input_column, self.output_column, time.time()))
        con.commit()
        con.close()

        self.renew_recentfile()
        try:
            self.win_progress.setValue(self.file_address, self.target ,self.sheet_name, self.input_column, self.output_column, self.option, self.start_row, self.end_row, self.stay_on_top_flag)
        except Exception as e:
            print('[Error]:', e)
            QtWidgets.QMessageBox.warning(self, 'Warning', '입력 내용을 확인하십시오.', QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.Yes)
            return

        self.hide()
        self.win_progress.show()

    def show_OSL(self):
        win32api.ShellExecute(0, 'open', 'https://mgylabs.gitlab.io/mulgyeol-distance-fetcher/OpenSourceLicense.txt', None, None, 1)

    def show_SL(self):
        url = 'https://mgylabs.gitlab.io/mulgyeol-distance-fetcher/LICENSE'
        if (PRODUCT_CONFIG['LICENSE_URL'] != None) and (PRODUCT_CONFIG['LICENSE_URL'] != url):
            url = PRODUCT_CONFIG['LICENSE_URL']
        win32api.ShellExecute(0, 'open', url, None, None, 1)

    def set_lineedit7_text(self, platform):
        if self.json_settings['defalut']['enabled']:
            p_to_id = {0:"naver", 1:"kakao"}
            self.lineedit7.setText(self.json_settings['defalut'][p_to_id[platform]]['arrivallocation'])
            self.lineedit7.setReadOnly(self.json_settings['defalut'][p_to_id[platform]]['readonly'])
            self.cbox1.setCurrentIndex(self.json_settings['defalut'][p_to_id[platform]]['option'])
            
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
            if self.auto_run:
                UDriver.run_update_with_autorun()
            else:
                UDriver.run_update()
        event.accept()

class WindowMgr:
    def __init__ (self, lock_pid):
        self.lock_pid = lock_pid

    def _window_enum_callback(self, hwnd, wildcard):
        tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        title = win32gui.GetWindowText(hwnd)
        if (self.lock_pid == pid) and (title == wildcard):    
            win32gui.ShowWindow(hwnd, 9)
            win32gui.SetForegroundWindow(hwnd)
            win32gui.SetActiveWindow(hwnd)

    def find_window_wildcard(self, wildcard=PRODUCT_CONFIG['PRODUCT_NAME']):
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

if __name__ == '__main__':
    lockfile = QtCore.QLockFile(os.getenv('TEMP') + '/MDFE.lock')

    if not lockfile.tryLock():
        mgr = WindowMgr(lockfile.getLockInfo()[1])
        mgr.find_window_wildcard()
        sys.exit()
    
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