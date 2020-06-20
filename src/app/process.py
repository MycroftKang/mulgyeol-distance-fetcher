from PyQt5 import QtCore, QtGui, QtWidgets
import time, datetime, json

from app.worker import Worker
from app.product import PRODUCT_CONFIG

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

    def setValue(self, file_address, target, sheet_name, addr_column, distance_column, option, start_row, end_row, _stay_on_top):
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
        self.worker.setValue(file_address, target, sheet_name, addr_column, distance_column, option, start_row, end_row)
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
        self.setWindowTitle(PRODUCT_CONFIG['PRODUCT_NAME'])

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
        self.label1.setText(str(data[0]) + '개 항목에서 <a style="color:#0078D7;">{}</a>까지의 거리 정보 수집 중'.format(data[6]))
        self.label3.setText('수집 항목: ' + str(data[1]) + ' (' + str(data[2]) + 'km)')
        self.label4.setText('남은 시간: 약 ' + str(self.label4_last_time) + ' 남음')
        self.label5.setText('남은 항목: ' + str(data[4]) + '개')
        if data[0] == data[5]:
            self.label1.setText(str(self.emit_data[0]) + '개 항목에서 <a style="color:#0078D7;">{}</a>까지의  거리 정보 수집이 완료되었습니다.'.format(data[6]))
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
        elif check == 3:
            QtWidgets.QMessageBox.warning(self, 'Warning', '도착 위치를 찾을 수 없습니다.')
            self.close()

    def closeEvent(self, event):
        self.worker.fetcher.online = False
        self.thread.terminate()
        self.closesignal.emit(True)
        event.accept()
        print('Thread 종료가 확인되었습니다.')