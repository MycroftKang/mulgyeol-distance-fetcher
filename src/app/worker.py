from PyQt5 import QtCore, QtGui, QtWidgets
import pickle

from app.mapmecro import Mecro

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

    def setValue(self, file_address, target, sheet_name, addr_column, distance_column, option, start_row, end_row):
        self.file_address = file_address
        self.sheet_name = sheet_name
        self.addr_column = addr_column
        self.distance_column = distance_column
        self.start_row = start_row
        self.target = target
        self.option = option
        self.end_row = end_row

    def start_work(self):
        with open('..\\data\\bin\\config.bin', 'rb') as (f):
            config_list = pickle.load(f)
        self.fetcher.setValue(config_list, self.file_address, self.target, self.sheet_name, self.addr_column, self.distance_column, self.option, self.start_row, self.end_row)
        print('pass')
        self.fetcher.fetch()

    def get_percent_signal(self, percent):
        self.percent_signal.emit(percent)

    def driver_update(self):
        self.fetcher.update_driver()
        self.driver_update_signal.emit(2)