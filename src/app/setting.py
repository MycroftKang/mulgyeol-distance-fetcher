from PyQt5 import QtCore, QtGui, QtWidgets
import pickle
from app.setting_window import setting_dialog

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