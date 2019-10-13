from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import time


class mainwin(QtWidgets.QWidget):

    def __init__(self):
        super(mainwin, self).__init__()
        self.setWindowTitle('Qthread')

        self.text_lable = QtWidgets.QLabel('0')
        self.number_lineedit = QtWidgets.QLineEdit()
        self.start_button = QtWidgets.QPushButton('Start')
        self.progressbar = QtWidgets.QProgressBar()
        self.progressbar.setValue(0)

        mainlayout = QtWidgets.QVBoxLayout()
        mainlayout.addWidget(self.start_button)
        mainlayout.addWidget(self.number_lineedit)
        mainlayout.addWidget(self.text_lable)
        mainlayout.addWidget(self.progressbar)

        self.setLayout(mainlayout)
        self.start_button.clicked.connect(self.run_qthread)
        self.t = thread()
        self.t.requestSignal_get.connect(self.get_lineedit_text)
        self.t.update_get.connect(self.set_lable)
        #set Lable vaule
        self.t.update_get.connect(self.set_progress)
        #set Progress bar total vaule
        self.t.requestSignal_get.connect(self.set_progress_max)
        #set Progress bar value

    def get_lineedit_text(self):
        self.t.update_get.emit(int(self.number_lineedit.text()))

    def run_qthread(self):
        self.t.start()

    def set_lable(self, number):
        self.text_lable.setText(str(number))

    def set_progress_max(self, max):
        self.progressbar.setMaximum(max)

    def set_progress(self, pnumber):
        self.progressbar.setValue(pnumber)


class thread(QtCore.QThread):
    requestSignal_get = QtCore.pyqtSignal()
    update_get = QtCore.pyqtSignal(int)

    def __init__(self):
        super(thread, self).__init__()
        self.update_get.connect(self.getnumber)
        self._getnumber = 0

    def getnumber(self, number):
        self._getnumber = number

    def run(self):
        self.requestSignal_get.emit()
        time.sleep(0.2)
        self.emit(QtCore.SIGNAL('maxpnumber'), self._getnumber)
        for i in range(0, self._getnumber+1):
            self.emit(QtCore.SIGNAL('number'), i)
            time.sleep(1)


def main():
    app = QtWidgets.QApplication(sys.argv)
    qm = mainwin()
    qm.show()
    app.exec_()


if __name__ == '__main__':
    main()
