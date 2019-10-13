import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

import queue  # If this template is not loaded, pyinstaller may not be able to run the requests template after packaging
import requests
################################################


################################################
class Widget(QWidget):
    def __init__(self, *args, **kwargs):
        super(Widget, self).__init__(*args, **kwargs)
        layout = QHBoxLayout(self)

        # Increase progress bar
        self.progressBar = QProgressBar(self, minimumWidth=400)
        self.progressBar.setValue(0)
        layout.addWidget(self.progressBar)

        # Add Download Button
        self.pushButton = QPushButton(self, minimumWidth=100)
        self.pushButton.setText("download")
        layout.addWidget(self.pushButton)

        # Binding Button Event
        self.pushButton.clicked.connect(self.on_pushButton_clicked)

    # Download button event

    def on_pushButton_clicked(self):
        the_url = 'http://cdn2.ime.sogou.com/b24a8eb9f06d6bfc08c26f0670a1feca/5c9de72d/dl/index/1553820076/sogou_pinyin_93e.exe'
        the_filesize = requests.get(
            the_url, stream=True).headers['Content-Length']
        the_filepath = "D:/sogou_pinyin_93e.exe"
        the_fileobj = open(the_filepath, 'wb')
        #### Create a download thread
        self.downloadThread = downloadThread(
            the_url, the_filesize, the_fileobj, buffer=10240)
        self.downloadThread.download_proess_signal.connect(
            self.set_progressbar_value)
        self.downloadThread.start()

    # Setting progress bar

    def set_progressbar_value(self, value):
        self.progressBar.setValue(value)
        if value == 100:
            QMessageBox.information(self, "Tips", "Download success!")
            return


##################################################################
#Download thread
##################################################################
class downloadThread(QThread):
    download_proess_signal = pyqtSignal(int)  # Create signal

    def __init__(self, url, filesize, fileobj, buffer):
        super(downloadThread, self).__init__()
        self.url = url
        self.filesize = filesize
        self.fileobj = fileobj
        self.buffer = buffer

    def run(self):
        try:
            # Streaming download mode
            rsp = requests.get(self.url, stream=True)
            offset = 0
            for chunk in rsp.iter_content(chunk_size=self.buffer):
                if not chunk:
                    break
                self.fileobj.seek(offset)  # Setting Pointer Position
                self.fileobj.write(chunk)  # write file
                offset = offset + len(chunk)
                proess = offset / int(self.filesize) * 100
                self.download_proess_signal.emit(int(proess))  # Sending signal
            #######################################################################
            self.fileobj.close()  # Close file
            self.exit(0)  # Close thread

        except Exception as e:
            print(e)


####################################
#Program entry
####################################
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Widget()
    w.show()
    sys.exit(app.exec_())
