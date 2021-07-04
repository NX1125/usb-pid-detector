import os
import re
import subprocess
import sys
import traceback
from io import BytesIO

from PyQt5.QtCore import QObject, pyqtSlot
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, QLineEdit

try:
    from subprocess import DEVNULL
except ImportError:
    DEVNULL = os.open(os.devnull, os.O_RDWR)

p = re.compile(rb'(?:^|[\\&])PID_([A-Za-z0-9]+)(?:$|[\\&])')

DEFAULT_PID = '5002'


class App(QObject):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.text_box: QLineEdit = None
        self.label: QLabel = None

    @pyqtSlot()
    def on_click(self):
        self.label.setText('CHECKING')
        self.label.setStyleSheet('color:blue;')
        app.processEvents()
        # print('Checking...')
        try:
            with open('err.txt', 'w', encoding='utf-8') as file:
                script = "Get-PnpDevice -PresentOnly | Select-Object -Property InstanceId"
                process = subprocess.Popen(f'powershell.exe -command "{script}"',
                                           # creationflags=subprocess.CREATE_NO_WINDOW,
                                           stdout=subprocess.PIPE,
                                           stdin=subprocess.DEVNULL,
                                           stderr=file,
                                           shell=True)
                out, err = process.communicate()
                data = out
        except BaseException:
            with open('log.txt', 'w', encoding='utf-8') as file:
                traceback.print_exc(file=file)
            label.setStyleSheet('color:black;')
            label.setText('SEE log.txt')
            traceback.print_exc()
            return
        lines: set[bytes] = {line.strip() for line in data.splitlines()}

        devices_pid = set()
        for line in lines:
            if not line.startswith(b'USB\\'):
                continue
            match = p.search(line)
            if match:
                devices_pid.add(match.group(1))

        print(f'There are {len(devices_pid)} devices')
        for pid in devices_pid:
            print(f'PID: {pid}')

        required_pid = (self.text_box.text() or str(DEFAULT_PID)).encode('utf-8')
        if required_pid in devices_pid:
            self.label.setStyleSheet('color:green;')
            self.label.setText('PASSED')
            print('PID Found!')
        else:
            self.label.setStyleSheet('color:red;')
            self.label.setText('FAIL')
            print('PID not Found!')


@pyqtSlot()
def on_return_press():
    button.clicked.emit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = QMainWindow()
    window.setWindowTitle('Check USB by PID')

    my_app = App(window)

    text_box = QLineEdit(str(DEFAULT_PID))
    text_box.setFixedHeight(25)
    text_box.setPlaceholderText(f'PID (default {DEFAULT_PID})')

    button = QPushButton('CHECK')
    button.clicked.connect(my_app.on_click)

    text_box.returnPressed.connect(on_return_press)

    label = QLabel('PRESS CHECK')
    label.setStyleSheet('color: blue;')
    label.setAlignment(Qt.AlignHCenter)

    layout = QVBoxLayout()

    layout.addWidget(text_box)
    layout.addWidget(button)
    layout.addWidget(label)

    widget = QWidget()
    widget.setLayout(layout)

    window.setCentralWidget(widget)
    window.adjustSize()

    my_app.text_box = text_box
    my_app.label = label

    window.show()

    app.exec_()
