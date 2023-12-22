from PyQt6 import QtWidgets, QtCore
from random import randint

class MyWindow(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__()
        QtWidgets.QWidget.__init__(self, parent)
        self.label_cube = QtWidgets.QLabel("Number of cubes:")
        self.label_cube.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        self.lineEdit_1 = QtWidgets.QLineEdit(self)
        self.label_throw = QtWidgets.QLabel("Number of throws:")
        self.label_cube.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        self.lineEdit_2 = QtWidgets.QLineEdit(self)
        self.button_start = QtWidgets.QPushButton("OK")
        self.button_restart = QtWidgets.QPushButton("Reset")
        self.button_exit = QtWidgets.QPushButton("Canceal")
        self.textEdit = QtWidgets.QTextEdit(self)
        self.vbox = QtWidgets.QVBoxLayout()
        self.vbox.addWidget(self.label_cube)
        self.vbox.addWidget(self.lineEdit_1)
        self.vbox.addWidget(self.label_throw)
        self.vbox.addWidget(self.lineEdit_2)
        self.vbox.addWidget(self.textEdit)
        self.vbox.addWidget(self.button_start)
        self.vbox.addWidget(self.button_restart)
        self.vbox.addWidget(self.button_exit)
        self.setLayout(self.vbox)
        self.button_start.clicked.connect(self.on_clicked)
        self.button_restart.clicked.connect(self.restart)
        self.button_exit.clicked.connect(QtWidgets.QApplication.instance().quit)

    def on_clicked(self):
        cube = int(self.lineEdit_1.text())
        self.lineEdit_1.clear()
        throw = int(self.lineEdit_2.text())
        self.lineEdit_2.clear()
        sum = []
        for i in range(throw):
            count = 0
            for j in range(cube):
                num = randint(1, 6)
                count += num
            sum.append(count)
        for i in range(cube, (cube * 6) + 1):
            kol_sum = sum.count(i)
            ver = (kol_sum / throw) * 100
            self.textEdit.append(f"Вероятность выпадения {i}: {ver}%")

    def restart(self):
        self.label_cube.setText("Введите количество кубиков:")
        self.label_throw.setText("Введите количество бросков:")
        self.textEdit.clear()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindow()
    window.setWindowTitle("Угадай число!")
    window.resize(500, 500)
    window.show()
    sys.exit(app.exec())