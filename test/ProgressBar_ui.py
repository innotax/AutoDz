"""
Data-Carrying PySide/PyQt Signals
"""
from PyQt5.QtCore import pyqtSlot, pyqtSignal, QObject


class Circle(QObject):
    # Signal emitted when the circle is resized,
    # carrying its integer radius
    resized = pyqtSignal(int)  # 반지름 크기 변할때, 반지름 보내는 signal
    # Signal emitted when the circle is moved, carrying
    # the x and y coordinates of its center.
    moved = pyqtSignal(int, int)  # 원 이동시, 중심자표 보내는 signal

    def __init__(self, x, y, r):
        # Initialize the Circle as a QObject so it can emit signals
        super().__init__()

        # "Hide" the values and expose them via properties
        self.__x = x
        self.__y = y
        self.__r = r

    @property
    def x(self):
        return self.__x

    @x.setter
    def x(self, new_x):
        self.__x = new_x
        # After the center is moved, emit the
        # moved signal with the new coordinates
        self.moved.emit(new_x, self.y)

    @property
    def y(self):
        return self.__y

    @y.setter
    def y(self, new_y):
        self.__y = new_y
        # After the center is moved, emit the moved
        # signal with the new coordinates
        self.moved.emit(self.x, new_y)

    @property
    def r(self):
        return self.__r

    @r.setter
    def r(self, new_r):
        self.__r = new_r
        # After the radius is changed, emit the
        # resized signal with the new radius
        self.resized.emit(new_r)


# A slot for the "moved" signal, accepting the x and y coordinates
@pyqtSlot(int, int)
def on_moved(x, y):
    print('Circle was moved to (%s, %s).' % (x, y))


# A slot for the "resized" signal, accepting the radius
@pyqtSlot(int)
def on_resized(r):
    print('Circle was resized to radius %s.' % r)


if __name__ == "__main__":
    c = Circle(5, 5, 4)

    # Connect the Circle's signals to our simple slots
    c.moved.connect(on_moved)
    c.resized.connect(on_resized)

    # Move the circle one unit to the right
    c.x += 1

    # Increase the circle's radius by one unit
    c.r += 1
