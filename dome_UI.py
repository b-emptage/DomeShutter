from PySide6.QtWidgets import QApplication, QMainWindow
from ui_dome import Ui_MainWindow
from dome_shutter import Dome_Control
import sys


class DomeWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.dome = Dome_Control()

        # --- Connect WEST buttons ---
        self.ui.west_open.clicked.connect(self.west_open)
        self.ui.west_stop.clicked.connect(self.west_stop)
        self.ui.west_close.clicked.connect(self.west_close)
        self.ui.west_set.clicked.connect(self.west_setpoint)

        # --- Connect EAST buttons ---
        self.ui.east_open.clicked.connect(self.east_open)
        self.ui.east_stop.clicked.connect(self.east_stop)
        self.ui.east_close.clicked.connect(self.east_close)
        self.ui.east_set.clicked.connect(self.east_setpoint)

        # --- Sliders ---
        self.ui.west_set_pos.valueChanged.connect(self.west_slider_changed)
        self.ui.east_set_pos.valueChanged.connect(self.east_slider_changed)

        # --- Menu actions ---
        self.ui.actionConnect_RainMon.triggered.connect(self.connect_rainmon)
        self.ui.actionDisable_RDP_Monitor.triggered.connect(self.toggle_rdp)

    # ===== WEST FUNCTIONS =====
    def west_open(self):
        self.dome.open_w()

    def west_stop(self):
        self.dome.stop_w()

    def west_close(self):
        self.dome.close_w()

    def west_setpoint(self):
        value = self.ui.west_set_pos.value()
        print(f"WEST SETPOINT → {value}")
        self.ui.west_progress.setValue(value)

    def west_slider_changed(self, value):
        print(f"WEST SLIDER → {value}")

    # ===== EAST FUNCTIONS =====
    def east_open(self):
        self.dome.open_e()

    def east_stop(self):
        self.dome.stop_e()

    def east_close(self):
        self.dome.close_e()

    def east_setpoint(self):
        value = self.ui.east_set_pos.value()
        print(f"EAST SETPOINT → {value}")
        self.ui.east_progress.setValue(value)

    def east_slider_changed(self, value):
        print(f"EAST SLIDER → {value}")

    # ===== MENU =====
    def connect_rainmon(self):
        print("Connecting Rain Monitor...")

    def toggle_rdp(self):
        print("Toggling RDP Monitor")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = DomeWindow()
    window.show()

    sys.exit(app.exec())