
import os
import sys

PRODUCT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(PRODUCT_DIR)
sys.path.append(r"C:\FalconX")
sys.path.append(PRODUCT_DIR)
sys.path.append(os.path.join(PRODUCT_DIR, "scripts"))
sys.path.append(os.path.join(PRODUCT_DIR, "scripts", "mfg-site-packages"))
sys.path.append(os.path.join(PRODUCT_DIR, "scripts", "sctl-sol-yeti-scripts"))
sys.path.append(os.path.join(PRODUCT_DIR, "scripts", "sctl-falcon"))


from yeti_service_window import MainWindow
from PySide6 import QtWidgets
from qt_material import apply_stylesheet



import conf
import version

def main(argv):


    myapp = QtWidgets.QApplication(argv)
    mainWin = MainWindow(conf.repair_rfu_dict, conf.audit_boid_dict, conf.product_names_dict, conf.fw_info_dict, conf.has_rtc_support_dict, conf.USB_PORTS, version=version.__version__)
    apply_stylesheet(myapp, theme='dark_blue.xml')
    mainWin.show()
    myapp.exec()


if __name__  == '__main__':
    main(sys.argv)
