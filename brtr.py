import sys
import time
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QLineEdit, QPushButton, QLabel, QTextEdit, QGridLayout, QMessageBox,
    QProgressBar, QSizePolicy, QFrame
)
from PyQt6.QtCore import QThread, QObject, pyqtSignal, Qt, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QIcon
import xml.etree.ElementTree as ET
import pandas as pd
import json

# --- Credential Loading ---
try:
    from cred import Cred
    DEFAULT_USERNAME = Cred.usname
    DEFAULT_PASSWORD = Cred.pws
except ImportError:
    DEFAULT_USERNAME = ""
    DEFAULT_PASSWORD = ""

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

TARGET_URL = "https://web.spaggiari.eu/home/app/default/login.php?custcode="
WAIT_TIMEOUT = 10
POPUP_CONFIRM_BUTTON_SELECTOR = "//button[text()='Conferma']"

class WorkerSignals(QObject):
    progress = pyqtSignal(str)
    step_progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

def xml_xls_to_json_filtered(xml_file_path, json_file_path):
    COLUMNS_TO_IGNORE = ['Nota Agenda','data_fine','tutto_il_giorno','data_inserimento','classe_desc','gruppo_desc','aula','materia']
    RENAME_MAP = {'data_inizio':'data'}
    NAMESPACES={'ss':'urn:schemas-microsoft-com:office:spreadsheet','':'urn:schemas-microsoft-com:office:spreadsheet'}
    tree=ET.parse(xml_file_path); root=tree.getroot(); data=[]
    worksheet=root.find('.//Worksheet',NAMESPACES)
    header_row=worksheet.find('.//Row',NAMESPACES)
    original_columns=[cell.find('ss:Data',NAMESPACES).text for cell in header_row.findall('ss:Cell',NAMESPACES) if cell.find('ss:Data',NAMESPACES) is not None]
    column_indices_to_include=[]; final_columns=[]
    for i,col in enumerate(original_columns):
        if col not in COLUMNS_TO_IGNORE:
            column_indices_to_include.append(i)
            final_columns.append(RENAME_MAP[col] if col in RENAME_MAP else col)
    for row in worksheet.findall('ss:Table/ss:Row',NAMESPACES)[1:]:
        row_full=[cell.find('ss:Data',NAMESPACES).text for cell in row.findall('ss:Cell',NAMESPACES)]
        while len(row_full)<len(original_columns): row_full.append(None)
        data.append([row_full[i] for i in column_indices_to_include])
    df=pd.DataFrame(data,columns=final_columns)
    with open(json_file_path,'w',encoding='utf-8') as f: f.write(df.to_json(orient='records',indent=4,force_ascii=False))

class AutomationWorker(QObject):
    def __init__(self, username, password, start_date, end_date):
        super().__init__()
        self.username = username
        self.password = password
        self.start_date = start_date
        self.end_date = end_date
        self.signals = WorkerSignals()

    def run(self):
        driver = None
        DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)

        # --- Clean old downloaded files ---
        for f in os.listdir(DOWNLOAD_DIR):
            try:
                os.remove(os.path.join(DOWNLOAD_DIR, f))
            except:
                pass
        try:
            self.signals.progress.emit("Starting Selenium web driver...")
            self.signals.step_progress.emit(10)

            chrome_options = Options()
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--ignore-ssl-errors')
            chrome_options.add_argument('--ignore-certificate-errors')
            prefs = {
                "download.default_directory": DOWNLOAD_DIR,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            chrome_options.add_experimental_option("prefs", prefs)

            driver = webdriver.Chrome(options=chrome_options)
            driver.get(TARGET_URL)
            wait = WebDriverWait(driver, WAIT_TIMEOUT)
            self.signals.step_progress.emit(20)

            self.signals.progress.emit("1. Performing login...")
            username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
            username_field.send_keys(self.username)

            password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
            password_field.send_keys(self.password)

            login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.accedi")))
            login_button.click()
            self.signals.step_progress.emit(40)

            self.signals.progress.emit("2. Navigating to Agenda...")
            agenda_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Agenda")))
            agenda_link.click()
            self.signals.step_progress.emit(55)

            self.signals.progress.emit("3. Configuring Export...")
            excel_export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@alt='Esportazione Excel']")))
            excel_export_button.click()
            self.signals.step_progress.emit(70)

            date_start_field = wait.until(EC.visibility_of_element_located((By.ID, "dal")))
            date_start_field.clear()
            date_start_field.send_keys(self.start_date)

            date_end_field = wait.until(EC.visibility_of_element_located((By.ID, "al")))
            date_end_field.clear()
            date_end_field.send_keys(self.end_date)
            self.signals.step_progress.emit(85)

            confirm_button = wait.until(EC.element_to_be_clickable((By.XPATH, POPUP_CONFIRM_BUTTON_SELECTOR)))
            confirm_button.click()
            self.signals.progress.emit("4. Download triggered.")
            self.signals.step_progress.emit(90)

            time.sleep(5)

            # --- Rename latest downloaded file to data.xls ---
            try:
                files = os.listdir(DOWNLOAD_DIR)
                if files:
                    latest = max([os.path.join(DOWNLOAD_DIR, f) for f in files], key=os.path.getctime)
                    new_path = os.path.join(DOWNLOAD_DIR, "data.xls")
                    os.replace(latest, new_path)
            except Exception as e:
                self.signals.progress.emit(f"Rename error: {e}")

            # --- Convert data.xls to output.json ---
            try:
                xml_xls_to_json_filtered(os.path.join(DOWNLOAD_DIR, "data.xls"), os.path.join(DOWNLOAD_DIR, "output.json"))
                self.signals.progress.emit("JSON conversion finished.")
            except Exception as e:
                self.signals.progress.emit(f"JSON error: {e}")(latest, new_path)
            except Exception as e:
                self.signals.progress.emit(f"Rename error: {e}")
            self.signals.step_progress.emit(100)
            self.signals.finished.emit("Automation complete! Check your downloads folder.")

        except Exception as e:
            self.signals.error.emit(f"Error: {type(e).__name__} - {str(e)}")
            self.signals.step_progress.emit(0)
        finally:
            if driver:
                driver.quit()

class AutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spaggiari Agenda Exporter")
        self.setWindowIcon(QIcon("icon.png"))
        self.setGeometry(200, 200, 900, 750)

        self.thread = None
        self.worker = None

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        self._apply_style()
        self._init_ui()

    def _apply_style(self):
        style = """
        QMainWindow { background-color: #1e1e1e; color: #f0f0f0; }
        QLabel { color: #f0f0f0; font-size: 12pt; }
        QLineEdit {
            padding: 10px;
            border: 1px solid #333;
            border-radius: 8px;
            background: #2b2b2b;
            color: white;
        }
        QPushButton {
            padding: 12px;
            font-size: 14pt;
            background-color: #3a7bd5;
            color: white;
            border-radius: 8px;
        }
        QPushButton:hover { background-color: #57a0ff; }
        QTextEdit {
            background: #2b2b2b;
            padding: 10px;
            border-radius: 8px;
            color: white;
            font-family: Consolas;
        }
        QProgressBar {
            height: 25px;
            background: #2b2b2b;
            border-radius: 8px;
        }
        QProgressBar::chunk {
            border-radius: 8px;
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #4facfe, stop:1 #00f2fe);
        }
        """
        self.setStyleSheet(style)

    def _init_ui(self):
        input_group = QWidget()
        input_layout = QGridLayout(input_group)

        self.username_input = QLineEdit(text=DEFAULT_USERNAME)
        self.password_input = QLineEdit(text=DEFAULT_PASSWORD)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.start_date_input = QLineEdit(text="01-11-2025")
        self.end_date_input = QLineEdit(text="01-10-2026")

        input_layout.addWidget(QLabel("Username:"), 0, 0)
        input_layout.addWidget(self.username_input, 0, 1)
        input_layout.addWidget(QLabel("Password:"), 1, 0)
        input_layout.addWidget(self.password_input, 1, 1)

        input_layout.addWidget(QLabel("Start Date:"), 2, 0)
        input_layout.addWidget(self.start_date_input, 2, 1)
        input_layout.addWidget(QLabel("End Date:"), 3, 0)
        input_layout.addWidget(self.end_date_input, 3, 1)

        self.main_layout.addWidget(input_group)

        self.start_button = QPushButton("Start Automation")
        self.start_button.clicked.connect(self.start_automation)
        self.main_layout.addWidget(self.start_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.main_layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.main_layout.addWidget(self.log_output)

    def start_automation(self):
        username = self.username_input.text().strip()
        password = self.password_input.text()
        start_date = self.start_date_input.text().strip()
        end_date = self.end_date_input.text().strip()

        if not all([username, password, start_date, end_date]):
            QMessageBox.warning(self, "Missing Input", "Fill all fields.")
            return

        self.start_button.setEnabled(False)
        self.start_button.setText("Running...")
        self.log_output.clear()

        self.thread = QThread()
        self.worker = AutomationWorker(username, password, start_date, end_date)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.signals.progress.connect(self.update_log)
        self.worker.signals.step_progress.connect(self.update_progress_bar)
        self.worker.signals.finished.connect(self.automation_finished)
        self.worker.signals.error.connect(self.automation_error)

        self.thread.start()

    def update_log(self, message):
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def update_progress_bar(self, value):
        if not hasattr(self, "_progress_anim"):
            self._progress_anim = QPropertyAnimation(self.progress_bar, b"value")
            self._progress_anim.setDuration(350)
            self._progress_anim.setEasingCurve(QEasingCurve.Type.InOutCubic)

        self._progress_anim.stop()
        self._progress_anim.setStartValue(self.progress_bar.value())
        self._progress_anim.setEndValue(value)
        self._progress_anim.start()

    def automation_finished(self, result):
        self.log_output.append("<span style='color:#4aff4a;'>" + result + "</span>")
        self._reset_ui()
        QMessageBox.information(self, "Done", "Export completed!")

    def automation_error(self, error):
        self.log_output.append("<span style='color:red;'>" + error + "</span>")
        self._reset_ui()
        QMessageBox.critical(self, "Error", error)

    def _reset_ui(self):
        self.start_button.setEnabled(True)
        self.start_button.setText("Start Automation")
        if self.thread:
            self.thread.quit()
            self.thread.wait()
            self.thread = None
            self.worker = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutomationApp()
    window.show()
    sys.exit(app.exec())
