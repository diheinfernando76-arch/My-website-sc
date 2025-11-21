import sys
import time
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QLineEdit, QPushButton, QLabel, QTextEdit, QGridLayout, QMessageBox,
    QProgressBar, QSizePolicy
)
from PyQt6.QtCore import QThread, QObject, pyqtSignal, Qt
from PyQt6.QtWidgets import QFrame
from PyQt6.QtCore import QPropertyAnimation, QEasingCurve



# --- Credential Loading ---
# Attempt to import credentials from cred.py
try:
    from cred import Cred
    DEFAULT_USERNAME = Cred.usname
    DEFAULT_PASSWORD = Cred.pws
except ImportError:
    DEFAULT_USERNAME = ""
    DEFAULT_PASSWORD = ""
    print("Warning: 'cred.py' not found. Credentials must be entered manually.")

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 

# --- CONFIGURATION (Selectors and URLs) ---
TARGET_URL = "https://web.spaggiari.eu/home/app/default/login.php?custcode="
WAIT_TIMEOUT = 10 
POPUP_CONFIRM_BUTTON_SELECTOR = "//button[text()='Conferma']" 

class WorkerSignals(QObject):
    """Defines the signals available from a running worker thread."""
    progress = pyqtSignal(str)      # For emitting status updates to the GUI log
    step_progress = pyqtSignal(int) # For updating the QProgressBar (0-100)
    finished = pyqtSignal(str)      # Emitted when the worker completes successfully
    error = pyqtSignal(str)         # Emitted when an error occurs

class AutomationWorker(QObject):
    """
    Worker class to handle the heavy Selenium automation in a separate thread.
    """
    def __init__(self, username, password, start_date, end_date):
        super().__init__()
        self.username = username
        self.password = password
        self.start_date = start_date
        self.end_date = end_date
        self.signals = WorkerSignals()

    def run(self):
        """The main execution function for the automation."""
        driver = None
        DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads") 
        os.makedirs(DOWNLOAD_DIR, exist_ok=True) 
        
        try:
            self.signals.progress.emit("Starting Selenium web driver...")
            self.signals.step_progress.emit(10) # 10%
            self.signals.progress.emit(f"Downloads will be saved to: {DOWNLOAD_DIR}")

            # Set up Chrome options for robustness and custom downloads
            chrome_options = Options()
            # *** ENABLE HEADLESS MODE ***
            chrome_options.add_argument('--headless')
            # ----------------------------
            chrome_options.add_argument('--ignore-ssl-errors')
            chrome_options.add_argument('--ignore-certificate-errors')
            
            # Custom Download Preferences
            prefs = {
                "download.default_directory": DOWNLOAD_DIR,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True 
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # Initialize Driver
            driver = webdriver.Chrome(options=chrome_options)
            self.signals.progress.emit(f"Opening URL: {TARGET_URL}")
            driver.get(TARGET_URL)
            wait = WebDriverWait(driver, WAIT_TIMEOUT)
            self.signals.step_progress.emit(20) # 20%

            # --- STEP 1 & 2: Login ---
            self.signals.progress.emit("1. Performing login sequence...")
            
            # Fill Username
            username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
            username_field.send_keys(self.username)
            self.signals.progress.emit("   -> Username/Codice Personale entered.")
            
            # Fill Password
            password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
            password_field.send_keys(self.password)
            self.signals.progress.emit("   -> Password entered.")

            # --- STEP 3: Click Login Button ---
            login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.accedi")))
            login_button.click()
            self.signals.progress.emit("   -> Clicked Login. Waiting for dashboard...")
            self.signals.step_progress.emit(40) # 40%

            # --- STEP 4: Click 'Agenda' Link ---
            self.signals.progress.emit("2. Navigating to Agenda...")
            agenda_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Agenda")))
            agenda_link.click()
            self.signals.progress.emit("   -> Clicked Agenda link.")
            self.signals.step_progress.emit(55) # 55%
            
            # --- STEP 5: Trigger Export & Interact with Configuration Area ---
            self.signals.progress.emit("3. Configuring Excel Export...")
            
            # Click the Excel Export Button
            excel_export_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//img[@alt='Esportazione Excel']"))
            )
            excel_export_button.click()
            self.signals.progress.emit("   -> Export button clicked. Waiting for modal...")
            self.signals.step_progress.emit(70) # 70%

            # 6.1 Set the start date ('dal') input field 
            date_start_field = wait.until(EC.visibility_of_element_located((By.ID, "dal")))
            date_start_field.clear() 
            date_start_field.send_keys(self.start_date)
            self.signals.progress.emit(f"   -> Set start date ('dal') to: {self.start_date}")

            # 6.2 Set the end date ('al') input field 
            date_end_field = wait.until(EC.visibility_of_element_located((By.ID, "al")))
            date_end_field.clear() 
            date_end_field.send_keys(self.end_date)
            self.signals.progress.emit(f"   -> Set end date ('al') to: {self.end_date}")
            self.signals.step_progress.emit(85) # 85%

            # 6.3 Click the confirmation button ('Conferma')
            self.signals.progress.emit("   -> Clicking final confirmation button...")
            confirm_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, POPUP_CONFIRM_BUTTON_SELECTOR))
            )
            confirm_button.click()
            self.signals.progress.emit("4. Download Triggered.")
            self.signals.step_progress.emit(90) # 90%
            
            # Final observation pause
            time.sleep(5)
            self.signals.progress.emit("5. Download pause complete.")
            
            self.signals.step_progress.emit(100) # 100%
            self.signals.finished.emit("Automation complete! Check your 'downloads' folder.")

        except Exception as e:
            error_msg = f"Automation failed at step. Error: {type(e).__name__} - {str(e)}"
            self.signals.error.emit(error_msg)
            self.signals.step_progress.emit(0) # Reset progress on error
        finally:
            if driver:
                self.signals.progress.emit("Closing browser...")
                driver.quit()
            
class AutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spaggiari Agenda Exporter")
        self.setGeometry(200, 100, 900, 650)

        self.thread = None
        self.worker = None

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(20)

        self._apply_style()
        self._init_ui()

    def _apply_style(self):
        """Modern glassy dark style."""
        style = """
        QMainWindow {
            background-color: #1e1f22;
        }

        QLabel {
            color: #f0f0f0;
            font-size: 14px;
        }

        .titleLabel {
            font-size: 26px;
            font-weight: bold;
            color: #61dafb;
            margin-bottom: 10px;
        }

        QFrame#card {
            background-color: #2a2d31;
            border-radius: 16px;
            padding: 20px;
            border: 1px solid #3a3d42;
        }

        QLineEdit {
            padding: 10px;
            border-radius: 10px;
            background-color: #3a3d42;
            border: 1px solid #4a4d52;
            color: white;
            font-size: 13px;
        }
        QLineEdit:focus {
            border: 1px solid #61dafb;
        }

        QPushButton {
            background-color: #61dafb;
            color: #000;
            padding: 14px;
            border-radius: 12px;
            font-size: 16px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #52c7e8;
        }
        QPushButton:disabled {
            background-color: #2e3a40;
            color: #777;
        }

        QTextEdit {
            background-color: #232527;
            border-radius: 12px;
            padding: 12px;
            color: #f0f0f0;
            font-family: Consolas;
            font-size: 13px;
        }

        QProgressBar {
            height: 22px;
            border-radius: 10px;
            background-color: #2e2f31;
            text-align: center;
            color: white;
        }
        QProgressBar::chunk {
            border-radius: 10px;
            background-color: #61dafb;
        }
        """
        self.setStyleSheet(style)

    def _init_ui(self):
        # ---------- HEADER ----------
        header = QLabel("Spaggiari Agenda Exporter")
        header.setObjectName("titleLabel")
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.main_layout.addWidget(header)

        # ---------- INPUT CARD ----------
        card = QFrame()
        card.setObjectName("card")
        card_layout = QGridLayout(card)
        card_layout.setSpacing(18)

        self.username_input = QLineEdit(text=DEFAULT_USERNAME)
        self.password_input = QLineEdit(text=DEFAULT_PASSWORD)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)

        self.start_date_input = QLineEdit("01-11-2025")
        self.end_date_input = QLineEdit("01-10-2026")

        labels = ["Username", "Password", "Start Date (DD-MM-YYYY)", "End Date (DD-MM-YYYY)"]
        widgets = [self.username_input, self.password_input, self.start_date_input, self.end_date_input]

        for i, (lbl, widget) in enumerate(zip(labels, widgets)):
            card_layout.addWidget(QLabel(lbl), i, 0)
            card_layout.addWidget(widget, i, 1)

        self.main_layout.addWidget(card)

        # ---------- BUTTON ----------
        self.start_button = QPushButton("Start Automation")
        self.start_button.clicked.connect(self.start_automation)
        self.start_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.main_layout.addWidget(self.start_button)

        # ---------- PROGRESS ----------
        self.progress_bar = QProgressBar()
        self.main_layout.addWidget(self.progress_bar)

        # ---------- LOG ----------
        log_label = QLabel("Execution Log")
        log_label.setStyleSheet("font-size: 20px; margin-top: 10px;")
        self.main_layout.addWidget(log_label)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.main_layout.addWidget(self.log_output)

        self.main_layout.setStretch(3, 1)

    def start_automation(self):
        """Initializes and starts the worker thread."""
        username = self.username_input.text().strip()
        password = self.password_input.text()
        start_date = self.start_date_input.text().strip()
        end_date = self.end_date_input.text().strip()

        if not all([username, password, start_date, end_date]):
            QMessageBox.warning(self, "Missing Input", "Please fill in all fields (Username, Password, and Dates).")
            return

        # Disable button and clear log
        self.start_button.setEnabled(False)
        self.start_button.setText("Automation Running...")
        self.progress_bar.setValue(0)
        self.log_output.clear()
        self.log_output.append(f"Starting automation for date range: {start_date} to {end_date}...")

        # 1. Create a QThread object
        self.thread = QThread()
        # 2. Create a worker object
        self.worker = AutomationWorker(username, password, start_date, end_date)
        # 3. Move worker to the thread
        self.worker.moveToThread(self.thread)

        # 4. Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.signals.progress.connect(self.update_log)
        self.worker.signals.step_progress.connect(self.update_progress_bar)
        self.worker.signals.finished.connect(self.automation_finished)
        self.worker.signals.error.connect(self.automation_error)

        # 5. Start the thread
        self.thread.start()

    def update_log(self, message):
        """Slot to receive progress messages from the worker."""
        self.log_output.append(message)
        # Scroll to the bottom automatically
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())

    def update_progress_bar(self, value):
        """Smooth animated progress bar updates."""
        if not hasattr(self, "_progress_anim"):
            self._progress_anim = QPropertyAnimation(self.progress_bar, b"value")
            self._progress_anim.setDuration(300)  # 0.3s animation
            self._progress_anim.setEasingCurve(QEasingCurve.Type.InOutCubic)

        self._progress_anim.stop()
        self._progress_anim.setStartValue(self.progress_bar.value())
        self._progress_anim.setEndValue(value)
        self._progress_anim.start()




    def automation_finished(self, result):
        """Slot to handle successful completion."""
        self.log_output.append(f"\n<span style='color: #2ecc71;'>✅ SUCCESS: {result}</span>")
        self.progress_bar.setValue(100)
        self._reset_ui()
        QMessageBox.information(self, "Success", "Export completed successfully!")

    def automation_error(self, error_msg):
        """Slot to handle errors."""
        self.log_output.append(f"\n<span style='color: #e74c3c;'>❌ ERROR: {error_msg}</span>")
        self.progress_bar.setValue(0)
        self._reset_ui()
        QMessageBox.critical(self, "Error", "Automation failed. Check the log for details.")

    def _reset_ui(self):
        """Resets the UI elements after completion or error."""
        self.start_button.setEnabled(True)
        self.start_button.setText("Start Automation")
        # Cleanup thread
        if self.thread:
            self.thread.quit()
            self.thread.wait()
            self.thread = None
            self.worker = None

    def closeEvent(self, event):
        """Ensures the thread is properly terminated when the app is closed."""
        if self.thread and self.thread.isRunning():
            self.thread.terminate()
            self.thread.wait()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutomationApp()
    window.show()
    sys.exit(app.exec())