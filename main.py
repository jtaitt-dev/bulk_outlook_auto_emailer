import sys
import os
import re
import random
import numpy as np
import pandas as pd
import win32com.client
import pythoncom
import logging
import urllib.parse
from datetime import datetime
from PyQt6.QtCore import Qt, QTimer, QRectF, QPointF
from PyQt6.QtGui import QPainter, QPainterPath, QLinearGradient, QColor, QPen, QFont, QBrush, QRadialGradient
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QLineEdit, QTextEdit, QFileDialog, QMessageBox, QCheckBox, QDialog
)

APP_NAME = "MailOps"
APP_TAGLINE = "Precision bulk email. Zero surprises."

class WaveSplashScreen(QWidget):
    def __init__(self, width: int = 854, height: int = 480):
        super().__init__()
        self.width, self.height = width, height
        self.t, self.particles = 0.0, []
        self.setFixedSize(self.width, self.height)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.title_font = QFont("sans-serif", 52, QFont.Weight.Bold)
        self.subtitle_font = QFont("sans-serif", 18, QFont.Weight.Light)
        self.credit_font = QFont("sans-serif", 10, QFont.Weight.Normal)
        self.timer = QTimer(self)
        self.timer.setInterval(16)
        self.timer.timeout.connect(self._update_animation)

    def showEvent(self, event):
        if self.screen():
            screen_geometry = self.screen().geometry()
            self.move((screen_geometry.width() - self.width) // 2, (screen_geometry.height() - self.height) // 2)
        self.timer.start()
        super().showEvent(event)

    def hideEvent(self, event):
        self.timer.stop()
        super().hideEvent(event)

    def _wave(self, x: float, t: float) -> float:
        y = np.sin((x * 0.005) + t * 1.0) * 30
        y += np.sin((x * 0.01) + t * 2.5) * 40
        y += np.sin((x * 0.02) + t * 0.5) * 20
        return y + self.height * 0.45

    def _update_animation(self):
        self.t += 0.01
        if len(self.particles) < 300:
            for _ in range(5):
                x_pos = random.uniform(0, self.width)
                self.particles.append(
                    {
                        'x': x_pos,
                        'y': self._wave(x_pos, self.t) + random.uniform(-20, 20),
                        'vx': random.uniform(-0.5, 0.5),
                        'vy': random.uniform(-0.2, 0.2),
                        'alpha': random.uniform(50, 150),
                        'max_alpha': random.uniform(100, 200),
                        'size': random.uniform(1.5, 4.5),
                        'life': 1.0,
                    }
                )
        self.particles = [p for p in self.particles if p['life'] > 0]
        for p in self.particles:
            p['x'] += p['vx']
            p['y'] += p['vy']
            p['life'] -= 0.005
            p['alpha'] = p['max_alpha'] * np.sin(p['life'] * np.pi)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        gradient = QLinearGradient(0, 0, 0, self.height)
        gradient.setColorAt(0, QColor(60, 80, 120))
        gradient.setColorAt(1, QColor(35, 45, 80))
        painter.fillRect(self.rect(), gradient)
        path = QPainterPath()
        path.moveTo(-10, self._wave(-10, self.t))
        for x in range(0, self.width + 11, 10):
            path.lineTo(x, self._wave(x, self.t))
        for width, color in [(25, QColor(0, 255, 255, 6)), (15, QColor(100, 255, 255, 12)), (5, QColor(200, 255, 255, 25))]:
            pen = QPen(color, width, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap)
            painter.setPen(pen)
            painter.drawPath(path)
        painter.setPen(Qt.PenStyle.NoPen)
        for p in self.particles:
            grad = QRadialGradient(QPointF(p['x'], p['y']), p['size'])
            grad.setColorAt(0, QColor(255, 255, 255, int(p['alpha'])))
            grad.setColorAt(1, QColor(255, 255, 255, 0))
            painter.setBrush(QBrush(grad))
            painter.drawEllipse(QPointF(p['x'], p['y']), p['size'], p['size'])
        text_opacity = max(0.0, min(1.0, (self.t - 0.5) / 1.5))
        if text_opacity > 0:
            painter.setPen(QColor(255, 255, 255, int(255 * text_opacity)))
            painter.setFont(self.title_font)
            painter.drawText(QRectF(0, self.height * 0.6, self.width, 100), Qt.AlignmentFlag.AlignCenter, APP_NAME)
            painter.setFont(self.subtitle_font)
            painter.drawText(QRectF(0, self.height * 0.75, self.width, 50), Qt.AlignmentFlag.AlignCenter, APP_TAGLINE)
            painter.setFont(self.credit_font)
            painter.drawText(QRectF(0, self.height - 30, self.width, 20), Qt.AlignmentFlag.AlignCenter, "Built by Joshua Taitt – Neta Scientific")

class EmailSender(QWidget):
    def __init__(self):
        super().__init__()
        self.outlook_app = None
        self.log_file_path = ""
        self.excel_path = None
        self._ui_ready = False
        self._log_buffer = []
        self.setup_logging()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(f"{APP_NAME} — Outlook Bulk Sender (Professional Edition)")
        self.setGeometry(300, 300, 700, 620)
        layout = QVBoxLayout()
        self.help_btn = QPushButton("📘 How to Use")
        self.subject_input = QLineEdit()
        self.cc_input = QLineEdit()
        self.body_input = QTextEdit()
        self.file_btn = QPushButton("Select Excel File")
        self.file_label = QLabel("No Excel file selected.")
        self.preview_checkbox = QCheckBox("Preview each email before sending (Recommended for testing)")
        self.send_btn = QPushButton("Start Sending Emails")
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.help_btn)
        layout.addWidget(QLabel("Subject:"))
        layout.addWidget(self.subject_input)
        layout.addWidget(QLabel("CC Email(s) (optional, separate with a semicolon ';'):"))
        layout.addWidget(self.cc_input)
        layout.addWidget(QLabel("Email Body (Copy from Word and Paste Here):"))
        layout.addWidget(self.body_input)
        layout.addWidget(self.file_btn)
        layout.addWidget(self.file_label)
        layout.addWidget(self.preview_checkbox)
        layout.addWidget(self.send_btn)
        layout.addWidget(QLabel("Log Output:"))
        layout.addWidget(self.log_output)
        self.setLayout(layout)
        self.file_btn.clicked.connect(self.select_excel_file)
        self.send_btn.clicked.connect(self.send_emails)
        self.help_btn.clicked.connect(self.show_help_dialog)
        self._ui_ready = True
        if self._log_buffer:
            for ts, msg in self._log_buffer:
                self.log_output.append(f"{ts} - {msg}")
            self._log_buffer.clear()
        QApplication.processEvents()

    def setup_logging(self):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.log_file_path = f"mailops_log_{timestamp}.log"
        logging.basicConfig(level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(lineno)d - %(message)s',
                            filename=self.log_file_path, filemode='w')
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(formatter)
        logging.getLogger().addHandler(console_handler)
        self.log_message("Logging initialized.", logging.DEBUG)
        self.log_message(f"{APP_NAME} starting.", logging.INFO)

    def log_message(self, message, level=logging.INFO):
        log_levels = {logging.DEBUG: logging.debug, logging.INFO: logging.info, logging.WARNING: logging.warning, logging.ERROR: logging.error}
        log_function = log_levels.get(level, logging.info)
        log_function(message)
        ts = datetime.now().strftime('%H:%M:%S')
        if level >= logging.INFO:
            if self._ui_ready:
                self.log_output.append(f"{ts} - {message}")
            else:
                self._log_buffer.append((ts, message))
        QApplication.processEvents()

    def get_signature_from_file(self):
        self.log_message("Reading Outlook signature.", logging.DEBUG)
        sig_dir = os.path.join(os.getenv("APPDATA"), "Microsoft", "Signatures")
        if not os.path.isdir(sig_dir):
            self.log_message(f"Signature directory not found: {sig_dir}", logging.ERROR)
            QMessageBox.critical(self, "Signature Error", f"Signature directory not found:\n{sig_dir}")
            return None, None
        htm_files = [f for f in os.listdir(sig_dir) if f.endswith(".htm")]
        if not htm_files:
            self.log_message("No .htm signature files found in directory.", logging.WARNING)
            QMessageBox.warning(self, "Signature Error", "No HTML signature files found.")
            return "", ""
        latest_file = max(htm_files, key=lambda f: os.path.getmtime(os.path.join(sig_dir, f)))
        sig_path = os.path.join(sig_dir, latest_file)
        self.log_message(f"Using signature file: {sig_path}", logging.INFO)
        sig_base_name = os.path.splitext(latest_file)[0]
        sig_files_dirname = f"{sig_base_name}_files"
        sig_files_dirpath = os.path.join(sig_dir, sig_files_dirname)
        html_content = ""
        try:
            with open(sig_path, "r", encoding="utf-8") as f:
                html_content = f.read()
        except Exception:
            self.log_message(f"Signature file {sig_path} not UTF-8, trying cp1252.", logging.DEBUG)
            with open(sig_path, "r", encoding="cp1252") as f:
                html_content = f.read()
        return html_content, sig_files_dirpath

    def embed_images_and_update_html(self, mail_item, signature_html, sig_files_dirpath):
        self.log_message("Embedding signature images.", logging.DEBUG)
        if not signature_html or not sig_files_dirpath or not os.path.isdir(sig_files_dirpath):
            self.log_message("Signature HTML empty or assets missing, using original HTML.", logging.WARNING)
            return signature_html
        updated_html = signature_html
        html_relative_folder_name = os.path.basename(sig_files_dirpath)
        referenced_files = {f.lower() for f in re.findall(r'src=["\'](?:[^"\']*/)?([^"\']+)["\']', signature_html, re.IGNORECASE)}
        for image_filename in os.listdir(sig_files_dirpath):
            try:
                if image_filename.lower() not in referenced_files:
                    continue
                image_path = os.path.join(sig_files_dirpath, image_filename)
                cid = image_filename
                attachment = mail_item.Attachments.Add(image_path)
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
                path_raw = f"{html_relative_folder_name}/{image_filename}"
                path_encoded = f"{html_relative_folder_name}/{urllib.parse.quote(image_filename)}"
                cid_link = f"cid:{cid}"
                updated_html = (
                    updated_html
                    .replace(f'src="{path_raw}"', f'src="{cid_link}"')
                    .replace(f"src='{path_raw}'", f"src='{cid_link}'")
                    .replace(f'src="{path_encoded}"', f'src="{cid_link}"')
                    .replace(f"src='{path_encoded}'", f"src='{cid_link}'")
                )
            except Exception as e:
                self.log_message(f"Failed to embed image '{image_filename}': {e}", logging.ERROR)
        return updated_html

    def try_attach_file(self, mail_item, attachment_path, email_address):
        if pd.isna(attachment_path) or not isinstance(attachment_path, str):
            return "No attachment specified."
        path_list = attachment_path.split(';')
        attachment_notes = []
        for path in path_list:
            clean_path = path.strip().strip('"').strip("'")
            if not clean_path:
                continue
            if not os.path.isfile(clean_path):
                self.log_message(f"Attachment file not found: {clean_path}", logging.WARNING)
                attachment_notes.append(f"NOT FOUND: {os.path.basename(clean_path)}")
            else:
                try:
                    mail_item.Attachments.Add(clean_path)
                    self.log_message(f"Attached: {clean_path}", logging.INFO)
                    attachment_notes.append(f"Attached: {os.path.basename(clean_path)}")
                except Exception as e:
                    self.log_message(f"Error attaching '{clean_path}' for '{email_address}': {e}", logging.ERROR)
                    attachment_notes.append(f"ERROR attaching: {os.path.basename(clean_path)}")
        if not attachment_notes:
            return "No valid attachments specified."
        return "\n".join(attachment_notes)

    def select_excel_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.excel_path = file
            self.file_label.setText(f"Selected: {os.path.basename(file)}")
            self.log_message(f"Excel file selected: {self.excel_path}", logging.INFO)

    def send_emails(self):
        self.log_message("Send_emails process started.", logging.INFO)
        if not self.excel_path:
            QMessageBox.critical(self, "Error", "An Excel file has not been selected.")
            return
        base_subject = self.subject_input.text().strip()
        body_html_from_ui = self.body_input.toHtml()
        if not base_subject or not self.body_input.toPlainText().strip():
            QMessageBox.warning(self, "Missing Information", "The 'Subject' and 'Email Body' fields are required.")
            return
        try:
            self.log_message("Initializing Outlook application.", logging.DEBUG)
            self.outlook_app = win32com.client.Dispatch("Outlook.Application")
            sig_html_original, sig_files_dirpath = self.get_signature_from_file()
            if sig_html_original is None:
                return
        except Exception as e:
            self.log_message(f"CRITICAL: Could not connect to Outlook or get signature. Error: {e}", logging.ERROR)
            QMessageBox.critical(self, "Outlook Error", f"Could not connect to Outlook.\nError: {e}")
            return
        try:
            df = pd.read_excel(self.excel_path)
            if 'Email' not in df.columns:
                raise ValueError("Excel file must contain an 'Email' column.")
            for col in ['Attachment', 'Greeting', 'Supplier Name']:
                if col not in df.columns:
                    df[col] = ''
            df.fillna('', inplace=True)
        except Exception as e:
            self.log_message(f"ERROR reading Excel file: {e}", logging.ERROR)
            QMessageBox.critical(self, "Excel Error", f"Failed to read the Excel file.\nError: {e}")
            return
        cc = self.cc_input.text().strip()
        preview_mode = self.preview_checkbox.isChecked()
        sent_count = 0
        total_count = len(df)
        self.log_message(f"Starting email batch for {total_count} records. Preview mode: {preview_mode}", logging.INFO)
        for index, row in df.iterrows():
            email_address = str(row['Email']).strip()
            if not email_address:
                self.log_message(f"Skipping row {index + 2}: Email address is missing.", logging.WARNING)
                continue
            self.log_message(f"Processing row {index + 2} for recipient: {email_address}", logging.DEBUG)
            try:
                mail = self.outlook_app.CreateItem(0)
                mail.To = email_address
                mail.CC = cc
                supplier_name = str(row.get('Supplier Name', '')).strip()
                final_subject = f"{supplier_name} - {base_subject}" if supplier_name else base_subject
                mail.Subject = final_subject
                updated_sig_html = self.embed_images_and_update_html(mail, sig_html_original, sig_files_dirpath)
                greeting = str(row['Greeting']).strip()
                greeting_html_part = f"<p>{greeting.rstrip(',')}," + "</p>" if greeting else ""
                styled_body_content = f"""
                <div style="font-family:Calibri, sans-serif; font-size:11pt;">
                    {greeting_html_part}
                    {body_html_from_ui}
                </div>
                """
                final_html_body = styled_body_content + updated_sig_html
                mail.HTMLBody = final_html_body
                self.log_message(f"Final HTMLBody set for {email_address}.", logging.DEBUG)
                attachment_path = row.get('Attachment', '')
                attachment_note = self.try_attach_file(mail, attachment_path, email_address)
                if preview_mode:
                    mail.Display()
                    reply = QMessageBox.question(
                        self,
                        "Confirm Send",
                        f"Review the email for {email_address}.\n\n{attachment_note}\n\nDo you want to send it?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        mail.Send()
                        self.log_message(f"SENT (after preview) to {email_address}", logging.INFO)
                        sent_count += 1
                    else:
                        self.log_message(f"SKIPPED (after preview) sending to {email_address} by user.", logging.INFO)
                else:
                    mail.Send()
                    self.log_message(f"SENT to {email_address}", logging.INFO)
                    sent_count += 1
            except pythoncom.com_error as e:
                self.log_message(f"A COM Error occurred for {email_address}: {e}", logging.ERROR)
            except Exception as e:
                self.log_message(f"An unexpected error occurred for {email_address}: {e}", logging.ERROR)
        summary_message = f"Email batch finished. Sent {sent_count} of {total_count} emails."
        self.log_message(summary_message, logging.INFO)
        QMessageBox.information(self, "Process Complete", f"{summary_message}\n\nA detailed log file has been saved to:\n{self.log_file_path}")
        self.outlook_app = None

    def show_help_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"How to Use {APP_NAME}")
        dialog.setMinimumSize(600, 500)
        layout = QVBoxLayout(dialog)
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml(f"""
        <h2 style="color:#003DA5;">How to Use {APP_NAME}</h2><hr>
        <ol>
            <li><b>Compose Your Email:</b>
                <ul>
                    <li>Fill in the <b>Subject</b>. This will be the base subject for all emails.</li>
                    <li>For the <b>Email Body</b>, compose your message in Microsoft Word or another editor, <b>copy it</b>, and then <b>paste it directly</b> into the text box. The font will be automatically set to Calibri 11.</li>
                </ul>
            </li>
            <li><b>Select Excel File:</b> Click to choose your spreadsheet of recipients.</li>
            <li><b>Excel Format:</b>
                <ul>
                    <li><b>Required:</b> A column named <b>Email</b>.</li>
                    <li><b>Optional:</b> A column named <b>Supplier Name</b>. The value from this column will be added to the start of the subject line (e.g., <i>Supplier Inc. - Your Subject</i>).</li>
                    <li><b>Optional:</b> A column named <b>Greeting</b> for a personalized opening (e.g., <i>Hi Mike,</i>).</li>
                    <li><b>Optional:</b> A column named <b>Attachment</b>. Provide the <u>full file path</u> for any attachments. <b>To add multiple attachments, separate each full path with a semicolon (;)</b>.</li>
                </ul>
            </li>
            <li><b>Preview (Recommended):</b> Check the "Preview" box to review each email before it is sent.</li>
            <li><b>Send:</b> Click "Start Sending Emails" to begin.</li>
        </ol><hr>
        <p style="font-size:12px; color:gray;">
        <b>Note:</b> Your default Outlook signature will be automatically appended to the end of every email.
        </p>
        """)
        layout.addWidget(help_text)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.exec()

def _launch_with_splash():
    logging.basicConfig(level=logging.INFO)
    logging.info("Launching splash.")
    splash = WaveSplashScreen()
    splash.show()
    main_window = EmailSender()
    def handoff():
        logging.info("Handoff: closing splash, showing main window.")
        splash.close()
        main_window.show()
    QTimer.singleShot(2600, handoff)
    return main_window

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = _launch_with_splash()
    sys.exit(app.exec())
