import sys
import os
import pandas as pd
import win32com.client as win32
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QLineEdit, QTextEdit, QFileDialog, QMessageBox, QCheckBox, QDialog
)
from PyQt6.QtCore import Qt
from pathlib import Path


class EmailSender(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MailOps: Outlook Bulk Sender")
        self.setGeometry(300, 300, 700, 620)

        layout = QVBoxLayout()

        self.help_btn = QPushButton("📘 How to Use")
        self.file_label = QLabel("No file selected")
        self.subject_input = QLineEdit()
        self.cc_input = QLineEdit()
        self.body_input = QTextEdit()
        self.preview_checkbox = QCheckBox("Preview each email before sending")
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        self.file_btn = QPushButton("Select Excel File")
        self.send_btn = QPushButton("Send Emails")

        layout.addWidget(self.help_btn)
        layout.addWidget(QLabel("Subject:"))
        layout.addWidget(self.subject_input)
        layout.addWidget(QLabel("CC Email(s): (comma or semicolon separated)"))
        layout.addWidget(self.cc_input)
        layout.addWidget(QLabel("Email Body (Plain text, will appear above your signature):"))
        layout.addWidget(self.body_input)
        layout.addWidget(self.file_btn)
        layout.addWidget(self.file_label)
        layout.addWidget(self.preview_checkbox)
        layout.addWidget(self.send_btn)
        layout.addWidget(QLabel("Log Output:"))
        layout.addWidget(self.log_output)

        self.setLayout(layout)

        self.file_btn.clicked.connect(self.select_file)
        self.send_btn.clicked.connect(self.send_emails)
        self.help_btn.clicked.connect(self.show_help_dialog)

        self.excel_path = None

    def get_signature_html(self, signature_prefix="default"):
        sig_dir = os.path.join(os.getenv("APPDATA"), "Microsoft", "Signatures")

        # Find the correct .htm file that starts with the signature_prefix
        matches = [f for f in os.listdir(sig_dir) if f.startswith(signature_prefix) and f.endswith(".htm")]
        if not matches:
            raise FileNotFoundError(f"No signature file starting with '{signature_prefix}' found in {sig_dir}")

        sig_file = matches[0]  # Use the first match
        sig_path = os.path.join(sig_dir, sig_file)
        sig_base = os.path.splitext(sig_file)[0]
        sig_files_dir = os.path.join(sig_dir, f"{sig_base}_files")

        try:
            with open(sig_path, "r", encoding="utf-8") as f:
                html = f.read()
        except UnicodeDecodeError:
            with open(sig_path, "r", encoding="cp1252") as f:
                html = f.read()

        return html, sig_files_dir

    def embed_signature_images(self, mail, html, sig_files_dir):
        if not os.path.isdir(sig_files_dir):
            return html  # No image folder? Nothing to embed.

        for file in os.listdir(sig_files_dir):
            file_path = os.path.join(sig_files_dir, file)
            content_id = f"image{file.split('.')[0]}"
            try:
                attachment = mail.Attachments.Add(file_path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"{file}"
                )
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3714001F", f"inline; filename={file}; cid:{file}"
                )
            except Exception as e:
                self.log_output.append(f"⚠️ Could not embed image: {file} — {e}")

        return html

    def select_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.excel_path = file
            self.file_label.setText(f"Selected: {file}")

    def send_emails(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            QMessageBox.critical(self, "Error", "Valid Excel file not selected.")
            return

        subject = self.subject_input.text().strip()
        cc = self.cc_input.text().strip()
        body_text = self.body_input.toPlainText().strip()
        preview_mode = self.preview_checkbox.isChecked()

        if not subject or not body_text:
            QMessageBox.warning(self, "Missing Info", "Subject and Body are required.")
            return

        try:
            df = pd.read_excel(self.excel_path)
            if 'Email' not in df.columns:
                raise ValueError("Excel must contain 'Email' column")

            sig_html, sig_img_dir = self.get_signature_html("default")
            outlook = win32.Dispatch("Outlook.Application")
            count = 0

            for index, row in df.iterrows():
                email_raw = str(row['Email']).strip()
                attachment_path = str(row['Attachment']).strip() if 'Attachment' in row and pd.notna(
                    row['Attachment']) else ""

                if not email_raw:
                    self.log_output.append(f"❌ Skipping - Missing email at row {index + 2}")
                    continue

                # Attachment is listed but missing
                if attachment_path and not os.path.isfile(attachment_path):
                    self.log_output.append(f"❌ Skipping - Attachment not found for {email_raw}: {attachment_path}")
                    continue

                try:
                    mail = outlook.CreateItem(0)
                    mail.To = email_raw
                    mail.CC = cc
                    mail.Subject = subject

                    # Construct HTML body
                    content_html = body_text.replace("\n", "<br>")
                    full_html = f"<p>{content_html}</p>{sig_html}"
                    mail.HTMLBody = full_html

                    if attachment_path:
                        mail.Attachments.Add(attachment_path)
                        attachment_note = f"Attached {os.path.basename(attachment_path)}"
                    else:
                        attachment_note = "No attachment provided"

                    self.embed_signature_images(mail, full_html, sig_img_dir)

                    if preview_mode:
                        mail.Display()
                        result = QMessageBox.question(
                            self, "Send Email?",
                            f"Send email to:\n{email_raw}\n\n{attachment_note}",
                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                        )
                        if result == QMessageBox.StandardButton.Yes:
                            mail.Send()
                            self.log_output.append(f"✅ Sent to {email_raw} — {attachment_note}")
                            count += 1
                        else:
                            self.log_output.append(f"⏭️ Skipped {email_raw}")
                    else:
                        mail.Send()
                        self.log_output.append(f"✅ Sent to {email_raw} — {attachment_note}")
                        count += 1

                except Exception as e:
                    self.log_output.append(f"❌ {email_raw} - Error: {str(e)}")

            QMessageBox.information(self, "Done", f"Finished sending. Emails sent: {count}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def show_help_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("How to Use MailOps")
        dialog.setMinimumSize(650, 550)

        layout = QVBoxLayout(dialog)

        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml("""
        <h2 style="color:#003DA5;">How to Use MailOps</h2>
        <hr>
        <h3>1. Prepare Your Excel File</h3>
        <ul>
          <li><b>Required:</b> <code>Email</code> column (e.g., someone@example.com)</li>
          <li><b>Optional:</b> <code>Attachment</code> column with full file path (e.g., C:\\Users\\You\\file.pdf)</li>
          <li>Leave <code>Attachment</code> blank to send emails without attachments</li>
          <li>Save the file as <b>.xlsx</b> format</li>
        </ul>
        <h3>2. Enter Email Details</h3>
        <ul>
          <li><b>Subject:</b> Required</li>
          <li><b>CC Emails:</b> Optional — separate with commas or semicolons</li>
          <li><b>Email Body:</b> Supports basic HTML (e.g., <code>&lt;b&gt;</code>, <code>&lt;br&gt;</code>, <code>&lt;p&gt;</code>)</li>
          <li>Avoid pasting from Word or Outlook — it can break formatting</li>
          <li>The body appears <b>above</b> your Outlook signature</li>
        </ul>
        <h3>3. Choose Preview Option</h3>
        <ul>
          <li>Enable preview to manually approve each email before sending</li>
          <li>Disable to send all emails automatically</li>
        </ul>
        <h3>4. Send Emails</h3>
        <ul>
          <li>Click <b>Send Emails</b> to begin</li>
          <li>Watch the <b>Log Output</b> section for status</li>
        </ul>
        <h3>Notes</h3>
        <ul>
          <li>Your Outlook signature (with embedded images) is automatically included</li>
          <li>If an attachment is missing, the email still sends and is logged</li>
          <li>Inline images in the signature are properly embedded using Content-ID</li>
        </ul>
        <h3>Tips</h3>
        <ul>
          <li>Test with 1–2 rows before bulk sending</li>
          <li>Use absolute paths for all attachments (not just filenames)</li>
          <li>Double-check email addresses to avoid failures</li>
        </ul>
        <hr>
<p style="text-align:center; font-size:12px; color:gray;">
  MailOps © 2025<br>
  Created and maintained by Joshua Taitt
</p>
        """)
        layout.addWidget(help_text)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)

        dialog.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailSender()
    window.show()
    sys.exit(app.exec())
