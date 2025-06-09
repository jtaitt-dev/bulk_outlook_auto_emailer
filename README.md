# 📧 MailOps: Outlook Bulk Email Sender (with Signature + Attachments)

A PyQt6 desktop tool for sending personalized Outlook emails in bulk using Excel. Supports embedded Outlook signatures, inline images, file attachments, optional previewing, and logging.

---

## 🚀 Features

- ✅ Load recipients from Excel (`Email` column required, `Attachment` optional)
- 🖋️ Automatically inserts your default Outlook HTML signature (with embedded images)
- 📎 Adds individual attachments per recipient if specified
- 👁️ Optional preview mode to confirm each email before sending
- 📝 Log output panel to track sent, skipped, and failed emails
- 📤 Sends via Outlook using `win32com.client` for full compatibility

---

## ⚠️ Outlook Compatibility

> **Requires the classic (legacy) version of Outlook for Windows**  
> This app uses `win32com.client`, which is not supported by the new Outlook experience (as of 2024).  
> To use MailOps:
> - You must use the **classic Outlook** desktop client.
> - Do **not** enable the "New Outlook" toggle.

---

## 📁 Excel Template Format

| Email               | Attachment                        |
|--------------------|------------------------------------|
| user1@example.com  | `C:\Files\file1.pdf`               |
| user2@example.com  | *(leave blank to skip attachment)* |

---

## 🛠️ How to Use

1. **Launch the App**
   - Run `python main.py` or use your preferred IDE (e.g., PyCharm).

2. **Prepare Your File**
   - Create an Excel file (`.xlsx`) with:
     - `Email` column (required)
     - `Attachment` column (optional, full path)

3. **Fill Out Email Details**
   - Subject
   - CC (comma or semicolon separated)
   - Email Body (plain text; will appear above your signature)

4. **Select File**
   - Click **Select Excel File** and load your recipient list.

5. **(Optional)** Enable Preview Mode
   - Preview and confirm each email manually before sending.

6. **Send Emails**
   - Click **Send Emails** and monitor the Log Output panel.

---

## 🖋️ Signature Handling

- Pulls your default Outlook HTML signature from:
%APPDATA%\Microsoft\Signatures\

- Embedded images are handled using `Content-ID` references for proper rendering.

---

## 🔧 Requirements

- Windows OS with **classic Outlook** installed and configured
- Python 3.10+
- `pip install -r requirements.txt`

### Dependencies:
```bash
pip install pyqt6 pandas pywin32


🧠 Tips

    Run a test with a few rows before sending to a large list.

    Ensure attachment paths are valid and absolute.

    Avoid pasting from Word/Outlook into the body field—use plain text or basic HTML.

📦 Packaging (Optional)

You can use tools like pyinstaller to package this into an .exe:
pyinstaller --noconfirm --windowed --onefile main.py

