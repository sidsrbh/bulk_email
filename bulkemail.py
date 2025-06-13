import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, simpledialog
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import threading

class SettingsDialog(simpledialog.Dialog):
    def body(self, master):
        tk.Label(master, text="Email:").grid(row=0)
        tk.Label(master, text="Password:").grid(row=1)
        tk.Label(master, text="SMTP Server:").grid(row=2)
        tk.Label(master, text="Port:").grid(row=3)

        self.email_entry = tk.Entry(master)
        self.password_entry = tk.Entry(master, show="*")
        self.smtp_entry = tk.Entry(master)
        self.port_entry = tk.Entry(master)

        self.email_entry.grid(row=0, column=1)
        self.password_entry.grid(row=1, column=1)
        self.smtp_entry.grid(row=2, column=1)
        self.port_entry.grid(row=3, column=1)

        self.email_entry.insert(0, app.sender_email)
        self.password_entry.insert(0, app.sender_password)
        self.smtp_entry.insert(0, app.smtp_server)
        self.port_entry.insert(0, str(app.smtp_port))

        return self.email_entry

    def apply(self):
        app.sender_email = self.email_entry.get()
        app.sender_password = self.password_entry.get()
        app.smtp_server = self.smtp_entry.get()
        app.smtp_port = int(self.port_entry.get())

class BulkEmailApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Bulk Email Sender')
        self.geometry("600x700")  # Adjusted for additional components

        self.excel_file = None
        self.emails_df = pd.DataFrame()

        # Email details (default values)
        self.sender_email = "your-email@domain.com"
        self.sender_password = "yourpassword"
        self.smtp_server = "smtp.hostinger.com"
        self.smtp_port = 465

        # UI Components setup
        tk.Button(self, text='Settings', command=self.open_settings).pack(pady=(5,0))

        tk.Label(self, text='Subject:').pack(pady=(5,0))
        self.subject_entry = tk.Entry(self, width=60)
        self.subject_entry.pack(pady=5)
        self.subject_entry.insert(0, "Your Subject Here")

        tk.Label(self, text='Body:').pack(pady=(5,0))
        self.body_text = scrolledtext.ScrolledText(self, height=10, width=45)
        self.body_text.pack(pady=5)
        self.body_text.insert(tk.END, "Dear User,\n\nThis is the main content of your email. Customize it as per your requirement.\n\nBest regards,\nYour Name")

        tk.Button(self, text='Upload Excel File', command=self.upload_excel).pack(pady=10)
        self.preview_label = tk.Label(self, text="")
        self.preview_label.pack(pady=5)

        tk.Button(self, text='Send Emails', command=self.initiate_email_sending).pack(pady=10)
        
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(pady=10)

        self.status_text = scrolledtext.ScrolledText(self, height=10, width=45, state='disabled')
        self.status_text.pack(pady=10)

    def open_settings(self):
        SettingsDialog(self)

    def log_message(self, message):
        self.status_text.configure(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.configure(state='disabled')
        self.status_text.see(tk.END)

    def upload_excel(self):
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.excel_file:
            self.emails_df = pd.read_excel(self.excel_file)
            if 'Email' in self.emails_df.columns:
                self.preview_label.configure(text=f"Loaded {len(self.emails_df)} emails.")
                messagebox.showinfo("Success", "Excel file uploaded successfully.")
            else:
                messagebox.showerror("Error", "No 'Email' column found in the Excel file.")
                self.preview_label.configure(text="")

    def initiate_email_sending(self):
        if not self.excel_file or self.emails_df.empty:
            messagebox.showerror("Error", "Please upload a valid Excel file first.")
            return

        # Correctly reference the method with self.send_emails_from_excel
        thread = threading.Thread(target=self.send_emails_from_excel, daemon=True)
        thread.start()


    def send_emails_from_excel(self):
        total_emails = len(self.emails_df['Email'].dropna().tolist())
        self.progress['maximum'] = total_emails
        self.progress['value'] = 0

        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", tk.END)

        for index, email in enumerate(self.emails_df['Email'].dropna().tolist(), start=1):
            try:
                self.send_bulk_email(self.sender_email, self.sender_password, [email], subject, body)
                self.log_message(f"Email sent to: {email}")
            except Exception as e:
                self.log_message(f"Failed to send to {email}: {e}")
            self.progress['value'] = index
            self.update_idletasks()

        self.log_message("Finished sending emails.")

    def send_bulk_email(self, sender_email, sender_password, recipient_emails, subject, body):
        server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
        server.login(sender_email, sender_password)
        for email in recipient_emails:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            server.send_message(msg)
        server.quit()

if __name__ == "__main__":
    app = BulkEmailApp()
    app.mainloop()

