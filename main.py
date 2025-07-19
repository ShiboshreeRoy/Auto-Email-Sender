import smtplib
from email.message import EmailMessage
from tkinter import *
from tkinter import filedialog, messagebox, scrolledtext, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import re
import logging
import threading

logging.basicConfig(
    filename="email_sender.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📧 Premium Auto Email Sender")
        self.root.geometry("800x820")
        self.root.configure(bg="#f2f2f7")  # macOS-like background
        self.attachments = []

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=6, relief="flat", background="#007AFF", foreground="white", font=('Helvetica Neue', 11, 'bold'))
        style.map("TButton",
                  foreground=[('pressed', 'white'), ('active', 'white')],
                  background=[('pressed', '#005BBB'), ('active', '#339CFF')])
        style.configure("TLabel", font=('Helvetica Neue', 11), background="#f2f2f7")
        style.configure("TEntry", font=('Helvetica Neue', 11))
        style.configure("TFrame", background="#f2f2f7")

        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=BOTH, expand=True)

        header = Label(main_frame, text="📬 Auto Email Sender", font=('Helvetica Neue', 20, 'bold'), bg="#f2f2f7", fg="#333")
        header.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        # Menu Bar for About Section
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        about_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=about_menu)
        about_menu.add_command(label="About Developer", command=self.show_about)


        entries = [
            ("Your Email:", "sender_entry"),
            ("App Password:", "password_entry"),
            ("SMTP Server (e.g., smtp.gmail.com):", "smtp_entry"),
            ("Port (e.g., 587):", "port_entry"),
            ("To Emails (comma separated):", "receiver_entry"),
            ("Subject:", "subject_entry"),
        ]
        for i, (label_text, attr) in enumerate(entries, start=1):
            ttk.Label(main_frame, text=label_text).grid(row=i, column=0, sticky=E, pady=6, padx=10)
            entry = ttk.Entry(main_frame, width=55, show='*' if 'password' in attr else '')
            entry.grid(row=i, column=1, pady=6)
            setattr(self, attr, entry)

        self.smtp_entry.insert(0, "smtp.gmail.com")
        self.port_entry.insert(0, "587")

        ttk.Label(main_frame, text="Message:").grid(row=7, column=0, sticky=NE, padx=10, pady=6)
        self.body_text = scrolledtext.ScrolledText(main_frame, width=55, height=10, font=('Helvetica Neue', 11), wrap=WORD, bg="white")
        self.body_text.grid(row=7, column=1, pady=6)

        ttk.Label(main_frame, text="Attachments:").grid(row=8, column=0, sticky=NE, padx=10, pady=6)
        self.drop_area = Listbox(main_frame, selectmode=MULTIPLE, width=55, height=5, bg="#ffffff", font=('Helvetica Neue', 10))
        self.drop_area.grid(row=8, column=1, pady=6)
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.drop_files)

        ttk.Button(main_frame, text="📁 Add Attachment", command=self.browse_files).grid(row=9, column=1, sticky=W, pady=(10, 0))

        self.send_btn = ttk.Button(main_frame, text="🚀 Send Email", command=self.send_email_threaded)
        self.send_btn.grid(row=9, column=1, sticky=E, pady=20)

        self.status_label = ttk.Label(main_frame, text="", foreground="#555")
        self.status_label.grid(row=11, column=1, sticky=W, pady=4)

    def browse_files(self):
        files = filedialog.askopenfilenames()
        self.add_attachments(files)

    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        self.add_attachments(files)

    def add_attachments(self, files):
        for file in files:
            if os.path.isfile(file) and file not in self.attachments:
                self.attachments.append(file)
                size = os.path.getsize(file) / (1024 * 1024)
                self.drop_area.insert(END, f"{os.path.basename(file)} ({size:.2f} MB)")

    def validate_email(self, email):
        pattern = r"[^@]+@[^@]+\.[^@]+"
        return re.match(pattern, email)

    def send_email_threaded(self):
        threading.Thread(target=self.send_email).start()

    def send_email(self):
        self.status_label.config(text="📤 Sending email...")
        self.send_btn.state(["disabled"])

        sender = self.sender_entry.get().strip()
        password = self.password_entry.get().strip()
        smtp_server = self.smtp_entry.get().strip()
        port = self.port_entry.get().strip()
        receivers = [email.strip() for email in self.receiver_entry.get().split(",") if email.strip()]
        subject = self.subject_entry.get().strip()
        body = self.body_text.get("1.0", END).strip()

        if not sender or not password or not receivers:
            messagebox.showerror("Error", "Sender email, password, and at least one recipient are required.")
            self.send_btn.state(["!disabled"])
            return

        if not self.validate_email(sender) or not all(self.validate_email(r) for r in receivers):
            messagebox.showerror("Error", "Please enter valid email addresses.")
            self.send_btn.state(["!disabled"])
            return

        msg = EmailMessage()
        msg['From'] = sender
        msg['To'] = ", ".join(receivers)
        msg['Subject'] = subject
        msg.set_content(body)

        for file_path in self.attachments:
            try:
                with open(file_path, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(file_path)
                    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
            except Exception as e:
                messagebox.showerror("Attachment Error", f"Failed to attach {file_path}:\n{e}")
                logging.error(f"Attachment error: {file_path} - {e}")
                self.send_btn.state(["!disabled"])
                return

        try:
            with smtplib.SMTP(smtp_server, int(port)) as server:
                server.starttls()
                server.login(sender, password)
                server.send_message(msg)
            messagebox.showinfo("Success", "✅ Email sent successfully!")
            logging.info(f"Email sent from {sender} to {receivers}")
            self.status_label.config(text="✅ Email sent successfully.")
        except smtplib.SMTPAuthenticationError as e:
            messagebox.showerror("Authentication Error", "❌ Login failed. Please use a valid App Password.\nCheck: https://support.google.com/mail/?p=BadCredentials")
            logging.error(f"Authentication error: {e}")
            self.status_label.config(text="❌ Authentication failed.")
        except Exception as e:
            messagebox.showerror("Error", f"❌ Failed to send email:\n{e}")
            logging.error(f"Email send error: {e}")
            self.status_label.config(text="❌ Failed to send email.")
        finally:
            self.send_btn.state(["!disabled"])

    def show_about(self):
        messagebox.showinfo(
        "About Developer",
        "👨‍💻 Developer: Shiboshree Roy\n"
        "📧 Email: shiboshreeroycse@gmail.com\n"
        "💼 Full Stack Web Developer\n"
        "🌐 Open Source Contributor\n"
        "⚙️ Researcher & Programmer\n"
        
    )



if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
