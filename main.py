import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import scrolledtext
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders




class EmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Sender")
        self.attachments = []

        # Recipients file
        recipients_label = tk.Label(root, text="Recipients File:")
        recipients_label.grid(row=0, column=0, padx=5, pady=5)
        self.recipients_entry = tk.Entry(root, width=50)
        self.recipients_entry.grid(row=0, column=1, padx=5, pady=5)
        recipients_button = tk.Button(root, text="Browse", command=self.select_recipients_file)
        recipients_button.grid(row=0, column=2, padx=5, pady=5)

        # Subject
        subject_label = tk.Label(root, text="Subject:")
        subject_label.grid(row=1, column=0, padx=5, pady=5)
        self.subject_entry = tk.Entry(root, width=50)
        self.subject_entry.grid(row=1, column=1, padx=5, pady=5)

        # Body
        body_label = tk.Label(root, text="Body:")
        body_label.grid(row=2, column=0, padx=5, pady=5)
        self.body_text = scrolledtext.ScrolledText(root, width=50, height=10)
        self.body_text.grid(row=2, column=1, padx=5, pady=5)

        # Attachments
        attachments_label = tk.Label(root, text="Attachments:")
        attachments_label.grid(row=3, column=0, padx=5, pady=5)
        self.attachments_text = tk.Text(root, width=50, height=5)
        self.attachments_text.grid(row=3, column=1, padx=5, pady=5)
        attachments_button_frame = tk.Frame(root)
        attachments_button_frame.grid(row=3, column=2, padx=5, pady=5)
        add_button = tk.Button(attachments_button_frame, text="Add", command=self.add_attachment)
        add_button.pack(side=tk.TOP)
        remove_button = tk.Button(attachments_button_frame, text="Remove", command=self.remove_attachment)
        remove_button.pack(side=tk.TOP)

        # Send button
        send_button = tk.Button(root, text="Send", command=self.send_email)
        send_button.grid(row=4, column=1, padx=5, pady=5)

    def select_recipients_file(self):
        file_path = filedialog.askopenfilename(title="Select Recipients File", filetypes=[("Text Files", "*.txt")])
        self.recipients_entry.delete(0, tk.END)
        self.recipients_entry.insert(0, file_path)

    def add_attachment(self):
        file_paths = filedialog.askopenfilenames(title="Select Attachments")
        self.attachments.extend(file_paths)
        self.attachments_text.delete(1.0, tk.END)
        for attachment in self.attachments:
            self.attachments_text.insert(tk.END, attachment + "\n")

    def remove_attachment(self):
        try:
            selected_attachment = self.attachments_text.selection_get()
            self.attachments.remove(selected_attachment.strip())
            self.attachments_text.delete(1.0, tk.END)
            for attachment in self.attachments:
                self.attachments_text.insert(tk.END, attachment + "\n")
        except tk.TclError:
            pass
    def send_email(self):
        try:
            # Get the values from the input fields
            recipients_file = self.recipients_entry.get()
            subject = self.subject_entry.get()
            body = self.body_text.get("1.0", tk.END)
            attachments = self.attachments

            # Open the file and read the recipients' email addresses
            with open(recipients_file, 'r') as f:
                recipients = [line.strip() for line in f]

            smtp_server = "smtp.gmail.com"
            port = 587

            # create smtp connection
            server = smtplib.SMTP(smtp_server, port)
            server.ehlo()
            server.starttls()

            # login to server
            email = "Insert Your Email Id here"
            app_password = "Insert 16-character password of your app password"
            server.login(email, app_password)

            # create message
            for recipient in recipients:
                msg = MIMEMultipart()
                msg['From'] = email
                msg['To'] = recipient
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'plain'))

                # attach files
                for attachment in attachments:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(open(attachment, 'rb').read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={attachment}')
                    msg.attach(part)

                # send mail to each recipient
                server.sendmail(email, recipient, msg.as_string())
                print(f"Mail sent successfully to {recipient}")

            server.quit()

            # Display a success message
            messagebox.showinfo("Email Sent", "Emails sent successfully to all recipients!")

        except Exception as e:
            print("Error sending email:", str(e))
            messagebox.showerror("Error", "Error sending email: " + str(e))



if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSender(root)
    root.mainloop()      