import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pywhatkit as pwk
import time
import os

def browse_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", ".xlsx"), ("All files", ".*")]
    )
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filepath)

def browse_media():
    filepath = filedialog.askopenfilename(
        filetypes=[("Image files", "*.jpg;*.jpeg;*.png"), ("Document files", "*.pdf;*.docx"), ("All files", "*.*")]
    )
    media_entry.delete(0, tk.END)
    media_entry.insert(0, filepath)

def send_whatsapp_message(recipient_number, message_body, media_file=None):
    try:
        if media_file:
            # Ensure the media file exists
            if not os.path.exists(media_file):
                messagebox.showerror("Error", "Media file not found.")
                return

            file_extension = os.path.splitext(media_file)[1].lower()

            if file_extension in [".jpg", ".jpeg", ".png"]:
                # Send image
                print(f"Sending image to {recipient_number}")
                pwk.sendwhats_image(recipient_number, media_file, message_body, 15)
            elif file_extension in [".pdf", ".docx"]:
                # Send document
                print(f"Sending document to {recipient_number}")
                pwk.sendwhats_document(recipient_number, media_file, message_body)
            else:
                messagebox.showerror("Error", "Unsupported media file format.")
                return
        else:
            # Send text message
            print(f"Sending message to {recipient_number}")
            pwk.sendwhatmsg_instantly(recipient_number, message_body)
    except Exception as e:
        print(f"Failed to send message to {recipient_number}. Error: {str(e)}")

def send_bulk_whatsapp():
    message_body = message_entry.get("1.0", tk.END).strip()
    excel_file = file_entry.get()
    media_file = media_entry.get()

    if not excel_file:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    if not message_body and not media_file:
        messagebox.showerror("Error", "Please enter a message or select a media file to send.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)

        if 'WhatsApp' not in df.columns:
            messagebox.showerror("Error", "'WhatsApp' column not found in Excel file.")
            return

        recipient_numbers = df['WhatsApp'].dropna().tolist()

        for recipient_number in recipient_numbers:
            recipient_number = str(recipient_number).strip()

            # Add country code if missing
            if not recipient_number.startswith("+"):
                country_code = "+91"  # Replace '91' with your desired country code
                recipient_number = f"{country_code}{recipient_number}"

            # Validate number format
            if len(recipient_number) <= 10 or not recipient_number[1:].isdigit():
                print(f"Skipping invalid number: {recipient_number}")
                continue

            print(f"Sending message to: {recipient_number}")

            # Send the media if available, otherwise send the message
            send_whatsapp_message(recipient_number, message_body, media_file if media_file else None)
            time.sleep(15)  # Delay to avoid conflicts and allow time for the browser to process

        messagebox.showinfo("Success", "WhatsApp messages have been sent!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to send WhatsApp messages. Error: {str(e)}")

# GUI setup
root = tk.Tk()
root.title("Bulk WhatsApp Sender from Excel")

tk.Label(root, text="Message:").grid(row=0, column=0, padx=10, pady=10)
message_entry = tk.Text(root, width=50, height=10)
message_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Excel File:").grid(row=1, column=0, padx=10, pady=10)
file_entry = tk.Entry(root, width=50)
file_entry.grid(row=1, column=1, padx=10, pady=10)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Select Media (Image/Document):").grid(row=2, column=0, padx=10, pady=10)
media_entry = tk.Entry(root, width=50)
media_entry.grid(row=2, column=1, padx=10, pady=10)

browse_media_button = tk.Button(root, text="Browse Media", command=browse_media)
browse_media_button.grid(row=2, column=2, padx=10, pady=10)

send_button = tk.Button(root, text="Send WhatsApp Messages", command=send_bulk_whatsapp)
send_button.grid(row=3, column=1, padx=10, pady=20)

root.mainloop()
