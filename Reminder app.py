import pandas as pd
from datetime import datetime
from plyer import notification
import tkinter as tk
from tkinter import Label, Entry, Button

REMINDERS_FILE = "reminders.xlsx"

class ReminderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Reminder App")
        self.root.attributes('-topmost', True)

        # Create and set up GUI elements
        self.title_label = Label(root, text="Title:")
        self.title_entry = Entry(root)

        self.message_label = Label(root, text="Message:")
        self.message_entry = Entry(root)

        self.date_label = Label(root, text="Date (YYYY-MM-DD):")
        self.date_entry = Entry(root)

        self.time_label = Label(root, text="Time (HH:MM):")
        self.time_entry = Entry(root)

        self.set_button = Button(root, text="Set Reminder", command=self.set_reminder)

        # Grid layout
        self.title_label.grid(row=0, column=0, pady=5)
        self.title_entry.grid(row=0, column=1, pady=5)

        self.message_label.grid(row=1, column=0, pady=5)
        self.message_entry.grid(row=1, column=1, pady=5)

        self.date_label.grid(row=2, column=0, pady=5)
        self.date_entry.grid(row=2, column=1, pady=5)

        self.time_label.grid(row=3, column=0, pady=5)
        self.time_entry.grid(row=3, column=1, pady=5)

        self.set_button.grid(row=4, column=0, columnspan=2, pady=10)

        # Check reminders periodically
        self.root.after(1000, self.check_reminders)

    def set_reminder(self):
        title = self.title_entry.get()
        message = self.message_entry.get()
        date_str = self.date_entry.get()
        time_str = self.time_entry.get()

        # Check if any field is empty before setting a reminder
        if not all([title, message, date_str, time_str]):
            self.show_error_message("Please fill in all fields before setting a reminder.")
            return

        try:
            reminder_datetime = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
            delay = (reminder_datetime - datetime.now()).total_seconds()

            if delay < 0:
                raise ValueError("Invalid date and time. Please enter a future date and time.")

            save_reminder_to_excel(reminder_datetime, title, message)

            # Schedule notification and remove reminder after notification
            self.root.after(int(delay * 1000), lambda: self.show_notification(title, message))

        except ValueError as e:
            self.show_error_message(str(e))

    def show_notification(self, title, message):
        notification.notify(
            title=title,
            message=message,
            app_icon=None,
            timeout=10,
        )

        remove_completed_reminders()
        self.clear_input_fields()

    def show_error_message(self, message):
        error_window = tk.Toplevel(self.root)
        error_window.title("Error")
        error_window.transient(self.root)  # Set the parent window
        Label(error_window, text=message, padx=20, pady=10).pack()
        Button(error_window, text="OK", command=error_window.destroy).pack()
        error_window.grab_set()  # Make the error window modal
        self.root.wait_window(error_window)  # Wait until the error window is closed

    def clear_input_fields(self):
        self.title_entry.delete(0, 'end')
        self.message_entry.delete(0, 'end')
        self.date_entry.delete(0, 'end')
        self.time_entry.delete(0, 'end')

    def check_reminders(self):
        # Schedule the next check
        self.root.after(1000, self.check_reminders)

def save_reminder_to_excel(reminder_datetime, title, message):
    try:
        df = pd.read_excel(REMINDERS_FILE)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Datetime", "Title", "Message"])

    new_row = {"Datetime": reminder_datetime, "Title": title, "Message": message}

    if df.empty:
        df = pd.DataFrame([new_row])
    else:
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True, sort=False)

    df.to_excel(REMINDERS_FILE, index=False)

def remove_completed_reminders():
    try:
        df = pd.read_excel(REMINDERS_FILE)
        now = datetime.now()

        # Filter out completed reminders
        df = df[df["Datetime"] > now]

        # Save the filtered DataFrame back to the Excel file
        df.to_excel(REMINDERS_FILE, index=False)

    except FileNotFoundError:
        # Create an empty DataFrame if the file does not exist
        df = pd.DataFrame(columns=["Datetime", "Title", "Message"])
        df.to_excel(REMINDERS_FILE, index=False)

if __name__ == "__main__":
    root = tk.Tk()
    app = ReminderApp(root)
    root.mainloop()
