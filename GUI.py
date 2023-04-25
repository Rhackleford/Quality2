import os
import glob
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from main import process_pdf_directory


class MainApp:
    def __init__(self, master):
        # Create the layout for the UI
        self.master = master
        master.title('PDF Converter')
        master.geometry('600x400')

        # Create the label for the file chooser
        self.label = tk.Label(text='Select the input folder:', font=('Arial', 14))
        self.label.pack(pady=10)

        # Create the file chooser button
        self.choose_button = tk.Button(text='Choose Folder', font=('Arial', 12), command=self.choose_folder)
        self.choose_button.pack(pady=10)

        # Create the button to start the script
        self.start_button = tk.Button(text='Start', font=('Arial', 12), state='disabled', command=self.run_script)
        self.start_button.pack(pady=10)

        # Create a listbox to display the folder contents
        self.file_listbox = tk.Listbox(self.master, width=80, height=10)
        self.file_listbox.pack(pady=10)

    def choose_folder(self):
        # Show the file dialog to choose a folder
        folder_path = filedialog.askdirectory(mustexist=True)



        # Check for PDFs with "- Part List" in their file names
        pdf_files = glob.glob(f"{folder_path}/*- Part List*.pdf")

        if not pdf_files:
            # No suitable PDFs found, show an error message box
            messagebox.showerror(title='No suitable PDFs found',
                                 message='The folder you chose has no suitable PDFs, please try another')
            self.start_button.config(state='disabled')
        else:
            # Enable the start button
            self.start_button.config(state='normal')
            self.folder_path = folder_path

            # Clear the listbox
            self.file_listbox.delete(0, tk.END)

            # Add folder contents to the listbox
            for file in os.listdir(folder_path):
                self.file_listbox.insert(tk.END, file)

    def process(self, folder_path):
        # Define the area_mm and column_positions_mm variables with the appropriate values
        area_mm = [0, 0, 196.85, 279.4]  # top, left, bottom, and right coordinates in millimeters
        column_positions_mm = [15, 76, 90, 96, 98, 113, 164, 260]  # Approximate column positions in millimeters

        # Process all PDF files in the folder_path directory
        process_pdf_directory(folder_path, area_mm, column_positions_mm)

    def run_script(self):
        # Disable the start button
        self.start_button.config(state='disabled')

        # Change the label text to "Converting Files......"
        self.label.config(text='Converting Files......')

        # Create a popup with the folder path, converting message, and a progress bar
        self.popup = tk.Toplevel()
        self.popup.title('Working on it....')
        self.popup.geometry('500x200')

        # Calculate the new position for the popup
        root_x = self.master.winfo_x()
        root_y = self.master.winfo_y()
        root_width = self.master.winfo_width()
        root_height = self.master.winfo_height()
        offset_x = (root_width - 500) // 2
        offset_y = int(root_height * 0.33)
        new_x = root_x + offset_x
        new_y = root_y + offset_y

        # Move the popup window to the new position
        self.popup.geometry(f"+{new_x}+{new_y}")

        popup_label = tk.Label(self.popup,
                               text=f'Selected folder path: {self.folder_path}\n\nConverting your files, please hold tight......',
                               font=('Arial', 10))
        popup_label.pack(pady=10)

        # Create a progress bar
        progress_bar = ttk.Progressbar(self.popup, mode='indeterminate', length=300)
        progress_bar.pack(pady=10)
        progress_bar.start()

        # Run the main.py script in a separate thread
        self.thread = threading.Thread(target=self.process, args=(self.folder_path,))
        self.thread.start()

        # Schedule a function to periodically check if the thread is still running
        self.master.after(1000, self.check_thread_status)

    def check_thread_status(self):
        if self.thread.is_alive():
            # Schedule the function to check the thread status again in 1 second
            self.master.after(1000, self.check_thread_status)
        else:
            # Thread is not running, change the label text
            self.label.config(text='Conversion Finished, The Window Is Safe To Close Now.')

            # Dismiss the popup
            self.popup.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
