import os
import glob
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from main import process_pdf


# Get the absolute path to the GIF file
gif_path = os.path.abspath('/home/justin/PycharmProjects/kivy-designer/venv/pig.gif')


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

    def choose_folder(self):
        # Show the file dialog to choose a folder
        folder_path = filedialog.askdirectory()

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

    def run_script(self):
        # Disable the start button
        self.start_button.config(state='disabled')

        # Change the label text to "Converting Files......"
        self.label.config(text='Converting Files......')

        # Create a popup with the folder path, converting message, and a spinner
        self.popup = tk.Toplevel()
        self.popup.title('Workin\' on it')
        self.popup.geometry('400x200')
        popup_label = tk.Label(self.popup,
                               text=f'Selected folder path: {self.folder_path}\n\nConverting your files, please hold tight......',
                               font=('Arial', 14))
        popup_label.pack(pady=10)

        # Replace the 'spinner.gif' with the path to your spinner gif
        spinner = tk.PhotoImage(file=gif_path)
        spinner_label = tk.Label(self.popup, image=spinner)
        spinner_label.image = spinner
        spinner_label.pack(pady=10)

        # Run the main.py script in a separate thread
        self.thread = threading.Thread(target=self.run_main_script, args=(self.folder_path,))
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