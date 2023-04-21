import os
import glob
import threading
import kivy
from kivy.app import App
from kivy.clock import Clock
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.image import Image
from kivy.uix.image import AsyncImage
import main
# Get the absolute path to the GIF file
gif_path = os.path.abspath('/home/justin/PycharmProjects/kivy-designer/venv/pig.gif')

class MainApp(App):
    def build(self):
        # Create the layout for the UI
        layout = BoxLayout(orientation='vertical')

        # Create the label for the file chooser
        self.label = Label(text='Select the input folder:')

        # Create the file chooser
        file_chooser = FileChooserIconView()

        # Create the button to start the script
        button = Button(text='Start', size_hint=(1, 0.2))

        # Bind the button press to the function that runs the script
        button.bind(on_press=self.run_script)

        # Add the label, file chooser, and button to the layout
        layout.add_widget(self.label)
        layout.add_widget(file_chooser)
        layout.add_widget(button)

        return layout

    def run_script(self, button):
        # Get the path to the selected folder from the file chooser
        folder_path = button.parent.children[1].path

        # Check for PDFs with "- Part List" in their file names
        pdf_files = glob.glob(f"{folder_path}/*- Part List*.pdf")

        if not pdf_files:
            # No suitable PDFs found, show an error popup
            content = Label(text='The folder you chose has no suitable PDFs, please try another')
            popup = Popup(title='No suitable PDFs found', content=content, size_hint=(0.6, 0.6))
            popup.open()
        else:
            # Change the label text to "Converting Files......"
            self.label.text = "Converting Files......"

            # Create a popup with the folder path, converting message, and a spinner
            content = BoxLayout(orientation='vertical')
            content.add_widget(
                Label(text=f'Selected folder path: {folder_path}\n\nConverting your files, please hold tight......'))

            # Replace the 'spinner.gif' with the path to your spinner gif
            spinner = AsyncImage(source=gif_path, anim_delay=0.1)

            content.add_widget(spinner)

            self.popup = Popup(title='Workin\' on it', content=content, size_hint=(0.6, 0.6))
            self.popup.open()

            # Run the main.py script in a separate thread
            self.thread = threading.Thread(target=self.run_main_script, args=(folder_path,))
            self.thread.start()

            # Schedule a function to periodically check if the thread is still running
            Clock.schedule_interval(self.check_thread_status, 1)

    def run_main_script(self, folder_path):
        area_mm = [0, 0, 196.85, 279.4]  # top, left, bottom, and right coordinates in millimeters
        column_positions_mm = [15, 76, 90, 96, 98, 113, 164, 260]  # Approximate column positions in millimeters
        main.process_pdf_directory(input_directory=folder_path, area_mm=area_mm,
                                   column_positions_mm=column_positions_mm)

    area_mm = [0, 0, 196.85, 279.4]  # top, left, bottom, and right coordinates in millimeters
    column_positions_mm = [15, 76, 90, 96, 98, 113, 164, 260]  # Approximate column positions in millimeters
    def check_thread_status(self, dt):
        if not self.thread.is_alive():
            # Thread is not running, change the label text
            self.label.text = "Conversion Finished, The Window Is Safe To Close Now."
            # Dismiss the popup
            self.popup.dismiss()

            # Unschedule the function
            Clock.unschedule(self.check_thread_status)

if __name__ == '__main__':
    MainApp().run()