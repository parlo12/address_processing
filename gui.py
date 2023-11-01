import tkinter as tk
from tkinter import filedialog
from threading import Thread
from excel_processor import ExcelProcessor
import shutil
import os
class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel Address Processor')
        self.input_file = ''
        self.output_file = 'output.xlsx'
        self.processor = ExcelProcessor()
        self.create_widgets()
        
    def create_widgets(self):
        tk.Label(self.root, text='Excel Address Processor', font=('Helvetica', 16)).pack(pady=20)
        tk.Button(self.root, text='Open Excel File', command=self.load_file).pack(pady=10)
        
        self.file_label = tk.Label(self.root, text='No file selected')
        self.file_label.pack(pady=10)
        
        tk.Button(self.root, text='Process File', command=self.process_file).pack(pady=10)
        self.status_label = tk.Label(self.root, text='')
        self.status_label.pack(pady=20)

        # Add the export button
        tk.Button(self.root, text='Export File', command=self.export_file).pack(pady=10)
    
    def load_file(self):
        self.input_file = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'), ('Excel Files', '*.xls'), ('All Files', '*.*')])
        if self.input_file:
            self.file_label.config(text=self.input_file.split("/")[-1])
            self.status_label.config(text='File loaded successfullly, ready to process.')
        else:
            self.status_label.config(text='No file selected')

    def process_file(self):
        if self.input_file:
            self.status_label.config(text='Processing file...')
            thread = Thread(target=self.processor.process_excel_file, args=(self.input_file, self.output_file, self.post_processing))
            thread.start()
        else:
            self.status_label.config(text='No file selected')
    def post_processing(self):
        # Code to run after processing the Excel file
        self.status_label.config(text='Process completed. The file is saved as {}'.format(self.output_file))
        print('Process completed. The file is saved as {}'.format(self.output_file))
    def export_file(self):
        if not os.path.exists(self.output_file):
            self.status_label.config(text='No file to export')
            return
        # Ask the user to select a fold to save the file
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not export_path:
            self.status_label.config(text='No file to export')
            return
        # Copy the file to the export path
        shutil.copy(self.output_file, export_path)
        self.status_label.config(text=f'File exported to {export_path}')   
        

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
