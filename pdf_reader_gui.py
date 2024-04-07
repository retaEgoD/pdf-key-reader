import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, askdirectory
import sv_ttk
import ctypes
from pdf_reader import find_and_save_values_in_pdf

ctypes.windll.shcore.SetProcessDpiAwareness(True)

# TODO Add entry field for output filename, add text explaination, reorder excel and pdf, format

class PDF_Reader_GUI:
        
        def __init__(self):
            
            self.root = tk.Tk()
            self.pdf_filepath = ''
            self.excel_filepath = ''
            self.output_directory = ''
            self.output_filename = 'testing.txt'
            
            
            self.pdf_filepath_label = ttk.Label(font=("Arial", 14))
            self.pdf_frame = ttk.Frame(master=self.root, borderwidth=1)
            self.pdf_label = ttk.Label(master=self.pdf_frame, text="Select a PDF file:", font=("Arial", 14))
            self.pdf_button = tk.Button(master=self.pdf_frame, text="Select PDF File", font=("Arial", 14), borderwidth=5, relief=tk.GROOVE, command=self.select_pdf_file)
            
            self.pdf_label.pack()
            self.pdf_button.pack()
            self.pdf_frame.grid(row=1, column=0)
            self.pdf_filepath_label.grid(row=1, column=1)
            
            self.excel_filepath_label = ttk.Label(font=("Arial", 14))
            self.excel_frame = ttk.Frame(master=self.root, borderwidth=1, width=2000, height=1000)
            self.excel_label = ttk.Label(master=self.excel_frame, text="Select an Excel file:", font=("Arial", 14), anchor='w')
            self.excel_button = tk.Button(master=self.excel_frame, text="Select Excel File", font=("Arial", 14), borderwidth=5, relief=tk.GROOVE, command=self.select_excel_file)
            
            self.excel_label.pack()
            self.excel_button.pack()
            self.excel_frame.grid(row=0, column=0)
            self.excel_filepath_label.grid(row=0, column=1)
            
            
            self.output_directory_label = ttk.Label(font=("Arial", 14))
            self.output_frame = ttk.Frame(master=self.root, borderwidth=1)
            self.output_label = ttk.Label(master=self.output_frame, text="Select an output location:", font=("Arial", 14))
            self.output_button = tk.Button(master=self.output_frame, text="Select Output Location", font=("Arial", 14), borderwidth=5, relief=tk.GROOVE, command=self.select_output_directory)
            
            self.output_label.pack()
            self.output_button.pack()
            self.output_frame.grid(row=2, column=0)
            self.output_directory_label.grid(row=2, column=1)
            
            self.find_values_button = tk.Button(text="Find Values", font=("Arial", 14), borderwidth=5, relief=tk.GROOVE, command=lambda: find_and_save_values_in_pdf(self.pdf_filepath, 
                                                                                                                                                                     self.excel_filepath, 
                                                                                                                                                                     self.output_directory, 
                                                                                                                                                                     self.output_filename))
            self.find_values_button.grid(row=3, column=1)
            
            sv_ttk.set_theme("dark")
            self.root.mainloop()
            
            
        def select_excel_file(self):
            self.excel_filepath_label.config(text=askopenfilename())
            
        
        def select_pdf_file(self):
            self.pdf_filepath_label.config(text=askopenfilename())
            
            
        def select_output_directory(self):
            self.output_directory_label.config(text=askdirectory())
            
            
        # def update_output_filename(self):
            
            
def main():
    pdf_filepath = 'sample.pdf'
    xlsx_filepath = 'sample.xlsx'
    gui = PDF_Reader_GUI()
    
if __name__ == '__main__':
    main()
