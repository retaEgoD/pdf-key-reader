import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, askdirectory
import sv_ttk
import ctypes
from pdf_reader import find_and_save_values_in_pdf
import traceback

ctypes.windll.shcore.SetProcessDpiAwareness(True)

# TODO reorder excel and pdf, format and docs

class PDFReaderGUI:
        
        def __init__(self):
            """
            Initialize the GUI application.
            """
            
            self.root = tk.Tk()
            self.root.geometry('950x350')
            self.root.title('PDF Key Reader')

            self.pdf_filepath = ''
            self.excel_filepath = ''
            self.output_directory = ''
            self.output_filename = ''
            
            # Create a label for the title
            self.title_label = ttk.Label(text="PDF Key Reader", font=("Arial", 24, "bold"))
            self.title_label.pack()

            # Create a label to explain what the program does
            self.description_label = ttk.Label(text="This program reads a set of key values from an excel file and checks which ones are present in a PDF, then outputs the result as a text file in the desired directory. By default, the key column is column A and the value column is column B. Press the buttons to select your files.\n", wraplength=800, font=("Arial", 11, "italic"), justify="center")
            self.description_label.pack(pady=10)
            
            
            self.create_widgets()
            self.configure_widgets()
            self.root.mainloop()


        def create_widgets(self):
            """
            Create the GUI widgets.
            """
            # Create a frame to hold all widgets
            self.widget_frame = ttk.Frame(self.root)
            self.widget_frame.pack()

            # Pack all widgets into the frame
            self.excel_button = ttk.Button(self.widget_frame, text="Select Excel File", command=self.select_excel_file)
            self.excel_button.grid(row=2, column=0, pady=3)
            self.excel_filepath_label = ttk.Label(self.widget_frame)
            self.excel_filepath_label.grid(row=2, column=1)
            
            self.pdf_button = ttk.Button(self.widget_frame, text="Select PDF File", command=self.select_pdf_file)
            self.pdf_button.grid(row=1, column=0, pady=3)
            self.pdf_filepath_label = ttk.Label(self.widget_frame)
            self.pdf_filepath_label.grid(row=1, column=1)

            self.output_button = ttk.Button(self.widget_frame, text="Select Output Location", command=self.select_output_directory)
            self.output_button.grid(row=3, column=0, pady=3, padx=6)
            self.output_directory_label = ttk.Label(self.widget_frame)
            self.output_directory_label.grid(row=3, column=1)

            self.output_filename_label = ttk.Label(self.widget_frame, text="Choose Output Filename")
            self.output_filename_entry = ttk.Entry(self.widget_frame, font=("Arial", 14), width=60)
            self.output_filename_label.grid(row=4, column=0)
            self.output_filename_entry.grid(row=4, column=1, padx=10)


            self.confirm_frame = ttk.Frame(self.root)
            self.confirm_frame.pack(pady=10)

            self.find_values_button = ttk.Button(self.confirm_frame, text="Find Values", command=self.find_values)
            self.find_values_button.grid(row=0, column=0)
            self.confirm_label = ttk.Label(self.confirm_frame, foreground='spring green')
            self.confirm_label.grid(row=0, column=1, padx=6)


        def configure_widgets(self):
            """
            Configure the GUI layout.
            """
            self.root.grid_columnconfigure(1, weight=1)
            sv_ttk.set_theme("dark")
                        
            
        def select_excel_file(self):
            """
            Select an Excel file.
            """
            filepath = askopenfilename()
            self.excel_filepath_label.config(text=filepath)
            self.excel_filepath = filepath
            
        
        def select_pdf_file(self):
            """
            Select a PDF file.
            """
            filepath = askopenfilename()
            self.pdf_filepath_label.config(text=filepath)
            self.pdf_filepath = filepath
            
            
        def select_output_directory(self):
            """
            Select an output directory.
            """
            directory_path = askdirectory()
            self.output_directory_label.config(text=directory_path)
            self.output_directory = directory_path
            
            
        def find_values(self):
            try:
                find_and_save_values_in_pdf(self.pdf_filepath, 
                                            self.excel_filepath, 
                                            self.output_directory, 
                                            self.output_filename_entry.get(), "A", "B")
                self.confirm_label.config(text="Success.")
                
            except Exception as error:
                error_modal = tk.Toplevel(self.root)
                error_modal.title("Error")
                error_modal.geometry("750x350")
                error_header = ttk.Label(error_modal, text="The following error has occurred:", font=("Arial", 24, "bold"), foreground='orange red')
                error_header.pack()
                error_text = ttk.Label(error_modal, text=error, font=("Arial", 14, "bold"), foreground='tomato')
                error_text.pack()
                error_traceback = ttk.Label(error_modal, text="\nThe full traceback is as follows:\n\n" + traceback.format_exc())
                error_traceback.pack()
            
        
        def mainloop(self):
            self.root.mainloop()
            
    
if __name__ == '__main__':
    gui = PDFReaderGUI()
    gui.mainloop()
