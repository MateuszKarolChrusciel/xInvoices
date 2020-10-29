import os
import tkinter as tk
import tkinter.filedialog as tk_filed
import tkinter.ttk as ttk

from common import functions as funcs
from common import program_objs as obj
from common.processing_modules import process_data, process_files

BASE_PATH = os.path.dirname(os.path.realpath(__file__))
PDFS_PATH = f"{BASE_PATH}/invoices_data"

CONSOLE = obj.Console

VERSION = 0.3

PDFS = list()
XLSXS = list()


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__(None)

        self.title(f"xInvoices v{VERSION}")
        self.geometry("640x480")
        self.resizable(False, False)

        self.main_frame = MainFrame(self)


class MainFrame(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        # Settings
        self.settings = dict()

        self.settings["Wielowątkowość"] = funcs.do_nothing
        self.settings[""] = funcs.do_nothing
        self.settings[""] = funcs.do_nothing
        self.settings[""] = funcs.do_nothing

        # Bar
        self.pb_overall = None

        # Frame for PDFs
        self.pdf_frame = None
        self.pdf_listbox = None
        self.pdf_open_button = None

        # Frame for XLSXs
        self.xlsx_frame = None
        self.xlsx_listbox = None
        self.xlsx_open_button = None

        # Frame for buttons
        self.b_frame = None

        # Options buttons
        self.b_options = list()

        # Process button
        self.b_process = None

        # Initialize frame's insides
        self.init_widgets()
        self.main_menu()

        # Show frame
        self.pack(fill=tk.BOTH, expand=tk.TRUE, padx=5, pady=5)

    def init_widgets(self):
        # Bar
        self.pb_overall = ttk.Progressbar(self, mode="determinate")

        # Frame for PDFs
        self.pdf_frame = tk.LabelFrame(self, labelanchor=tk.N, text="Faktury elektroniczne (.pdf)", relief="groove")

        self.pdf_listbox = tk.Listbox(self.pdf_frame)
        self.pdf_open_button = ttk.Button(self.pdf_frame,
                                          command=lambda: self.add_to_listbox(self.pdf_listbox, "pdf"),
                                          text='Dodaj plik .pdf')

        # Frame for XLSXs
        self.xlsx_frame = tk.LabelFrame(self, labelanchor=tk.N, text="Arkusze excel (.xlsx)", relief="groove")

        self.xlsx_listbox = tk.Listbox(self.xlsx_frame)
        self.xlsx_open_button = ttk.Button(self.xlsx_frame,
                                           command=lambda: self.add_to_listbox(self.xlsx_listbox, "xlsx"),
                                           text='Dodaj plik .xlsx')

        # Frame for buttons
        self.b_frame = tk.LabelFrame(self, labelanchor=tk.N, text="Opcje", relief="groove")

        # Options buttons
        self.b_options = [ttk.Button(self.b_frame, text=f"Button{num + 1}", command=None) for num in range(10)]

        # Process button
        self.b_process = ttk.Button(self.b_frame, text="Prześlij dane z e-faktur do arkuszy excel",
                                    command=execute_main_program)

    def main_menu(self):
        # Bar
        self.pb_overall.pack(expand=tk.FALSE, fill=tk.X, side=tk.BOTTOM)

        # Frame for PDFs
        self.pdf_frame.pack(expand=tk.TRUE, fill=tk.BOTH, side=tk.LEFT)
        self.pdf_listbox.pack(expand=tk.TRUE, fill=tk.BOTH, side=tk.TOP)
        self.pdf_open_button.pack(expand=tk.FALSE, fill=tk.X, side=tk.BOTTOM)

        # Frame for XLSs
        self.xlsx_frame.pack(expand=tk.TRUE, fill=tk.BOTH, side=tk.RIGHT)
        self.xlsx_listbox.pack(expand=tk.TRUE, fill=tk.BOTH, side=tk.TOP)
        self.xlsx_open_button.pack(expand=tk.FALSE, fill=tk.X, side=tk.BOTTOM)

        # Options buttons
        self.b_frame.pack(expand=tk.TRUE, fill=tk.BOTH, side=tk.TOP)

        for num, button in enumerate(self.b_options):
            button.pack(expand=tk.FALSE, fill=tk.X, side=tk.TOP)

        # Process button
        self.b_process.pack(expand=tk.FALSE, fill=tk.X, side=tk.BOTTOM)

    @staticmethod
    def upload_files(ext):
        files = [process_files.check_ext(file, ext) for file in tk_filed.askopenfilenames()]

        if not all(files):
            raise FileNotFoundError

        if ext == "pdf" or ext == ".pdf":
            for file in files:
                global PDFS
                PDFS.append(file)

        elif ext == "xlsx" or ext == ".xlsx":
            for file in files:
                global XLSXS
                XLSXS.append(file)

        return files

    def add_to_listbox(self, listbox, ext):
        try:
            items = self.upload_files(ext)
        except FileNotFoundError:
            return
        else:
            for item in items:
                listbox.insert(tk.END, process_files.filename_from_path(item).replace(" ", ""))


def execute_main_program():
    try:
        process_data.process(PDFS, XLSXS)
    except FileNotFoundError:
        pass
    except EnvironmentError:
        pass


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
