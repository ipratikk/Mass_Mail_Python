from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
from tkinter import messagebox

from setup import Setup
from mail_merge import MailMerge
from mailgui import MailGUI
from init_config import InitSetup

import threading
import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class MergeGUI:
    def __init__(self):
        root = Tk()
        root.wm_title("Mail Merge Utility")
        root.minsize(500,200)

        self.pdf_str = StringVar()
        self.excel_str = StringVar()
        self.pdf_files_str = StringVar()

        self.root = root

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.resizable(False,False)

        logger.info("Initialising PDF Input")
        self.add_pdf_inp()
        logger.info("Initialising Merge Field display")
        self.add_pdf_files_count()
        logger.info("Initialising Excel Input")
        self.add_excel_inp()
        logger.info("Initialising Excel Field display")
        self.add_excel_fields()
        logger.info("Initialising Configure Button")
        self.add_configure_btn()
        logger.info("Initialising Generate Button")
        self.add_generate_btn()
        logger.info("Initialising copyright Label")
        self.add_copyright_lbl()

        if not os.path.exists("configuration.json"):
            self.init_configure()
        
        setup = Setup()
        self.screen_height,self.screen_width = setup.get_display_size()
        
        self.root.mainloop()

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.root.destroy()
            self.root.quit()
            logger.info("Quiting Application")

    def show_alert(self,code,message):
        if code == "showinfo":
            messagebox.showinfo(code,message)
            logger.info(f"{message}")
            return
        messagebox.showerror(code,message)
        logger.error(f"{message}")
    
    def add_pdf_inp(self):
        self.pdf_lbl = Label(self.root,text = "PDF Directory")
        self.pdf_lbl.grid(row = 1, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.pdf_dir = Entry(self.root,font = "Consolas 12",textvariable = self.pdf_str,width = "30",state="readonly")
        self.pdf_dir.grid(row = 1, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
        self.pdf_browse = Button(self.root,text = "Browse",command = lambda:self.open_pdf(self.pdf_str))
        self.pdf_browse.grid(row = 1, padx = 10, pady = 10, column = 20, columnspan = 3,sticky = N+S+E+W)

    def add_pdf_files_count(self):
        self.pdf_files_lbl = Label(self.root,textvariable=self.pdf_files_str)
        
    def add_excel_inp(self):
        self.excel_lbl = Label(self.root,text = "Data file")
        self.excel_lbl.grid(row = 3, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.excel_dir = Entry(self.root,font = "Consolas 12",textvariable = self.excel_str,width = "30",state="readonly")
        self.excel_dir.grid(row = 3, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
        self.excel_browse = Button(self.root,text = "Browse",command = lambda:self.open_excel(self.excel_str))
        self.excel_browse.grid(row = 3, padx = 10, pady = 10, column = 20, columnspan = 3,sticky = N+S+E+W)

    def add_excel_fields(self):
        self.excel_header_lbl = Label(self.root,text = "Unique Identifier")
        self.excel_headers = Combobox(self.root,width="5",state = "readonly")

    def add_configure_btn(self):
        self.gen = Button(self.root,text = "Configure",command = self.init_configure)
        self.gen.grid(row = 7, padx = 10, pady = 10, column = 1, columnspan = 10,sticky = N+S+E+W)

    def add_generate_btn(self):
        self.gen = Button(self.root,text = "Start Email Client",command = lambda:self.run_script(self.pdf_str,self.excel_str))
        self.gen.grid(row = 7, padx = 10, pady = 10, column = 11, columnspan = 18,sticky = N+S+E+W)

    def add_copyright_lbl(self):
        copyright = "© 2021 Pratik Goel, Published in India"
        self.cp_lbl = Label(self.root,text = copyright)
        self.cp_lbl.grid(row = 17, padx = 10 , pady = 10, column = 5, columnspan = 12,sticky = N+S+E+W)

    def init_configure(self):
        self.root.withdraw()
        InitSetup(self.root)

    def run_script(self,pdf_dir,excel_dir):
        if len(pdf_dir.get()) < 1:
            self.show_alert("showerror","Enter PDF directory")
            return
        if len(excel_dir.get()) < 1:
            self.show_alert("showerror","Enter Data file")
            return
        self.excel_data['Files'] = self.pdf_list
        MailGUI(self.excel_data)

    def data_str(self,data):
        return ", ".join(sorted(data))

    def open_pdf(self,pdf_str):
        directory = askdirectory()
        pdf_str.set(directory)
        self.pdf_files_lbl.grid_remove()

    def open_excel(self,excel_str):
        filename = askopenfilename(filetypes = [('Microsoft Excel','*.xlsx')])
        self.excel_str.set(filename)
        if filename == "":
            self.excel_header_lbl.grid_remove()
            self.excel_headers.grid_remove()
            self.pdf_files_lbl.grid_remove()
            return
        excel_obj = MailMerge(excel_file=filename).read_excel()
        self.excel_data = excel_obj.read_data()
        
        logger.info("Displaying Excel Headers")
        self.excel_header_lbl.grid(row = 4, padx = 10 , pady = 5, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.excel_headers['values'] = tuple(self.excel_data.columns)
        self.excel_headers.grid(row = 4, column = 3,padx = 10, pady = 5,columnspan = 10,sticky = N+S+E+W)


        def find_files(eventObject):
            choice = self.excel_headers.get()
            pdf_list = []
            for id in self.excel_data[choice]:
                filename = os.path.join(self.pdf_str.get(),f"{id}.pdf")
                if os.path.exists(filename):
                    pdf_list.append(filename)
            self.pdf_files_str.set(f"Found {len(pdf_list)} files")
            self.pdf_files_lbl.grid(row = 2, column = 3,padx = 10, pady = 5,columnspan = 20,sticky = N+S+E+W)
            self.pdf_list = pdf_list

        self.excel_headers.bind("<<ComboboxSelected>>", find_files)

        
            
if __name__ == "__main__":
    MergeGUI()