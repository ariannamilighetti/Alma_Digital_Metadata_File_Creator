import create_alma_xml
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


class From_Spreadsheet (ttk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.selection_frm = ttk.LabelFrame(self, text="Items")
        self.selection_frm.grid(row=0, column=0, sticky='nwse', padx=5, pady=5)
        self.spreadsheeet_field()
        self.folder_fileds()
        self.run_button()
        self.run_frame()
        self.run_xml_creation()
        self.columnconfigure(0, weight = 1)
        self.rowconfigure(0, weight =0)
        self.selection_frm.columnconfigure(0, weight = 1)
        self.selection_frm.columnconfigure(1, weight = 0)
        self.selection_frm.columnconfigure(2, weight = 2)

    def spreadsheeet_field(self):
        # Spreadsheet Fields
        spreadsheet_label = tk.Label(master=self.selection_frm, text="Import metadata spreadsheet", anchor='nw')
        spreadsheet_button = tk.Button(self.selection_frm, text='Open', command=self.upload_spreadsheet,bg='#ff8962')
        self.spreadsheet_filename = tk.Entry(master=self.selection_frm)
        spreadsheet_label.grid(column=0, row=1, columnspan=1, sticky="ew", padx=5, pady=5)
        spreadsheet_button.grid(column=1, row=1, columnspan=1, sticky="ew", padx=5, pady=5)
        self.spreadsheet_filename.grid(column=2, row=1, columnspan=3, sticky="ew", padx=5, pady=5)
    
    def folder_fileds(self):
        # Folder Fields
        dir_label = tk.Label(master=self.selection_frm, text="Select item(s) parent folder", anchor='nw')
        dir_button = tk.Button(self.selection_frm, text='Open', command=self.askDirectory_ss, bg='#ff8962')
        self.ss_directory_label = tk.Entry(master=self.selection_frm)
        dir_label.grid(column=0, row=2, columnspan=1, sticky="ew", padx=5, pady=5)
        dir_button.grid(column=1, row=2, columnspan=1, sticky="ew", padx=5, pady=5)
        self.ss_directory_label.grid(column=2, row=2, columnspan=3, sticky="ew", padx=5, pady=5)

    def run_button(self):
        self.run_button = tk.Button(self, text="Run", command=lambda: self.get_metadata(), bg='#7fd1ae')
        self.run_button.grid(row=1, column=0, sticky='nwse', padx=5, pady=5)
    
    def get_metadata(self):
        self.metadata = create_alma_xml.create_check_window(self.spreadsheet_filename.get(), self.ss_directory_label.get())
        self.display_metadata()
    
    def run_frame(self):
        # Running frame
        self.run_frame = ttk.LabelFrame(self, text="Check Metadata")
        run_shmark = tk.Label(master=self.run_frame, text="Barcode", anchor='nw')
        run_verification = tk.Label(master=self.run_frame, text="Title", anchor='nw')
        run_status = tk.Label(master=self.run_frame, text="MMS ID", anchor='nw')
        run_leader = tk.Label(master=self.run_frame, text="Leader", anchor='nw')
        run_lable = tk.Label(master=self.run_frame, text="Label", anchor='nw')
        run_access_r = tk.Label(master=self.run_frame, text="Access Rights", anchor='nw')

        self.run_frame.grid(row=2, column=0, sticky='nwse', padx=5, pady=5)
        run_shmark.grid(column=0, row=0, sticky="ew", padx=20, pady=5)
        run_verification.grid(column=1, row=0, sticky="ew", padx=20, pady=5)
        run_status.grid(column=2, row=0, sticky="ew", padx=20, pady=5)
        run_leader.grid(column=3, row=0, sticky="ew", padx=20, pady=5)
        run_lable.grid(column=4, row=0, sticky="ew", padx=20, pady=5)
        run_access_r.grid(column=5, row=0, sticky="ew", padx=20, pady=5)

    def display_metadata(self):
        values = ['Restricted', 'Unrestricted']
        self.metadata = create_alma_xml.get_metadata()
        self.label_dic = dict()
        row = 1
        for n in self.metadata.values():
            
            for m, value in enumerate(n):
                if m != len(n)-1 :
                    met_label = tk.Label(master = self.run_frame, text = value, anchor = 'nw', width = 20, justify= 'left')
                    met_label.grid(column=m , row=row, sticky="nw", padx=20, pady=5)
                else:
                    label_name = "E{0}".format(n)
                    button_name = "B{0}".format(n)
                    acc_name = "A{0}".format(n)
                    var = tk.StringVar(value=value) 
                    self.label_dic[label_name] = tk.Entry(self.run_frame, textvariable=var, width = 30)
                    self.label_dic[label_name].grid(column=4, row=row, sticky="ew", padx=20, pady=5)
                    
            self.label_dic[acc_name] = ttk.Combobox(self.run_frame, values = values)
            self.label_dic[acc_name].bind("<MouseWheel>", self.empty_scroll_command)
            self.label_dic[acc_name].grid(column=5, row=row, sticky="ew", padx=20, pady=5)
            self.label_dic [button_name] = tk.Button(self.run_frame, text="update", command=self.get_on_click(self.label_dic, label_name, acc_name))
            self.label_dic [button_name].grid(column=6, row=row, sticky="ew", padx=20, pady=5)
            row += 1
    
    def empty_scroll_command(self, event):
        return "break"
        
    def get_on_click(self, widget_dict, entry_name, access_name):
        def on_click():
            result_label = widget_dict[entry_name].get()
            index = int(entry_name[1:])
            self.metadata[index][4] = result_label
            result_access = widget_dict[access_name].get()
            index = int(access_name[1:])
            self.metadata[index].append(result_access)
        return on_click

    def run_xml_creation(self):
        self.xml_button = tk.Button(self, text="Create xml", command=lambda: create_alma_xml.create_xml(self.metadata, self.ss_directory_label.get(), self.xml_button), bg='#ff8962')
        self.xml_button.grid(row=3, column=0, sticky='nwse', padx=5, pady=5)

    def upload_spreadsheet (self,event=None):
        global spreadsheet_input
        spreadsheet_input = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("Any file", "*")))
        self.spreadsheet_filename.delete(0, "end")
        self.spreadsheet_filename.insert(0, spreadsheet_input)

    def askDirectory_ss(self, event=None):
        global end_directory
        end_directory = filedialog.askdirectory()
        self.ss_directory_label.delete(0, "end")
        self.ss_directory_label.insert(0, end_directory)

    def askDirectory_man(self, event=None):
        global end_directory
        end_directory = filedialog.askdirectory()
        self.man_directory_label.delete(0, "end")
        self.man_directory_label.insert(0, end_directory)
    

# to test this class
def __main__():
    app = tk.Tk()
    data_input = From_Spreadsheet()
    data_input.grid(row=2, column = 0, sticky="nwse")
    tk.mainloop()

if __name__ == '__main__':
    __main__()