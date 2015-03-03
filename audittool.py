from Tkinter import *
import ttk
import tkMessageBox
import tkFileDialog
from os import listdir
import subprocess
import os
import csv
import dicom
from PIL import ImageTk, Image
import time, datetime
import shutil
import sqlite3

class AuditTool(object):

    def __init__(self, master):
        """Initialize the gui and fields in x-ray and ultrasound
        files"""

        self.active_stat = BooleanVar()
        self.master = master

        self.reportType = StringVar()
        names = ['X-ray', 'Ultrasound']
        self.typeData= (
            ["1", "X-ray"],
            ["2", "Ultrasound"])

        self.xray_data_list  = ['Accession Number', 'Acquisition Date', 'Acquisition Date Time', 'Acquisition Time', 'Bits Allocated', 'Bits Stored', 'Body Part Examined', 'Burned In Annotation', 'Cassette Orientation', 'Cassette Size', 'Columns', 'Content Date', 'Content Time', 'Contrast Bolus Agent', 'Device Serial Number', 'Gantry ID', 'Grid', 'High Bit', 'Instance Creation Date', 'Instance Creation Time', 'Instance Number', 'Institution Address', 'Institution Name', 'Institutional Department Name', 'Laterality', 'Lossy Image Compression', 'Manufacturer', 'Manufacturer Model Name', 'Modality', 'Operators Name', 'Patient Birth Date', 'Patient Comments', 'Patient ID', 'Patient Name', 'Patient Sex', 'Performed Procedure Step Start Date', 'Performed Procedure Step Start Time', 'Photometric Interpretation', 'Plate ID', 'Plate Type', 'Quality Control Image', 'Referring Physician Name', 'Samples Per Pixel', 'Series Date', 'Series Description', 'Series Instance UID', 'Series Number', 'Series Time', 'Software Versions', 'Station Name', 'Study Date', 'Study ID', 'Study Time', 'View Position']
        self.xray_data_req = ['Acquisition Date', 'Acquisition Time', ' Body Part Examined', 'Content Date', 'Content Time', ' Device Serial Number', 'Instance Creation Date', 'Instance Creation Time', 'Instance Number', 'Institution Address', 'Institution Name', 'Institutional Department Name', 'Operators Name', 'Patient Birth Date', 'Patient Comments', 'Patient ID', 'Patient Name', 'Patient Sex', 'Performed Procedure Step Start Date', 'Performed Procedure Step Start Time', '  Referring Physician Name', ' Study Date', 'Study ID', 'Study Time', 'View Position']
        self.xray_data = {}      
        for data in self.xray_data_list:
            self.xray_data[data.replace(' ', '')] = 0
        
        self.ultrasound_data_list = ['Accession Number', 'Bits Allocated', 'Bits Stored', 'Columns', 'Content Date', 'Content Time', 'High Bit', 'Image Type', 'Instance Number', 'Institution Name', 'Institutional Department Name', 'Manufacturer', 'Manufacturer Model Name', 'Modality', 'Patient Birth Date', 'Patient Birth Time', 'Patient ID', 'Patient Name', 'Patient Orientation', 'Patient Sex', 'Photometric Interpretation', 'Planar Configuration', 'Referring Physician Name', 'Rows', 'Samples Per Pixel', 'Series Date', 'Series Instance UID', 'Series Number', 'Series Time', 'Software Versions', 'Station Name', 'Study Date', 'Study Description', 'Study Time']
        self.ultrasound_data_req = ['Accession Number'', ''Image Type', 'Instance Number', 'Institution Name', 'Institutional Department Name'', ''Modality', 'Patient Birth Date', 'Patient Birth Time', 'Patient ID', 'Patient Name', 'Patient Orientation', 'Patient Sex', 'Referring Physician Name', 'Name', 'Study Date', 'Study Description', 'Study Time']
        self.ultrasound_data = {}
        for data in self.ultrasound_data_list:
            self.ultrasound_data[data.replace(' ', '')] = 0

        master.title('Auditing Tool')
        master.resizable(False, False)

        self.file_opt = options = {}
        options['filetypes'] = [('CSV Files', '.csv'), ('CSV File', '.csv')]
        
        self.frame_header = ttk.Frame(master)
        self.frame_header.pack(pady = (10, 0), padx = 20)

        img = ImageTk.PhotoImage(Image.open(resource_path("appolo_logo.gif")))
        self.logo = Label(self.frame_header, image=img)
        self.logo.image = img
        self.logo.grid(row = 0, column = 0, sticky=W, pady=0, padx = 0)
        
        ttk.Label(self.frame_header, text = 'Report Type:').grid(row = 0, column = 1, padx = (10, 0), sticky = S)

        self.combobox = ttk.Combobox(self.frame_header, textvariable = self.reportType, width = 10, values = names)
        self.combobox.bind('<<ComboboxSelected>>', self.combobox_handler)
        self.combobox.grid(row = 0, column = 2, padx = (5, 20), sticky = S)

        ttk.Label(self.frame_header, text = 'Source Folder:').grid(row = 0, column = 3, sticky = S)
        self.source_entry = ttk.Entry(self.frame_header, width = 23, font = ('Arial', 10))
        self.source_entry.grid(row = 0, column = 4, padx = (5, 0), sticky = S)

        ttk.Button(self.frame_header, text = 'Browse',
                   command = self.browseSourceFolder).grid(row = 0, column = 5, padx = (5, 20), sticky = S)

        self.clinic_var = StringVar()
        ttk.Label(self.frame_header, text = 'Clinic Name:').grid(row = 0, column = 6, sticky = S)
        self.clinic_entry = ttk.Entry(self.frame_header, width = 24, textvariable=self.clinic_var, font = ('Arial', 10))
        self.clinic_entry.grid(row = 0, column = 7, padx = (5, 0), sticky = S)

        self.frame_progress = ttk.Frame(master)
        self.frame_progress.pack(fill = 'both', pady = (10, 10), padx = (20, 20))

        self.process = StringVar()
        self.process.set("")
        self.progress_label = Label(self.frame_progress, textvariable=self.process, font = ('Arial', 8))
        self.progress_label.grid(row = 0, column = 1, pady = 0, padx = (28, 20), sticky = SW)
        
        self.progress_var = StringVar()
        self.progress_var.set("Progress:")
        self.progress_label = Label(self.frame_progress, textvariable=self.progress_var)
        self.progress_label.grid(row = 1, column=0)

        self.progress = ttk.Progressbar(self.frame_progress, mode='determinate', length=586)
        self.progress.grid(row=1, sticky=E, column=1, padx = (31, 20), pady = 0)
        self.progress["value"] = 0

        ttk.Button(self.frame_progress, text = 'Export to CSV',
                   command = self.saveCSV).grid(row = 1, column = 2, padx = (15, 0))
        ttk.Button(self.frame_progress, text = 'Cancel',
                   command = self.cancelSave).grid(row = 1, column = 3, padx = (5, 0))
        
        self.frame_attributes = ttk.Frame(master)
        self.frame_attributes.pack(fill = 'both', pady = (0, 10), padx = 20)

        ttk.Label(self.frame_attributes, text = 'Select the fields:').grid(row = 0, column = 0, sticky = W, pady = 10)

        self.s_all = Variable()
        self.s_all.set('0')
        self.select_all = Checkbutton(self.frame_attributes, text="Select All", variable = self.s_all, command = self.selectAll)
        self.select_all.grid(row = 0, column = 0, sticky = W, padx = (100, 0))

        self.canvas = Canvas(self.frame_attributes, scrollregion = (0, 0, 800, 350), bg="#f1f1f1", width = 810, relief = SUNKEN)
        self.xscroll = ttk.Scrollbar(self.frame_attributes, orient = HORIZONTAL, command = self.canvas.xview)
        self.yscroll = ttk.Scrollbar(self.frame_attributes, orient = VERTICAL, command = self.canvas.yview)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.config(xscrollcommand = self.xscroll.set, yscrollcommand = self.yscroll.set)

        self.canvas.grid(row = 1, column = 0, padx = (20, 20), pady=(0, 20))
        self.yscroll.grid(row = 1, column = 1, sticky = NS)

        self.frameC = Frame(self.canvas)
        self.canvas.create_window(0,0, anchor = NW, window = self.frameC)

        self.frame_footer = ttk.Frame(master)
        self.frame_footer.pack(fill = 'both', pady = (0, 10), padx = 20)
        
        ttk.Button(self.frame_footer, text = 'Exit',
                   command = master.destroy).grid(row = 0, column = 0, padx = (5, 0), sticky = W)

        img = ImageTk.PhotoImage(Image.open(resource_path("logo.gif")))
        self.powered_by = Label(self.frame_footer, image=img)
        self.powered_by.image = img
        self.powered_by.grid(row = 0, column = 1, columnspan =5, sticky=E, pady=(10,0), padx = (625, 0))

        self.set_clinic_name("12081626")
        
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(-1*(event.delta/120), "units")

    def set_clinic_name(self, name):
        conn = sqlite3.connect('clinic.db')
        cur = conn.cursor()
        cur.execute('''CREATE TABLE IF NOT EXISTS clinic (name text)''')
        cur.execute("SELECT Count(*), name FROM clinic")
        rows = cur.fetchall()
        for row in rows:
            n_rows = row[0]
            clinic_name = row[1]
        
        if n_rows == 0:
            if name != '12081626':
                cur.execute("""INSERT INTO clinic (name) VALUES (?)""", (name,))
        else:
            self.clinic_var.set(clinic_name)
        conn.commit()
        conn.close()

    def combobox_handler(self, event):
        """ Populates the canvas area with parameters that are assigned to the selected report type """
        self.clearProgressbar()
        item = self.combobox.current()
        if self.typeData[item][0] == "1":
            i =0;
            self.frameC.destroy()
            self.frameC = Frame(self.canvas, bg = "#ffffff")
            self.canvas.create_window(0,0, anchor = 'nw', window = self.frameC)
            for data in self.xray_data_list:
                self.checkBoxName = data
                dat = data.replace(' ', '')
                self.xray_data[dat] = Variable()
                if data in self.xray_data_req:
                    self.xray_data[dat].set('1')
                else:
                    self.xray_data[dat].set('0')
                self.c = Checkbutton(self.frameC, text=self.checkBoxName, variable = self.xray_data[dat], bg = "#ffffff")
                self.c.grid(row = i/4, column = i % 4, sticky = 'w')
                i += 1
        elif self.typeData[item][0] == "2":
            i =0;
            self.frameC.destroy()
            self.frameC = Frame(self.canvas, bg = "#ffffff")
            self.canvas.create_window(0,0, anchor = 'nw', window = self.frameC)
            for data in self.ultrasound_data_list:
                self.checkBoxName = data
                dat = data.replace(' ', '')
                self.ultrasound_data[dat] = Variable()
                if data in self.ultrasound_data_req:
                    self.ultrasound_data[dat].set('1')
                else:
                    self.ultrasound_data[dat].set('0')
                self.c = Checkbutton(self.frameC, text=self.checkBoxName, variable = self.ultrasound_data[dat], bg = "#ffffff")
                self.c.grid(row = i/5, column = i % 5, sticky = 'w')
                i += 1

    def browseSourceFolder(self):
        """ Creates a dialog box for selecting a directory and adds the selected path to the source_entry field in gui """
        directory = tkFileDialog.askdirectory()
        self.clearProgressbar()
        if listdir(directory) == []:
            tkMessageBox.showwarning("Empty", "Selected folder is empty. Choose another one.")
        else:
            files = listdir(directory)
            dicom = False

        self.source_entry.delete(0, END)
        self.source_entry.insert(0, directory)


    def saveCSV(self):
        """ Creates and saves the csv file """

        clinic_name = self.clinic_entry.get()
        if clinic_name == '':
            tkMessageBox.showwarning("Empty", "Please enter the clinic name.")
            return 0
        
        self.set_clinic_name(clinic_name)

        if self.active_stat.get() == True:
            tkMessageBox.showinfo("Ok", "One task is already running.")
            return
        
        source = self.source_entry.get()
        if source == '':
            tkMessageBox.showwarning("Empty", "You haven't selected any folder.")
            return 0
        files = listdir(source)

        if len(files) == 0:
            tkMessageBox.showwarning("Empty", "The selected folder is either empty or it has no dicom files.")
            return 0

        file_name = self.clinic_entry.get() + time.strftime("(%d_%m_%Y)") + time.strftime("(%H_%M_%S)") + '.csv' #+ time.strftime("%c")+ '.csv'
        
        csvfile = open(file_name, 'wb')
        wr = csv.writer(csvfile, delimiter=',', quotechar=',', escapechar = '\\', quoting=csv.QUOTE_NONE)

        item = self.combobox.current()

        self.count = 0
        self.active_stat.set(True)

        csvheader = []
        
        #write the file headers
        if self.typeData[item][0] == "1":
            for data in self.xray_data:
                if self.xray_data[data].get() == '1':
                    csvheader.append(data)
            wr.writerow(csvheader)

        elif self.typeData[item][0] == "2":
            for data in self.ultrasound_data:
                if self.ultrasound_data[data].get() == '1':
                    csvheader.append(data)
            wr.writerow(csvheader)

        #write the values
        if self.typeData[item][0] == "1":
            n_files = 0
            files = []
            for root, dirs, fils in os.walk(source):
                if len(fils) > 0:
                    for f in fils:
                        files.append(os.path.join(root, f))
            n_files = len(files)
            self.progress["maximum"] = n_files
            for dicom_file in files:
                if self.active_stat.get():
                    try:
                        d_file = dicom.read_file(dicom_file)
                    except:
                        n_files -= 1
                        self.progress["maximum"] = n_files
                        continue
                    csvlist = []
                
                    for data in self.xray_data:
                        if self.xray_data[data].get() == '1':
                            if data in d_file.dir():
                                value = str(getattr(d_file, data, ' '))
                                #print data, value
                                value = value.replace(',', ' ')
                                value = value.replace('\r\n', ' ')
                                csvlist.append(value)
                            else:
                                csvlist.append(' ')
                    wr.writerow(csvlist)
                    self.count += 1
                    self.process.set("Processing " + str(self.count) + " of " + str(n_files) + " files.")
                    self.progress["value"] = self.count
                    self.progress.update()
                else:
                    self.clearProgressbar()
                    tkMessageBox.showinfo("Cancel", "Operation cancelled.")
                    break

        elif self.typeData[item][0] == "2":
            n_files = 0
            files = []
            for root, dirs, fils in os.walk(source):
                if len(fils) > 0:
                    files.append(os.path.join(root, fils[0]))
            n_files = len(files)
            self.progress["maximum"] = n_files
            
            for dicom_file in files:
                if self.active_stat.get():
                    try:
                        d_file = dicom.read_file(dicom_file)
                    except:
                        n_files -= 1
                        self.progress["maximum"] = n_files
                        continue
                    csvlist = []

                    for data in self.ultrasound_data:
                        if self.ultrasound_data[data].get() == '1':
                            if data in d_file.dir():
                                value = str(getattr(d_file,  data, ''))
                                value = value.replace(',', ' ')
                                csvlist.append(value)
                            else:
                                csvlist.append(' ')
                    wr.writerow(csvlist)

                    self.count += 1
                    self.process.set("Processing " + str(self.count) + " of " + str(n_files) + " files.")
                    self.progress["value"] = self.count
                    self.progress.update()
                else:
                    self.clearProgressbar()
                    tkMessageBox.showinfo("Cancel", "Operation cancelled.")
                    break               

        self.active_stat.set(False)    
        self.process.set("Processed " + str(self.count) + " of " + str(n_files) + " files.")

        csvfile.close()

        self.sendFile(file_name)
        
    def cancelSave(self):
        """ Cancels the ongoing auditing process """
        self.active_stat.set(False)

    def clearProgressbar(self):
        """sets the progressbar to the initial position"""
        self.process.set("")
        self.progress["value"] = 0
        self.progress.update()

    def selectAll(self):
        """for selecting all the options in comboboxes"""
        item = self.combobox.current()
        if self.s_all.get() == "1":
            if self.typeData[item][0] == "1" and item >= 0:
                for data in self.xray_data:
                    self.xray_data[data].set('1')
            elif self.typeData[item][0] == "2" and item >= 0:
                for data in self.ultrasound_data:
                    self.ultrasound_data[data].set('1')
            else:
                tkMessageBox.showwarning("Empty Report Type", "Select one report type.")
        else:
            if self.typeData[item][0] == "1" and item >= 0:
                for data in self.xray_data:
                    self.xray_data[data].set('0')
            elif self.typeData[item][0] == "2" and item >= 0:
                for data in self.ultrasound_data:
                    self.ultrasound_data[data].set('0')
            else:
                tkMessageBox.showwarning("Empty Report Type", "Select one report type.")
        

    def close_window(self):
        self.destroy()

    def sendFile(self, file_to_send):
        """sends the file to the shared folder on the network"""
        try:
            shutil.copy(file_to_send, '\\\\10.42.52.39\\Clinic-Audit')
            tkMessageBox.showinfo("Report Generated", str(self.count) + " records exported.")
        except:
            tkMessageBox.showwarning("File not sent!", "Not able to send the file. Make sure your system is connected to the local network and try again.")
        finally:
            os.remove(file_to_send)
        

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = getattr(sys, '_MEIPASS', os.getcwd())
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path) 
      
def main():            
    
    root = Tk()
        
    audit = AuditTool(root)
    
    if "nt" == os.name:
        root.wm_iconbitmap(bitmap = resource_path("appolo.ico"))
    root.mainloop()
    
if __name__ == "__main__": main()
