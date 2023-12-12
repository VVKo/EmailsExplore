import json
from tkinter import *
from tkinter.ttk import *
import pandas as pd
import os, platform, pickle
from tkinter import filedialog, messagebox, simpledialog

from pandastable import Table, images, dialogs, util

from collections import OrderedDict

class EmailsExplore(Frame):
    
    def __init__(self, parent=None, data=None):
        self.parent = parent
        if not self.parent:
            Frame.__init__(self)
            self.main = self.master
        else:
            self.main=Toplevel()
            self.master = self.main

        self.currplatform = platform.system()
        self.setConfigDir()
        self.defaultsavedir = os.path.join(os.path.expanduser('~'))
        self.loadAppOptions()
        # start logging
        # self.start_logging()


        self.main.title('EmailsExplore')
        self.createMenuBar()
        self.setupGUI()
        self.setStyles()
        self.clipboarddf = None
        self.projopen = False

        self.newProject()

        self.main.lift()
        return

    # def start_logging(self):
    #     import logging
    #     from .core import logfile
    #     logging.basicConfig(filename=logfile, format='%(asctime)s %(message)s')
    def setStyles(self):
        """Set theme and widget styles"""

        style = self.style = Style(self)
        available_themes = self.style.theme_names()
        plf = util.checkOS()
        if plf == 'linux':
            style.theme_use('default')
        elif plf == 'darwin':
            style.theme_use('clam')

        self.bg = bg = self.style.lookup('TLabel.label', 'background')
        style.configure('Horizontal.TScale', background=bg)
        # set common background style for all widgets because of color issues
        # if plf in ['linux','darwin']:
        #    self.option_add("*background", bg)
        dialogs.applyStyle(self.menu)
        return

    def setupGUI(self):
        """Add all GUI elements"""

        self.m = PanedWindow(self.main, orient=HORIZONTAL)
        self.m.pack(fill=BOTH, expand=1)
        self.nb = Notebook(self.main)
        self.m.add(self.nb)
        self.setGeometry()
        return

    def getBestGeometry(self):
        """Calculate optimal geometry from screen size"""

        ws = self.main.winfo_screenwidth()
        hs = self.main.winfo_screenheight()
        if ws < 1400:
            g = '%dx%d+%d+%d' % (ws, hs, 0, 0)
            self.w = ws
        else:
            self.w = w = ws / 1.3;
            h = hs * 0.7
            x = (ws / 2) - (w / 2);
            y = (hs / 2) - (h / 2)
            g = '%dx%d+%d+%d' % (w, h, x, y)
        return g
    def setGeometry(self):
        self.winsize = self.getBestGeometry()
        self.main.geometry(self.winsize)
        return
    def createMenuBar(self):
        self.menu = Menu(self.main)
        file_menu = Menu(self.menu,tearoff=0)
        #add recent first
        self.createRecentMenu(file_menu)
        filemenuitems = {'01New Project':{'cmd': self.newProject},
                    '02Open Project':{'cmd': lambda: self.loadProject(asksave=True)},
                    '03Close':{'cmd':self.closeProject},
                    '04Save':{'cmd':self.saveProject},
                    '05Save As':{'cmd':self.saveasProject},
                    '06sep':'',
                    '07Import CSV':{'cmd':self.importCSV},
                    '08Import HDF5':{'cmd':self.importHDF},
                    '09Import from URL':{'cmd':self.importURL},
                    '10Import Excel':{'cmd':self.importExcel},
                    '11Import JSON': {'cmd': self.importJSON},
                    '12Export CSV':{'cmd':self.exportCSV},
                    '13sep':'',
                    '14Quit':{'cmd':self.quit}}

        self.file_menu = self.createPulldown(self.menu, filemenuitems, var=file_menu)
        self.menu.add_cascade(label='File', menu=self.file_menu['var'])
        editmenuitems = {'01Undo Last Change': {'cmd': self.undo},
                         '02Copy Table': {'cmd': self.copyTable},
                         '03Find/Replace': {'cmd': self.findText},
                         '04Preferences': {'cmd': self.currentTablePrefs}
                         }
        self.edit_menu = self.createPulldown(self.menu, editmenuitems)
        self.menu.add_cascade(label='Edit', menu=self.edit_menu['var'])

        self.sheet_menu = {'01Add Sheet': {'cmd': lambda: self.addSheet(select=True)},
                           '02Remove Sheet': {'cmd': lambda: self.deleteSheet(ask=True)},
                           '03Copy Sheet': {'cmd': self.copySheet},
                           '04Rename Sheet': {'cmd': self.renameSheet},
                           # '05Sheet Description':{'cmd':self.editSheetDescription}
                           }
        self.sheet_menu = self.createPulldown(self.menu, self.sheet_menu)
        self.menu.add_cascade(label='Sheet', menu=self.sheet_menu['var'])

        self.view_menu = {'01Zoom In': {'cmd': lambda: self._call('zoomIn')},
                          '02Zoom Out': {'cmd': lambda: self._call('zoomOut')},
                          '03Wrap Columns': {'cmd': lambda: self._call('setWrap')},
                          '04sep': '',
                          '05Dark Theme': {'cmd': lambda: self._call('setTheme', name='dark')},
                          '06Bold Theme': {'cmd': lambda: self._call('setTheme', name='bold')},
                          '07Default Theme': {'cmd': lambda: self._call('setTheme', name='default')},
                          }
        self.view_menu = self.createPulldown(self.menu, self.view_menu)
        self.menu.add_cascade(label='View', menu=self.view_menu['var'])

        self.table_menu = {'01Describe Table': {'cmd': self.describe},
                           '02Convert Column Names': {'cmd': lambda: self._call('convertColumnNames')},
                           '03Convert Numeric': {'cmd': lambda: self._call('convertNumeric')},
                           '04Clean Data': {'cmd': lambda: self._call('cleanData')},
                           '05Find Duplicates': {'cmd': lambda: self._call('findDuplicates')},
                           '06Correlation Matrix': {'cmd': lambda: self._call('corrMatrix')},
                           '07Concatenate Tables': {'cmd': self.concat},
                           '08Table to Text': {'cmd': lambda: self._call('showasText')},
                           '09Table Info': {'cmd': lambda: self._call('showInfo')},
                           '10sep': '',
                           '11Transform Values': {'cmd': lambda: self._call('transform')},
                           '12Group-Aggregate': {'cmd': lambda: self._call('aggregate')},
                           '13Cross Tabulation': {'cmd': lambda: self._call('crosstab')},
                           '14Merge/Concat Tables': {'cmd': lambda: self._call('doCombine')},
                           '15Pivot Table': {'cmd': lambda: self._call('pivot')},
                           '16Melt Table': {'cmd': lambda: self._call('melt')},
                           '17Time Series Resampling': {'cmd': lambda: self._call('resample')}
                           }
        self.table_menu = self.createPulldown(self.menu, self.table_menu)
        self.menu.add_cascade(label='Tools', menu=self.table_menu['var'])

        self.dataset_menu = {'01Sample Data': {'cmd': self.sampleData},
                             '03Iris Data': {'cmd': lambda: self.getData('iris.csv')},
                             '03Tips Data': {'cmd': lambda: self.getData('tips.csv')},
                             '04Stacked Data': {'cmd': self.getStackedData},
                             '05Pima Diabetes':
                                 {'cmd': lambda: self.getData('pima.csv')},
                             '06Titanic':
                                 {'cmd': lambda: self.getData('titanic3.csv')},
                             '07miRNA expression':
                                 {'cmd': lambda: self.getData('miRNA.csv')},
                             '08CO2 time series':
                                 {'cmd': lambda: self.getData('co2-ppm-mauna-loa.csv')},
                             '09Zoo Dataset':
                                 {'cmd': lambda: self.getData('zoo_dataset.csv')},
                             }
        self.dataset_menu = self.createPulldown(self.menu, self.dataset_menu)
        self.menu.add_cascade(label='Datasets', menu=self.dataset_menu['var'])

        self.plots_menu = {'01Store plot': {'cmd': self.addPlot},
                           '02Clear plots': {'cmd': self.updatePlotsMenu},
                           '03PDF report': {'cmd': self.pdfReport},
                           '04sep': ''}
        self.plots_menu = self.createPulldown(self.menu, self.plots_menu)
        self.menu.add_cascade(label='Plots', menu=self.plots_menu['var'])

        self.plugin_menu = {'01Update Plugins': {'cmd': self.discoverPlugins},
                            '02Install Plugin': {'cmd': self.installPlugin},
                            '03sep': ''}
        self.plugin_menu = self.createPulldown(self.menu, self.plugin_menu)
        self.menu.add_cascade(label='Plugins', menu=self.plugin_menu['var'])

        self.work_menu = {'01Завантажити співробіткиків': {'cmd': self.load_staff},
                          '02Завантажити корпоративки': {'cmd': self.load_emails},
                          '03sep':'',
                          '04Співробітник->Пошта':{'cmd': self.staff_email}}
        self.work_menu = self.createPulldown(self.menu, self.work_menu)
        self.menu.add_cascade(label='РОБОЧА', menu=self.work_menu['var'])

        self.help_menu = {'01Online Help': {'cmd': self.online_documentation},
                          '02View Error Log': {'cmd': self.showErrorLog},
                          '03About': {'cmd': self.about}}
        self.help_menu = self.createPulldown(self.menu, self.help_menu)
        self.menu.add_cascade(label='Help', menu=self.help_menu['var'])


        self.main.config(menu=self.menu)
        return

    def staff_email(self):
        staff = self.staff.copy()
        staff['Last Name [Required]'] = staff['Last Name [Required]'].str.strip().str.replace("'",'`').str.replace("ʼ",'`')
        staff['First Name [Required]'] = (staff['first name'].str.replace("'",'`').str.replace("ʼ",'`') + ' ' + staff["second name"].fillna('').str.replace("'",'`').str.replace("ʼ",'`')).str.strip()
        all_emails = self.all_emails.copy()
        all_emails['First Name [Required]'] = all_emails['First Name [Required]'].str.strip().str.replace("'",'`').str.replace("ʼ",'`')
        all_emails['Last Name [Required]'] = all_emails['Last Name [Required]'].str.strip().str.replace("'",'`').str.replace("ʼ",'`')
        staff_emails_in = pd.merge(staff, all_emails, on=['First Name [Required]', "Last Name [Required]"], how='inner')
        staff_emails_out = pd.merge(staff, all_emails, on=['First Name [Required]', "Last Name [Required]"], how='outer', indicator=True)
        not_in_emails = staff_emails_out[staff_emails_out['_merge'] == 'left_only'][self.staff.columns]

        self.addSheet('мають корпоративки', df=staff_emails_in, select=True)
        self.addSheet('не мають корпоративок', df=not_in_emails, select=True)

        return

    def load_emails(self, filename=None):
        if filename is None:
            filename = filedialog.askopenfilename(parent=self.master,
                                                  defaultextension='.json',
                                                  initialdir=os.getcwd(),
                                                  filetypes=[("json", "*.json")])
        with open(filename) as f:
            jsondata = json.load(f)

        self.all_emails = pd.json_normalize(jsondata['users'])

        self.addSheet('all_emails', df=self.all_emails, select=True)
        return
    def load_staff(self, filename='data/staff.xlsx'):
        if filename is None:
            filename = filedialog.askopenfilename(parent=self.master,
                                                  defaultextension='.xlsx',
                                                  initialdir=os.getcwd(),
                                                  filetypes=[("xls", "*.xls"),
                                                             ("xlsx", "*.xlsx"),
                                                             ("All files", "*.*")])

        self.staff = pd.read_excel(filename, sheet_name='staff')
        self.addSheet('staff', df=self.staff, select=True)
        return
    def describe(self):
        pass

    def addPlot(self):
        pass

    def updatePlotsMenu(self, clear=True):
        pass

    def online_documentation(self, event=None):
        pass

    def showErrorLog(self):
        pass

    def about(self):
        abwin = Toplevel()
        x, y, w, h = dialogs.getParentGeometry(self.main)
        abwin.geometry('+%d+%d' % (x + w / 2 - 200, y + h / 2 - 200))
        abwin.title('About')
        abwin.transient(self)
        abwin.grab_set()
        abwin.resizable(width=False, height=False)
        abwin.configure(background=self.bg)
        logo = images.tableapp_logo()
        label = Label(abwin, image=logo, anchor=CENTER)
        label.image = logo
        label.grid(row=0, column=0, sticky='ew', padx=4, pady=4)
        style = Style()
        style.configure("BW.TLabel", font='arial 11')
        text = 'EmailsExplore Application\n' \
               + 'version ' + '0.0.1' + '\n' \
               + 'Copyright (C) VVKo\n' \
               + 'This program is free software; you can redistribute it and/or\n' \
               + 'modify it under the terms of the GNU General Public License\n' \
               + 'as published by the Free Software Foundation; either version 3\n' \
               + 'of the License, or (at your option) any later version.\n' \

        row = 1
        # for line in text:
        tmp = Label(abwin, text=text, style="BW.TLabel")
        tmp.grid(row=row, column=0, sticky='news', pady=2, padx=4)

        return

    def discoverPlugins(self):
        pass

    def installPlugin(self):
        pass

    def pdfReport(self):
        pass

    def sampleData(self):
        pass

    def getData(self, name):
        pass

    def getStackedData(self):
        pass

    def concat(self):
        pass
    def _call(self, func, **args):
        pass
    def undo(self):
        pass

    def findText(self):
        pass

    def addSheet(self, sheetname=None, df=None, meta=None, select=False):
        """Add a sheet with new or existing data"""

        names = [self.nb.tab(i, "text") for i in self.nb.tabs()]

        def checkName(name):
            if name == '':
                messagebox.showwarning("Whoops", "Name should not be blank.")
                return 0
            if name in names:
                messagebox.showwarning("Name exists", "Sheet name already exists!")
                return 0

        noshts = len(self.nb.tabs())
        if sheetname == None:
            sheetname = simpledialog.askstring("New sheet name?", "Enter sheet name:",
                                               initialvalue='sheet' + str(noshts + 1))
        if sheetname == None:
            return
        if checkName(sheetname) == 0:
            return
        # Create the table
        main = PanedWindow(orient=HORIZONTAL)
        self.sheetframes[sheetname] = main
        self.nb.add(main, text=sheetname)
        f1 = Frame(main)
        table = Table(f1, dataframe=df, showtoolbar=1, showstatusbar=1)
        f2 = Frame(main)
        # show the plot frame
        pf = table.showPlotViewer(f2)
        # load meta data
        if meta != None:
            self.loadMeta(table, meta)
        # add table last so we have save options loaded already
        main.add(f1, weight=3)
        table.show()
        main.add(f2, weight=4)

        if table.plotted == 'main':
            table.plotSelected()
        elif table.plotted == 'child' and table.child != None:
            table.child.plotSelected()
        self.saved = 0
        self.currenttable = table
        # attach menu state of undo item so that it's disabled after an undo
        # table.undo_callback = lambda: self.toggleUndoMenu('active')
        self.sheets[sheetname] = table

        if select == True:
            ind = self.nb.index('end') - 1
            s = self.nb.tabs()[ind]
            self.nb.select(s)
        return sheetname

    def deleteSheet(self, ask=False):
        """Delete a sheet"""

        s = self.nb.index(self.nb.select())
        name = self.nb.tab(s, 'text')
        w = True
        if ask == True:
            w = messagebox.askyesno("Delete Sheet",
                                    "Remove this sheet?",
                                    parent=self.master)
        if w == False:
            return
        self.nb.forget(s)
        del self.sheets[name]
        del self.sheetframes[name]
        return

    def copySheet(self, newname=None):
        pass

    def renameSheet(self):
        pass
    def currentTablePrefs(self):
        pass
    def copyTable(self, subtable=False):
        pass
    def exportCSV(self):
        pass

    def importExcel(self, filename=None):
        if filename is None:
            filename = filedialog.askopenfilename(parent=self.master,
                                                  defaultextension='.xlsx',
                                                  initialdir=os.getcwd(),
                                                  filetypes=[("xls", "*.xls"),
                                                             ("xlsx", "*.xlsx"),
                                                             ("All files", "*.*")])

        data = pd.read_excel(filename, sheet_name=None)
        for n in data:
            self.addSheet(n, df=data[n], select=True)
        return

    def importJSON(self, filename=None):
        if filename is None:
            filename = filedialog.askopenfilename(parent=self.master,
                                                  defaultextension='.json',
                                                  initialdir=os.getcwd(),
                                                  filetypes=[("json", "*.json")])
        with open(filename) as f:
            jsondata = json.load(f)

        data = pd.json_normalize(jsondata['users'])

        self.addSheet('all_emails', df=data, select=True)

        unique_paths = data["Org Unit Path [Required]"].unique().tolist()

        # Count the number of emails for each unique path
        counts = data["Org Unit Path [Required]"].value_counts().reset_index()
        counts.columns = ['Org Unit Path [Required]', 'Кількість пошт']

        # Save the unique paths with their respective counts to the 'Підрозділи' sheet
        unique_paths_df = pd.DataFrame({'Org Unit Path [Required]': unique_paths})
        unique_paths_with_counts = pd.merge(unique_paths_df, counts, on='Org Unit Path [Required]', how='left')
        self.addSheet('Структурні підрозділи', df=unique_paths_with_counts, select=True)

        filtered_df = unique_paths_with_counts[unique_paths_with_counts['Org Unit Path [Required]'].str.startswith('/СПІВРОБІТНИКИ/')]
        self.addSheet('СПІВРОБІТНИКИ', df=filtered_df, select=True)

        return

    def getCurrentTable(self):

        s = self.nb.index(self.nb.select())
        name = self.nb.tab(s, 'text')
        table = self.sheets[name]
        return table
    def importURL(self):
        pass

    def importHDF(self):
        pass

    def importCSV(self):
        """Import json to a new sheet"""

        self.addSheet(select=True)
        table = self.getCurrentTable()
        table.importCSV(dialog=True)
        return

    def saveasProject(self):
        pass

    def saveProject(self, filename=None):
        pass

    def closeProject(self):
        """Close"""

        if self.projopen == False:
            w = False
        else:
            w = messagebox.askyesnocancel("Close Project",
                                          "Save this project?",
                                          parent=self.master)
        if w == None:
            return
        elif w == True:
            self.saveProject()
        else:
            pass
        for n in self.nb.tabs():
            self.nb.forget(n)
        self.filename = None
        self.projopen = False
        self.main.title('EmailsExplore')
        return w
    
    def newProject(self, data=None, df=None):
        """Create a new project from data or empty"""

        w = self.closeProject()
        if w == None:
            return
        self.sheets = OrderedDict()
        self.sheetframes = {}  # store references to enclosing widgets
        self.openplugins = {}  # refs to running plugins
        self.updatePlotsMenu()
        for n in self.nb.tabs():
            self.nb.forget(n)
        if data != None:
            for s in sorted(data.keys()):
                if s == 'meta':
                    continue
                df = data[s]['table']
                if 'meta' in data[s]:
                    meta = data[s]['meta']
                else:
                    meta = None
                # try:
                self.addSheet(s, df, meta)
                '''except Exception as e:
                    print ('error reading in options?')
                    print (e)'''
        else:
            self.addSheet('sheet1')
        self.filename = None
        self.projopen = True
        self.main.title('EmailsExplore')
        return
    
    def loadProject(self, filename=None, asksave=False):
        pass

    def createRecentMenu(self, menu):
        """Recent projects menu"""

        from functools import partial
        recent = self.appoptions['recent']
        recentmenu = Menu(menu)
        menu.add_cascade(label="Open Recent", menu=recentmenu)
        for r in recent:
            recentmenu.add_command(label=r, command=partial(self.loadProject, r))
        return

    def setConfigDir(self):
        """Set up config folder"""

        homepath = os.path.join(os.path.expanduser('~'))
        path = '.emailsexplore'
        self.configpath = os.path.join(homepath, path)
        self.pluginpath = os.path.join(self.configpath, 'plugins')
        if not os.path.exists(self.configpath):
            os.mkdir(self.configpath)
            os.makedirs(self.pluginpath)
        return

    def loadAppOptions(self):
        """Load global app options if present"""

        appfile = os.path.join(self.configpath, 'app.p')
        if os.path.exists(appfile):
            self.appoptions = pickle.load(open(appfile,'rb'))
        else:
            self.appoptions = {}
            self.appoptions['recent'] = []
        return


     
    def createPulldown(self, menu, dict, var=None):
        """Create pulldown menu, returns a dict.
        Args:
            menu: parent menu bar
            dict: dictionary of the form -
            {'01item name':{'cmd':function name, 'sc': shortcut key}}
            var: an already created menu
        """

        if var is None:
            var = Menu(menu,tearoff=0)
        dialogs.applyStyle(var)
        items = list(dict.keys())
        items.sort()
        for item in items:
            if item[-3:] == 'sep':
                var.add_separator()
            else:
                command = dict[item]['cmd']
                label = '%-25s' %(item[2:])
                if 'img' in dict[item]:
                    img = dict[item]['img']
                else:
                    img = None
                if 'sc' in dict[item]:
                    sc = dict[item]['sc']
                    #bind command
                    #self.main.bind(sc, command)
                else:
                    sc = None
                var.add('command', label=label, command=command, image=img,
                        compound="left")#, accelerator=sc)
        dict['var'] = var
        return dict

    def quit(self):
        self.main.destroy()
        return


if __name__ == '__main__':
    app = EmailsExplore()
    app.mainloop()
