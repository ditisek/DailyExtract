import wx
import codecs
import pandas as pd
import openpyxl
import psutil

p = psutil.Process()

# begin wxGlade: dependencies
# end wxGlade

# begin wxGlade: extracode
# end wxGlade


class MainFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MainFrame.__init__
        kwds["style"] = kwds.get("style", 0) | wx.CAPTION | \
                        wx.CLIP_CHILDREN | wx.CLOSE_BOX | wx.MAXIMIZE_BOX | \
                        wx.RESIZE_BORDER | wx.SYSTEM_MENU
        wx.Frame.__init__(self, *args, **kwds)
        self.SetSize((930, 773))
        self.SetTitle("Extract Messages")

        self.window_1 = wx.SplitterWindow(self, wx.ID_ANY)
        # self.window_1.SetMinSize((628, 734))
        self.window_1.SetSize((380, 750))
        self.window_1.SetMinimumPaneSize(20)

        self.window_1_pane_1 = wx.Panel(self.window_1, wx.ID_ANY)
        # self.window_1_pane_1.SetMinSize((140, 734))
        self.window_1_pane_1.SetSize((240, 734))

        sizer_1 = wx.BoxSizer(wx.VERTICAL)

        sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(sizer_3, 0, wx.EXPAND, 0)

        label_1 = wx.StaticText(self.window_1_pane_1, wx.ID_ANY,
                                "File extension: *.")
        sizer_3.Add(label_1, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 2)

        self.text_ctrl_1 = wx.TextCtrl(self.window_1_pane_1, wx.ID_ANY,
                                       "xlsm")
        sizer_3.Add(self.text_ctrl_1, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        self.button_1 = wx.Button(self.window_1_pane_1, wx.ID_ANY,
                                  "Select files")
        sizer_1.Add(self.button_1, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 4)

        self.button_3 = wx.Button(self.window_1_pane_1, wx.ID_ANY,
                                  "Read headers from files")
        sizer_1.Add(self.button_3, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 4)

        self.check_list_box_2 = wx.CheckListBox(
            self.window_1_pane_1, wx.ID_ANY, choices=["headers"])
        self.check_list_box_2.SetMinSize((101, 200))
        sizer_1.Add(self.check_list_box_2, 0, wx.ALIGN_CENTER_HORIZONTAL, 0)

        self.button_2 = wx.Button(self.window_1_pane_1, wx.ID_ANY,
                                  "Export selected")
        sizer_1.Add(self.button_2, 0,
                    wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 4)

        self.window_1_pane_2 = wx.Panel(self.window_1, wx.ID_ANY)
        self.window_1_pane_2.SetSize((300, 750))

        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)

        self.check_list_box_1 = wx.CheckListBox(
            self.window_1_pane_2, wx.ID_ANY, choices=["choice 1"])
        # self.check_list_box_1.SetSize((300, 750))
        self.check_list_box_1.SetMinSize((650, 730))
        sizer_2.Add(self.check_list_box_1, 0, wx.ALL | wx.EXPAND, 0)
        # sizer_2.Add(self.check_list_box_1, 0, wx.ALL, 0)

        self.window_1_pane_2.SetSizer(sizer_2)

        self.window_1_pane_1.SetSizer(sizer_1)

        self.window_1.SplitVertically(self.window_1_pane_1,
                                      self.window_1_pane_2)

        self.Layout()

        self.Bind(wx.EVT_BUTTON, self.SelectFiles, self.button_1)
        self.Bind(wx.EVT_BUTTON, self.ReadHeaders, self.button_3)
        self.Bind(wx.EVT_BUTTON, self.ExportSelected, self.button_2)
        # end wxGlade

    def SelectFiles(self, event):  # wxGlade: MainFrame.<event_handler>
        # data = []

        print(p.open_files())

        extension = self.text_ctrl_1.GetValue()
        extension = "*." + extension

        path = MyFileDialog(None, wildcard=extension)
        files = path.EventHandler.Paths
        sysno = []
        datum = []
        sysinfo = []
        for file in files:
            # print(file)
            xl = pd.ExcelFile(file)
            # df1 = xl.parse(xl.sheet_names[0])
            df1 = xl.parse('Daily Report')
            df2 = xl.parse('Email Summary')

            status = df2.iloc[19, 0].split()[-1]
            try:
                comment = df2.iloc[25, 0].replace(',', ' ')
            except:
                comment = ''

            for item in df1.columns:
                if (item.upper().startswith('VTEM')) or \
                        (item.upper().startswith('ZTEM')) or \
                        (item.upper().startswith('FWZTEM')):
                    # print(item)
                    sysno_f = item

            sysconfig = df1.loc[0]
            no_rows = df1.shape[0]

            # print('Columns:')
            # for col in df1.columns:
            #     print(col)

            found = False

            data = [[] for i in range(no_rows + 1)]

            # print('Rows:')
            for row in df1.index:
                # print(f'Row {row} contains:')

                for item in df1.loc[row].dropna():
                    # print(item)
                    data[row].append(item)
                # print(df1.loc[row].dropna())

            # conf_row = 'Revision 2.02 Confidential Information - for Geotech personnel only. '
            # date_row = ' Send to: fieldreports@geotech.ca or fieldreports@geotechairbone.com'

            # print(p.open_files())

            for row in data:
                # row = str(row)
                # print(row)
                if len(row) > 1:
                    row[0] = str(row[0])
                    if 'Revision' in row[0]:
                    # if row[0] == conf_row:
                        system_conf = row[-2]
                    # elif row[0] == date_row:
                    elif 'Send' in row[0]:
                        datuminside = row[-2].strftime('%Y/%m/%d')
                    # elif row[0] == 'Client':
                    elif 'Client' in row[0]:
                        client = row[1].replace(',', ' ')
                        job_no = row[-1]
                    elif 'Survey Location' in row[0]:
                        location = row[1].replace(',', ' ')
                    elif 'Crew Chief' in row[0]:
                        chief = row[1]
                    # elif 'Operator' in row[0]:
                    #     if row[1] != 'Registration':
                    #         operator = row[1]
                    elif 'Project Manager' in row[0]:
                        pm = row[1]
                    elif 'DataQC Processor' in row[0]:
                        qc = row[1]

            if data[6][1] == 'Registration':
                operator = 'No operator'
            else:
                operator = data[6][1]
            newlines = data[13][0]
            reflights = data[13][1]
            dailytotal = data[13][2]
            qckm = data[13][3]
            qcperc = data[13][4]
            totalkm = data[13][5]
            acctotal = data[13][6]
            perccomp = data[13][7]
            tofly = data[13][8]

            print(f'{datuminside},{sysno_f},{system_conf},{job_no},{location},'
                  f'{client},{pm},{qc},{chief},{operator},{status},'
                  f'{newlines},{reflights},{dailytotal},{qckm},{totalkm},'
                  f'{acctotal},{perccomp},{tofly},{comment}')

            # for no, row in enumerate(df1.loc[1].dropna()):
            #     print(row)
            #     if row == 'Date':
            #         print('yes')
            #         datum.append(df.loc[1][no + 1])

            file_name = file.split('\\')[-1]
            for field in file_name.split():
                if field.startswith('VTEM'):
                    sysno.append(field)
                elif field.endswith('xlsm'):
                    datum.append(field.split('.')[0][-8:])
                # print(field)
            # print(file_name)
        for s in range(0, len(sysno)):
            sysinfo.append(sysno[s] + ' ' + datum[s])
        # sysno = [file.rsplit('\\')[-1].split(' ')[-2] for file in files]
        # datum = [file.rsplit('\\')[-1].split(' ')[-1].split('.')[0][-8:] for file in files]
        # print()


        self.check_list_box_1.SetItems(sysinfo)

        # print(p.open_files())

        event.Skip()

    def ReadHeaders(self, event):  # wxGlade: MainFrame.<event_handler>
        data = []
        headers = []

        extension = self.text_ctrl_1.GetValue()

        files = self.check_list_box_1.GetItems()

        for file_no in range(0, len(files)):

            try:
                f = open(files[file_no], "r")
                print(f"Reading file: {files[file_no]}")
                for line in f:
                    headers.append(line.rsplit(',')[0].rstrip())
                self.check_list_box_1.SetSelection(file_no)
                # self.check_list_box_1.SetCheckedItems(file_no)
            except:
                print("error in file {0}".format(files[file_no]))

        headers = list(set(headers))
        # headers = [header.rstrip() for header in headers]
        # headers = list(set(headers))
        headers = sorted(headers)
        self.check_list_box_2.SetItems(headers)

        print("Done reading headers")
        event.Skip()

    def ExportSelected(self, event):  # wxGlade: MainFrame.<event_handler>
        headers = self.check_list_box_2.GetCheckedStrings()

        path = MyNewFileDialog(None)
        output_file = path.EventHandler.Path
        print(output_file)

        print(self.check_list_box_1.GetItems())

        files = self.check_list_box_1.GetItems()

        with open(output_file, 'w') as out:
            for no, file in enumerate(files):
                print(f"Reading from: {file}")
                self.check_list_box_1.SetSelection(no)
                with open(file, 'r') as data_file:
                    for line in data_file:
                        for header in headers:
                            if line.startswith(header):
                                # print(line)
                                out.write(line)
        print(headers)
        print("Created file")
        event.Skip()

# end of class MainFrame

class MyFileDialog(wx.FileDialog):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MyFileDialog.__init__
        kwds["style"] = kwds.get("style", 0) | wx.FD_OPEN | \
                        wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE
        wx.FileDialog.__init__(self, *args, **kwds)
        self.SetTitle("Select files")

        self.ShowModal()
# end of class MyFileDialog


class MyNewFileDialog(wx.FileDialog):
    def __init__(self, *args, **kwds):
        kwds["style"] = kwds.get("style", 0) | wx.FD_SAVE
        kwds['defaultFile'] = 'data.csv'
        wx.FileDialog.__init__(self, *args, **kwds)
        self.SetTitle("Enter filename")

        self.ShowModal()
# end of class MyFileDialog


class MyApp(wx.App):
    def OnInit(self):
        self.frame = MainFrame(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True

# end of class MyApp

if __name__ == "__main__":
    app = MyApp(0)
    app.MainLoop()