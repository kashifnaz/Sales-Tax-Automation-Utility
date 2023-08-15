from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow
from PyQt5.QtCore import QDir
import numpy as np
import pandas as pd
import sys
import STUtility
import os


class Main(QMainWindow, STUtility.Ui_MainWindow):
    def __init__(self, parent=None):
        super(Main, self).__init__(parent)
        self.setupUi(self)


        # Processing utility buttons
        self.browse_btn.clicked.connect(self.browse)
        self.load_btn.clicked.connect(self.load)
        self.process_btn.clicked.connect(self.process)
        self.output_btn.clicked.connect(self.output_directory)
        # self.save_btn.clicked.connect(self.save_files)

        # Converting utility buttons
        self.input_dir.clicked.connect(self.input_direct)
        self.move_right.clicked.connect(self.move_right_method)
        self.move_left.clicked.connect(self.move_left_method)
        self.move_right_all.clicked.connect(self.move_right_all_method)
        self.move_left_all.clicked.connect(self.move_left_all_method)
        self.file_select.itemSelectionChanged.connect(self.update_button_STATUS)
        self.file_convert.itemSelectionChanged.connect(self.update_button_STATUS)
        self.convert_save_btn.clicked.connect(self.convert_format)
        self.submit.clicked.connect(self.submit_method)
        self.actionExit.triggered.connect(self.exit_menu)

        self.update_button_STATUS()


    def browse(self):
        '''Browse excel files for processing utility'''
        self.filename = QFileDialog.getOpenFileName(self, 'Open excel file', directory=QDir.currentPath(), filter='Excel files (*.xlsx *.xls)')
        self.browse_line.setText(self.filename[0])

    def load(self):
        '''Load excel files and displaying sheets'''
        xl = pd.ExcelFile(self.filename[0])
        choices = [x for x in xl.sheet_names]
        self.load_status.clear()
        self.comboBox.clear()
        self.comboBox.addItems(choices)
        if len(choices) == 0:
            self.load_status.setText('Please load a file.')
        else:
            self.load_status.setText('File is loaded!')

    def output_directory(self):
        '''Folder for the output directory'''
        self.output = str(QFileDialog.getExistingDirectory(self, 'Output directory', directory=QDir.currentPath()))
        self.output_line.setText(self.output)

    def process(self):
        '''Conversion process including data wrangling, categorization, summarizing data'''
        # reading the excel sheet
        change_values = ['', '-', 'N/A', 'NOT REGISTERED', 'N?A', 'not', 'n/a', 'nan', ',', '.', 'N\A']
        sheet = self.comboBox.currentText()
        self.df = pd.read_excel(self.filename[0], sheet_name=sheet, skiprows=2, keep_default_na=False, na_values=change_values)
        self.df.fillna('', inplace=True)

        ## Removing spaces with underscore and renaming columns
        self.df = self.df.rename(columns={'ORDER NUMBER': 'INVOICE', 'ORDERED DATE': 'DATE', 'CUSTOMER NUMBER': 'CODE'})
        self.df = self.df.rename(columns=lambda x: x.replace(' ', '_'))

        ## Dropping unnecessary columns with ignoring errors
        self.df.drop(['ORG_ID', 'ADDRESS'], axis=1, inplace=True, errors='ignore')

        ## Rearraning columns
        column_names = ['STR_NEW','NTN','CNIC','CODE','SOLD_TO','TAX_CODE','INVOICE','DATE','ITEM_CODE','ITEM_NAME','UNIT_SELLING_PRICE','ORDERED_QTY','SALES','FREE','TRADE_PRICE','GROSS_AMOUNT','V_DISS','VALUE_FOR_GST','SALE_TAX','INCOME_TAX','NET_VALUE_TAX','BRANCH']
        self.df.reindex(column_names)

        ## Validate the CNIC DIGITS (13 without hypen) check by converting dtype to string and replacing last two characters with regex
        self.df = self.df.assign(CNIC=lambda x: x['CNIC'].str.replace(r'\D', '', regex=True))
        self.df['CNIC_DIGITS'] = self.df['CNIC'].map(len)
        self.df= self.df.assign(NTN_DIGITS=lambda x: x['NTN'].str.replace(r'\D', '', regex=True).map(len))

        ## Converting TAX_CODE into categories
        conditions = [
            (self.df['TAX_CODE'].str.contains('EXEMPTED')) | (self.df['TAX_CODE'].str.contains('O/P GST EXEMPT')),
            self.df['TAX_CODE'].str.contains('O/P GST ZERO'),
            self.df['TAX_CODE'].str.contains('MRP'),
            self.df['TAX_CODE'].str.contains('O/P GST 17'),
            self.df['TAX_CODE'].str.contains('O/P GST 12'),
        ]
        choices = [
            'EXEMPTED O/P GST EXEMPT',
            'O/P GST ZERO RATE',
            'O/P STAX MRP',
            'O/P GST 17%_N',
            'O/P GST 12%'
        ]

        ## Creating categorical column through np.select(for multiple choices)
        self.df['TAX_CODE_ST'] = np.select(conditions, choices, default=self.df['TAX_CODE'])

        ## Creating categorical column for Status on the basis of STRN number with np.where(two choices only)
        self.df['STATUS'] = np.where(self.df['STR_NEW'].eq(''), 'Unregistered', 'Registered')

        self.df_exp = self.df[self.df['STATUS'].eq('Unregistered')].groupby(['NTN','CNIC'])
        self.df_exp_data = self.df_exp.filter(lambda x: x.groupby(['NTN','CNIC'])['NET_VALUE_TAX'].sum().gt(100000000/12))
        self.df_exp_data_g = self.df_exp_data.groupby(['NTN', 'CNIC', 'SOLD_TO', 'BRANCH'])['NET_VALUE_TAX'].sum().reset_index()

        ## Exception reporting where STR_NEW is null, CNIC_DIGITS are less than 13 and NTN digits are less than 8 digits
        self.dfe = self.df[
            (self.df['STR_NEW'].eq(''))
            & ((self.df['CNIC_DIGITS'].lt(13) & (self.df['CNIC_DIGITS'].gt(0))) | (self.df['CNIC_DIGITS'].gt(13)))
            | (self.df['NTN_DIGITS'].lt(8) & self.df['NTN_DIGITS'].gt(0))
        ]

        ## Output exception in excel file
        options = {}
        options['strings_to_formulas'] = False
        writer1 = pd.ExcelWriter(self.output+'/'+'Exception.xlsx', engine='xlsxwriter', options=options)
        self.dfe.to_excel(writer1, sheet_name='CNIC_NTN', index=False, float_format='%.2f')
        self.df_exp_data_g.to_excel(writer1, sheet_name='Sale>8.3M_Pivot', index=False, float_format='%.2f')
        self.df_exp_data.to_excel(writer1, sheet_name='Sale>8.3M_data', index=False, float_format='%.2f')
        writer1.save()
        writer1.close()

        ## Define filenames for the filtered data
        dict_name = {
            'Exempt': 'EXEMPTED O/P GST EXEMPT',
            'Taxable_10': 'O/P S/TAX 12%',
            'Zero_Rate': 'O/P GST ZERO RATE',
            'Taxable_17': 'O/P GST 17%_N',
            'MRP': 'O/P STAX MRP'
        }
        ## Getting filtered data according to the tax categories into a dictionary
        df_dict= {}
        for name in dict_name.keys():
            df_dict[name] = self.df[self.df['TAX_CODE_ST'].eq(dict_name.get(name))]

        ## Making excel output of tax categories in df_dict and exporting seperate sheets for registered and unregistered customers
        annex_column = ['STR_NEW', 'NTN', 'CNIC', 'CODE', 'SOLD_TO', 'INVOICE', 'DATE','HS_CODE', 'SALES', 'FREE', 'VALUE_FOR_GST', 'SALE_TAX', 'INCOME_TAX']
        group_column = ['STR_NEW', 'NTN', 'CNIC', 'CODE', 'SOLD_TO', 'INVOICE', 'DATE', 'HS_CODE']

        for name in df_dict.keys():
            options = {}
            options['strings_to_formulas'] = False
            writer = pd.ExcelWriter(self.output+'/'+name+".xlsx", engine='xlsxwriter', options=options)
            temp = df_dict.get(name)
            # Pivot table of registered
            tempr = temp[temp['STATUS'].eq('Registered')].loc[:,annex_column].groupby(group_column)['SALES', 'FREE','VALUE_FOR_GST','SALE_TAX', 'INCOME_TAX'].sum().reset_index()
            # Eliminating zero values
            tempr[(tempr['VALUE_FOR_GST'].gt(0))].to_excel(writer, sheet_name='Annex_Reg', index=False, float_format='%.2f')
            # Seperate sheet for less than zero value for adjustment
            tempr[tempr['VALUE_FOR_GST'].lt(0) | (tempr['VALUE_FOR_GST'].eq(0))].to_excel(writer, sheet_name='Annex_Reg<0', index=False, float_format='%.2f')
            # Pivot table of Unregistered
            tempu = temp[temp['STATUS'].eq('Unregistered')].loc[:,annex_column].groupby(group_column)['SALES', 'FREE','VALUE_FOR_GST','SALE_TAX', 'INCOME_TAX'].sum().reset_index()
            # Eliminating zero values
            tempu[tempu['VALUE_FOR_GST'].gt(0)].to_excel(writer, sheet_name='Annex_Unreg', index=False, float_format='%.2f')
            # Seperate sheet for less than zero value for adjustment
            tempu[(tempu['VALUE_FOR_GST'].lt(0)) | (tempu['VALUE_FOR_GST'].eq(0))].to_excel(writer, sheet_name='Annex_Unreg_0', index=False, float_format='%.2f')
            temp[temp['STATUS'].eq('Registered')].to_excel(writer, sheet_name='Reg', index=False,  float_format='%.2f')
            temp[temp['STATUS'].eq('Unregistered')].to_excel(writer, sheet_name='Unreg', index=False,  float_format='%.2f')
            temp.to_excel(writer, sheet_name=name, index=False)
            writer.save()
            writer.close()

    def input_direct(self):
        self.input = QFileDialog.getExistingDirectory(self, 'Select input directory', directory=QDir.currentPath())
        self.input_dir_path.setText(self.input)
        self.file_select.clear()
        with os.scandir(self.input) as files:
            for i, j in enumerate(files):
                if not j.name.startswith('.') and j.is_file() and j.name.endswith('.xlsx'):
                    self.file_select.addItem(j.name)
                    if j.name == 'Exception.xlsx':
                        self.file_select.takeItem(i)

    def move_right_method(self, index):
        '''Move files from list to converted list and back'''
        row = self.file_select.currentRow()
        rowItem = self.file_select.takeItem(row)
        self.file_convert.addItem(rowItem)

    def move_left_method(self):
        '''Vice versa of move_right()'''
        row = self.file_convert.currentRow()
        rowItem = self.file_convert.takeItem(row)
        self.file_select.addItem(rowItem)

    def move_right_all_method(self):
        '''Move all items to the right'''
        for i in range(self.file_select.count()):
            self.file_convert.addItem(self.file_select.takeItem(0))

    def move_left_all_method(self):
        '''Move all items to the left'''
        for i in range(self.file_convert.count()):
            self.file_select.addItem(self.file_convert.takeItem(0))

    def submit_method(self):
        self.no_of_lines = self.textEdit.toPlainText()
        return self.no_of_lines

    def convert_format(self):
        '''Split and convert the file into SRB format for uploading'''
        files = []
        for index in range(self.file_convert.count()):
            files.append(self.file_convert.item(index))

        for file in files:
            xl_path = self.input+'/'+file.text()
            xl = pd.ExcelFile(xl_path)
            path_file = os.path.join(self.input, file.text().split('.')[0])
            os.makedirs(path_file, exist_ok=True)
            for sheet in ['Annex_Reg', 'Annex_Unreg']:
                path_sheet = os.path.join(path_file, sheet)
                os.makedirs(path_sheet, exist_ok=True)
                df_convert = xl.parse(sheet_name=sheet)
                chunk_size = 300 if self.no_of_lines == '' else int(self.no_of_lines)
                x = 1
                group = df_convert.groupby(np.arange(len(df_convert.index)) // chunk_size)
                for i, g in group:
                    os.chdir(path_sheet)
                    g.to_excel(('{}_{}_{:02d}.xls'.format(file.text().split('.')[0], sheet, x)), index=False, engine='xlwt')
                    x += 1

    def update_button_STATUS(self):
        self.move_right.setDisabled(not bool(self.file_select.selectedItems()) or self.file_select.count() == 0)
        self.move_left.setDisabled(not bool(self.file_convert.selectedItems()) or self.file_convert.count() == 0)

    def exit_menu(self):
        self.close()


def main():
    app = QApplication(sys.argv)
    form = Main()
    form.show()
    app.exec_()

if __name__=='__main__':
    main()

