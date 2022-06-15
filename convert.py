import tabula
import os
import re
import pandas as pd
from openpyxl import load_workbook




class Statement():
    def __init__(self, pdf_name):
        self.options = dict()  # Parameters for tabula.read_pdf()
        self.bank = None
        self.location = None
        self.office = None
        self.pdf_name = pdf_name
        self.dir = './'
        self.ADD_DIGITS = False
        self.OCR = False
        self.get_bank()
        self.get_office()
        self.year = int(re.compile(r'^\d{4}').search(pdf_name).group())
        self.set_workarounds()
        self.xls_name = self.pdf_name.replace(".pdf", ".xlsx")
 
    def check_xls(self):
        return self.get_file_size() > 5_500

    def check_dfs_have_data(self):
        """
        Check if any of the dfs in df_list have data.
        """
        for df in self.df_list:
            if len(df) > 0:
                return True
        return False

    def get_office(self):
        """
        Get the office from the filename.
        """
        if re.search("Dar", self.pdf_name):
            return "Dar"
        if re.search("Dodoma", self.pdf_name):
            return "Dodoma"
        if re.search("Musoma", self.pdf_name):
            return "Musoma"
        if re.search("Mbeya", self.pdf_name):
            return "Mbeya"
        if re.search("Katavi", self.pdf_name):
            return "Katavi"

    def get_bank(self):
        """
        Get the bank name from the filename.
        """
        if re.search("NBC", self.pdf_name):
            self.bank = 'NBC'
        if re.search("CRDB", self.pdf_name):
            self.bank = 'CRDB'
        return self.bank

    def set_workarounds(self):
        """
        Workarounds for statements that are not in the correct format.
        """
        if self.bank == 'CRDB' and self.office == 'Dar':
            self.options['lattice'] = False  # Parameters for tabula.read_pdf()
            self.options['stream'] = False
        if self.bank == 'NBC' and self.office != 'Dodoma':
            self.ADD_DIGITS = True
        if self.office == 'Katavi':
            self.OCR = True
        if self.bank == 'NBC':
            self.OCR = True
            self.options['area'] = (153.0,30.0,840.525,564.99)



    def convert_pdf_to_df(self):
        # print(len(tabula.read_pdf(pdf_name, pages='all')))\
        tables = tabula.read_pdf(self.dir + self.pdf_name, pages='all', **self.options)
        print(tables)
        self.df_list = []
        for table in tables:
            if len(table) > 0:
                # print(table)
                self.df_list.append(table)
        if len(self.df_list) == 0:
            print("No data found in PDF")
        # df = tabula.read_pdf(pdf_name, pages='all')[1]
    
    def fill_na(self):
        """
        Fill in missing values in the df.
        """
        for df in self.df_list:
            df.fillna("", inplace=True)

    def convert_df_to_excel(self):
        rows = 0
        if not self.df_list:
            with open(self.dir + self.xls_name, 'w') as _:
                print("Making empty file")  # Make an empty file, even though it has failed
                pass
        for df in self.df_list:
            if rows > 0:  # There is an existing sheet with some data already written, that needs to be opened and the new data written below
                print(f"Existing file, adding to it from row {rows}")
                book = load_workbook(self.dir + self.xls_name)
                writer = pd.ExcelWriter(self.dir + self.xls_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') 
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                # print(df)
                df.to_excel(writer, index=False, header=False, startrow=rows)
                writer.save()

                rows += len(df)
            else:
                print("New file, writing to it")
                df.to_excel(self.dir + self.xls_name, index=False, header=False)
                rows += len(df)
        print("Saving: " + self.xls_name)

    def add_digits(self):
        """
        Workaround for NBC statements where the final digit of the date gets chopped off.
        """
        # year = self.get_year(self.df_list[0].iloc[1,0])
        self.add_year(str(self.year)) # Just using the year from the filename

    @staticmethod
    def get_year(text):
        """
        Workaround for NBC statements where the final digit of the date gets chopped off.
        Find the year from the text at the top of the statement, that gives the statement dates.
        """
        year = ''
        try:
            year = re.compile(r"Account Statement For Period:.*[/]\d*1(\d)").match(text).group(1)
        except:
            pass
        return year

    @staticmethod
    def add_digit(text, digit):
        """
        Workaround for NBC statements where the final digit of the date gets chopped off.
        Manually add a digit to text.
        """
        return re.sub(r'(\d{2}[/]\d{2}[/]\d)(.*)', r'\g<1>' + digit, str(text))

    def add_year(self, year):
        """
        Workaround for NBC statements where the final digit of the date gets chopped off.
        Manually add the year to the second and fourth columns.
        """
        print(self.df_list[0])
        self.df_list[0].iloc[:, 1] = self.df_list[0].iloc[:, 1].apply(self.add_digit, args=(year,))
        self.df_list[0].iloc[:, 3] = self.df_list[0].iloc[:, 3].apply(self.add_digit, args=(year,))

    def get_file_size(self):
        """
        Get the file size of the PDF.
        """
        return os.path.getsize(self.dir + self.xls_name)

    def move_file(self, dir):
        """
        Move the file to a new directory.
        """
        os.rename(self.dir + self.xls_name, dir + self.xls_name)
        os.rename(self.dir + self.pdf_name, dir + self.pdf_name)
        self.dir = dir

    def ocr(self):
        """
        OCR the PDF and convert it to an Excel file.
        """
        dir = "ocr/"
        pdf_name = self.pdf_name.replace(" ", "\\ ")
        os.system("qpdf --decrypt --replace-input "+ self.dir + pdf_name)
        os.system("ocrmypdf " + self.dir + pdf_name + " " + dir + pdf_name + " --force-ocr")
        self.options['lattice'] = False
        self.options['stream'] = True
        self.options['area'] = (153.0,30.0,840.525,564.99)
        os.remove(self.dir + self.pdf_name)
        self.dir = dir

    def adjust_col_width(self):
        """
        Adjust the column width of the Excel file.
        """
        df = pd.read_excel(self.dir + self.xls_name)
        writer = pd.ExcelWriter(self.dir + self.xls_name, engine='xlsxwriter')
        sheetname = "Sheet1"
        # for sheetname, df in dfs.items():  # loop through `dict` of dataframes
        df.to_excel(writer, sheet_name=sheetname, header=False, index=False)  # send df to writer
        worksheet = writer.sheets[sheetname]  # pull worksheet object
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) *1.1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width
        writer.save()

    def replace_strings(self):
        """
        Replace strings in the Excel file.
        """
        for df in self.df_list:
            df.replace(r'\r', ' ', regex=True, inplace=True)


# df = tabula.read_pdf("2016-08 Dar CRDB TAS.pdf", pages='all')[0]
# df = df.iloc[1: , :]

# print(df)

# tabula.convert_into("2016-08 Dar CRDB TAS.pdf", "2016-08 Dar CRDB TAS.xls", output_format="xls", pages='all')

# df.to_excel("2016-08 Dar CRDB TAS.xls", header=False, index=False)

def list_files_in_dir(dir_path):
    file_list = []
    for file in os.listdir(dir_path):
        if file.endswith(".pdf"):
            file_list.append(file)
    return file_list



def main():
    dir_path = "."
    file_list = list_files_in_dir(dir_path)
    for file in file_list:
        print("Processing file: " + file)
        statement = Statement(file)
        if statement.OCR:
            statement.ocr()
        statement.convert_pdf_to_df()
        print(statement.dir)
        if not statement.check_dfs_have_data() and not statement.OCR:
            print("No data found in PDF, trying OCR")
            statement.OCR = True
            statement.ocr()
            statement.convert_pdf_to_df()
        statement.fill_na()
        statement.replace_strings()
        if statement.ADD_DIGITS:
            statement.add_digits()
        statement.convert_df_to_excel()
        statement.adjust_col_width()
        if statement.check_xls():
            statement.move_file("processed/")
        else:
            statement.move_file("failed/")
    

def test():
    pdf_name = '2015-11 Mbeya NBC TAS.pdf'
    dir = './'
    pdf_name = pdf_name.replace(" ", "\\ ")

    os.system("qpdf --decrypt --replace-input "+ dir + pdf_name)

    # df = tabula.read_pdf(dir + pdf_name, pages='all', lattice = True, stream=True, multiple_tables=True, area=(156.0,30.0,600.0,580.035))
    # print(df)
    

if __name__ == "__main__":
    main()
    # test()
    # convert_pdf_to_excel("2016-09 Mbeya NBC TAS.pdf")
    # print(add_digit('09/09/16', '6'))

    # print(tabula.read_pdf(pdf_name, pages='all', lattice = True))
    # print(tables)
    # data = []
    # for table in tables:
    #     if len(table) > 0:
    #         for row in table:
    #             data.append(row)
    # # print(f"Data is: {data}")
    # df = pd.DataFrame(data)
    # print(df)