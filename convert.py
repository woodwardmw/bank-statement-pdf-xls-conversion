import tabula
import os
import re
import pandas as pd
from openpyxl import load_workbook

class Statement():
    def __init__(self, pdf_name):
        self.options = dict()  # Parameters for tabula.read_pdf()
        self.bank = None
        self.office = None
        self.pdf_name = pdf_name
        self.dir = './'
        self.ADD_DIGITS = False
        self.OCR = False
        self.get_bank()    # Get the bank name from the filename
        self.get_office()  # Get the office from the filename
        self.year = int(re.compile(r'^\d{4}').search(pdf_name).group())  # Get the year from the filename
        self.set_workarounds()  # Set the workarounds according to the type of statement
        self.xls_name = self.pdf_name.replace(".pdf", ".xlsx")
 
    def check_xls(self):
        return self.get_file_size() > 5_500  # If the file size is less than 5500 bytes, then it is probably a blank file

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
            self.options['area'] = (153.0,30.0,840.525,564.99)  # This area should generally cover the main part of a statement, and is given as an argument for tabula.read_pdf()

    def convert_pdf_to_df(self):
        tables = tabula.read_pdf(self.dir + self.pdf_name, pages='all', **self.options)
        self.df_list = []  # List of dataframes, to store the multiple dataframes from the pdf
        for table in tables:
            if len(table) > 0:
                self.df_list.append(table)
        if len(self.df_list) == 0:  # If there are no dataframes, then the file is probably blank
            print("No data found in PDF")
    
    def fill_na(self):
        """
        Fill in missing values in the df with empty strings.
        """
        for df in self.df_list:
            df.fillna("", inplace=True)

    def convert_df_to_excel(self):
        rows = 0
        if not self.df_list:
            with open(self.dir + self.xls_name, 'w') as _:
                print("Making empty file")  # Make an empty file, even though it has failed
        for df in self.df_list:
            if rows == 0:  # There is no existing sheet, so make a new one
                print("New file, writing to it")
                df.to_excel(self.dir + self.xls_name, index=False, header=False)
                rows += len(df)
            else:  # There is an existing sheet, so append the new data to it, starting at row numeber 'row'
                book = load_workbook(self.dir + self.xls_name)
                writer = pd.ExcelWriter(self.dir + self.xls_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') 
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                df.to_excel(writer, index=False, header=False, startrow=rows)
                writer.save()
                rows += len(df)
        print("Saving: " + self.xls_name)

    def add_digits(self):
        """
        Workaround for NBC statements where the final digit of the date gets chopped off.
        """
        self.add_year(str(self.year)) # Using the year from the filename

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
        self.df_list[0].iloc[:, 1] = self.df_list[0].iloc[:, 1].apply(self.add_digit, args=(year,))
        self.df_list[0].iloc[:, 3] = self.df_list[0].iloc[:, 3].apply(self.add_digit, args=(year,))

    def get_file_size(self):
        """
        Get the file size of the XLS file.
        """
        return os.path.getsize(self.dir + self.xls_name)

    def move_file(self, dir):
        """
        Move the file (both pdf and xls) to a new directory and update its self.dir attribute.
        """
        os.rename(self.dir + self.xls_name, dir + self.xls_name)
        os.rename(self.dir + self.pdf_name, dir + self.pdf_name)
        self.dir = dir

    def ocr(self):
        """
        OCR the PDF so it can be processed by tabula.
        """
        dir = "ocr/"
        pdf_name = self.pdf_name.replace(" ", "\\ ")  # Escape the spaces to avoid errors in the command
        os.system("qpdf --decrypt --replace-input "+ self.dir + pdf_name)  # Decrypt the PDF
        os.system("ocrmypdf " + self.dir + pdf_name + " " + dir + pdf_name + " --force-ocr") # OCR the PDF
        self.options['lattice'] = False
        self.options['stream'] = True  # Parameters for tabula.read_pdf() - these seem to work best with OCR files
        self.options['area'] = (153.0,30.0,840.525,564.99)
        os.remove(self.dir + self.pdf_name)  # Delete the original PDF
        self.dir = dir

    def adjust_col_width(self):
        """
        Adjust the column width of the Excel file to fit the data.
        """
        df = pd.read_excel(self.dir + self.xls_name)  # Read the Excel file to a df
        writer = pd.ExcelWriter(self.dir + self.xls_name, engine='xlsxwriter')  # Create a writer object from the Excel file
        sheetname = "Sheet1"  # Name of the sheet
        df.to_excel(writer, sheet_name=sheetname, header=False, index=False)  # write the df back to the xls file
        worksheet = writer.sheets[sheetname]  # pull worksheet object
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) * 1.1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width
        writer.save()

    def replace_strings(self):
        """
        Replace '\r' strings in the Excel file, which don't display well in Excel.
        """
        for df in self.df_list:
            df.replace(r'\r', ' ', regex=True, inplace=True)

def list_files_in_dir(dir_path):
    """
    List all files in a directory.
    """
    file_list = []
    for file in os.listdir(dir_path):
        if file.endswith(".pdf"):
            file_list.append(file)
    return file_list

def has_pdf_extention(filename):
    """
    Check if the filename has the .pdf extension.
    """
    return filename.endswith(".pdf")

def main():
    dir_path = "./"
    file_list = list_files_in_dir(dir_path) # Get a list of all files in the directory
    for file in file_list:
        if has_pdf_extention(file): # Check if the file has the .pdf extension, otherwise ignore the file
            print("Processing file: " + file)
            statement = Statement(file)
            if statement.OCR:
                statement.ocr()
            statement.convert_pdf_to_df()
            if not statement.check_dfs_have_data() and not statement.OCR:  # If there is no data, and OCR hasn't been tried, try OCR
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
            if statement.check_xls():  # Move the file, based on whether the xls file seems large enough to contain genuine data
                statement.move_file("processed/")
            else:
                statement.move_file("failed/")

if __name__ == "__main__":
    main()