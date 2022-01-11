from slugify import slugify
import requests
import xlsxwriter
import os
import sys


class ExcelFile:
    """
    Class to handle building the Excel file.
    Required arguments:
    :::filename::: This specifies the name of the document.
    Optional arguments:
    :::path::: Specifies where to save the Excel file. By default, this will be the users documents.
    """
    def __init__(self, filename, path=os.path.join(os.path.expanduser("~"), "Documents")):
        if filename[-5:] == ".xlsx":
            self.filename = filename
        else:
            self.filename = f"{filename}.xlsx"

        self.path = path
        self.full_filepath = os.path.join(self.path,self.filename)
        self.workbook = xlsxwriter.Workbook(self.full_filepath)
        self.sheet_list = []
        self.bold = self.workbook.add_format({"bold":True})
        self.boldCentral = self.workbook.add_format({"bold":True,"align":"center"})
        self.central =  self.workbook.add_format({"align":"center"})

    def longest_value(self, list_name, list_index_to_search):
        """
        Adjusts column widths based on the longest column
        Required arguments:
        :::list_name::: The list of values to check values
        ::list_index_to_search::: The index of the list to check
        """
        maxValue = 0
        for item in list_name:
            try:
                if len(str(item[list_index_to_search])) > maxValue:
                    maxValue = len(str(item[list_index_to_search]))
            except:
                pass
        return maxValue
    
    def create_sheet(self, sheet_name):
        """
        Creates a sheet in the excel file.
        Required arguments:
        :::sheet_name::: The name of the sheet to create
        """
        if sheet_name in self.sheet_list:
            return "Sheet already created"
        else:
            self.sheet_list.append(sheet_name)
            exec(f"self.{sheet_name} = self.workbook.add_worksheet("{sheet_name}")")

    def add_data(self, list_of_dict, sheet_name):
        """
        Adds the data (and headers) to the excel file
        Required arguments:
        :::list_of_dict::: A list of dictionaries to iterate through to add to the excel file
        :::sheet_name::: The name of the sheet for the values to be added to
        """

        self.create_sheet(sheet_name)

        # Create the headers
        header_col = "A"
        header_col_no = 0

        if len(list_of_dict) > 0:
            header_list = list(list_of_dict[0].keys())
            data_list = []
            data_list.append(header_list)
            for d in list_of_dict:
                data_list.append(list(d.values()))

            for i, d in enumerate(data_list):
                col = "A"
                col_no = 0
                for d1 in d:
                    d1 = str(d1).replace("'", "\\\'")
                    if i == 0:
                        exec(f"self.{sheet_name}.write({i}, {col_no}, '{d1}', self.bold)")
                    else:
                        exec(f"self.{sheet_name}.write({i}, {col_no}, '{d1}')")
                    exec(f"self.{sheet_name}.set_column("{col}:{col}",max(self.longest_value(data_list,{col_no}),len('{d1}'))+5)")
                    col = chr(ord(col) + 1) 
                    col_no = col_no + 1
            exec(f"self.{sheet_name}.freeze_panes(1, 0)")

    def save(self):
        """
        Closes the newly created workbook
        """
        self.workbook.close()
    
try:
    API_URL_ROOT = "https://interviewbom.herokuapp.com/"

    # Get the initial BoM data
    r = requests.get(f"{API_URL_ROOT}bom/")

    data = r.json()["data"]
    results = []

    # For each part in the BoM, get the part number and quantity and add the dictionary to a list
    for d in data:
        parent_part_id = d["parent_part_id"]
        part_id = d["part_id"]

        if part_id:
            r = requests.get(f"{API_URL_ROOT}part/{part_id}/")
            results.append(
                {
                    "parent_part": parent_part_id,
                    "part_number": r.json()["part_number"],
                    "quantity": d["quantity"]
                }
            )

    # Get the file name from the system argument, and process the excel file
    file_name = sys.argv[1]

    excel_file = ExcelFile(file_name)
    excel_file.create_sheet("PartList")
    excel_file.add_data(results, "PartList")
    excel_file.save()

except IndexError:
    print("Please specify a file name")