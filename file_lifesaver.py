import openpyxl
import os
import csv

def file_lifesaver(output_path="File_Lifesaver_Output", file_format="csv"):
    """Put the script in the folder with .xlsx files that you want to save as txt
    file or any other extension with a delimiter of your choice. """
    try:
        os.mkdir(output_path)
    except FileExistsError:
        pass

    for excelFile in os.listdir():
        if excelFile.endswith(".xlsx"):
            workbook = openpyxl.load_workbook(excelFile, data_only=True)

            for sheet_name in workbook.sheetnames:
                current_sheet = workbook[sheet_name]

                with open(output_path + "/{}_{}.{}".format(excelFile, sheet_name, file_format),
                          "w+", newline="") as text_file:
                    excel_writer = csv.writer(text_file, delimiter="\t") #, quotechar='"', quoting=csv.QUOTE_ALL
                    for row in current_sheet:
                        row_data = []

                        for column in row:
                            row_data.append(column.value)
                        excel_writer.writerow(row_data)


file_lifesaver()
