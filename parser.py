from xlrd import open_workbook
from os import listdir
import sys
import re

STATUSES = ("Complete", "Blank", "Not Applicable", "Partially Complete", "Not Coded")
EXCEL_FILE_REGEX = re.compile(r".+\.xlsx?$", re.IGNORECASE)
EXCEL_TEMP_FILE_REGEX = re.compile(r"^~.+", re.IGNORECASE)

class SheetData():
    '''
    Class doc would go here
    '''

    '''
    Key -> field name
    Value -> list that looks like
        0 -> dict that looks like:
            Key -> STATUSES, strings
            Value -> int count
        1 -> dict that looks like:
            Key -> string comment
            Value -> int count
        2 -> int, number of times data added for this column, field count
    '''        
    data = {}

    def __init__(self):
        self.data = {}

    def add_data(self, field_name, status, comment):
        if(field_name not in self.data):
            self.data[field_name] = [self.new_status_dict(),{}, 0]

        # Status    
        if(status in STATUSES):
            self.data[field_name][0][status] += 1
        else:
            self.data[field_name][0]["Not Coded"] += 1

        # Comments
        if(comment is None or comment is ''):
            pass
        elif(comment in self.data[field_name][1]):
            self.data[field_name][1][comment] += 1
        else:
            self.data[field_name][1][comment] = 1

        # Field count
        self.data[field_name][2] += 1

    def new_status_dict(self):
        status_dict = {}
        for status in STATUSES:
            status_dict[status] = 0
        return status_dict

    def get_fields(self):
        return self.data.keys()

    def get_status_count(self, field_name, status):
        return self.data[field_name][0][status]

    def get_comments(self, field_name):
        return self.data[field_name][1].keys()

    def get_comment_count(self, field_name, comment):
        return self.data[field_name][1][comment]

    def get_field_count(self, field_name):
        return self.data[field_name][2]



# --- MAIN ---
COLUMNS = {
    "FIELD": 0, # A
    "STATUS": 1, # B
    "COMMENT": 3, # D
}

'''
Key -> string, sheet name
Value -> SheetData
'''
data = {}

# Directory with only Excel sheets in it
data_dir = sys.argv[1]

# Collect data from all files in the spreadsheets dir
for file in listdir(data_dir):
    # Only do this for Excel files (skips OS meta files like .DS_Store)
    if (not EXCEL_FILE_REGEX.match(file)) or EXCEL_TEMP_FILE_REGEX.match(file):
        continue

    workbook = open_workbook("{}/{}".format(data_dir, file))
    for sheet in workbook.sheets():
        # Setup the sheet if not seen before
        if sheet.name not in data:
            data[sheet.name] = SheetData()

        # Loop through rows, starting at row 2 (i.e. ignoring header row)
        for row in range(1, sheet.nrows):
            field_name_val = sheet.cell(row, COLUMNS["FIELD"]).value.replace(",","")
            if field_name_val:
                field_name = "[{}] {}".format(row+1, field_name_val)
                status = sheet.cell(row, COLUMNS["STATUS"]).value
                try:
                    comment = sheet.cell(row, COLUMNS["COMMENT"]).value.replace(",","")
                except:
                    comment = None
                
                data[sheet.name].add_data(field_name, status, comment)

# Make CSV
print("SECTION,FIELD,FIELD COUNT,COMPLETE,BLANK,NOT APPLICABLE,PARTIALLY COMPLETE,NOT CODED,COMMENTS")
for sheet_name in data:
    for field_name in data[sheet_name].get_fields():
        comment_str = ""
        for comment in data[sheet_name].get_comments(field_name):
            count = data[sheet_name].get_comment_count(field_name, comment)
            comment_str += "-- {} ({})".format(comment,count)

        field_count = float(data[sheet_name].get_field_count(field_name))
        complete_count = float(data[sheet_name].get_status_count(field_name,"Complete"))
        blank_count = float(data[sheet_name].get_status_count(field_name,"Blank"))
        na_count = float(data[sheet_name].get_status_count(field_name,"Not Applicable"))
        partial_count = float(data[sheet_name].get_status_count(field_name,"Partially Complete"))
        not_coded_count = float(data[sheet_name].get_status_count(field_name,"Not Coded"))

        print("{},{},{},{},{},{},{},{},{}".format(
            sheet_name,
            field_name,
            field_count,
            complete_count / field_count * 100,
            blank_count / field_count * 100,
            na_count / field_count * 100,
            partial_count / field_count * 100,
            na_count / field_count * 100,
            comment_str
        ))
