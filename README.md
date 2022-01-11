# QuickReleaseTask
Python process to generate an .xlsx file based on the results of an API query. Built using Python 3.8.10

# Dependencies
Dependencies found in requirements.txt

# Process
Process calls the provided API to get parts_id, quantity, and parent_part_id. Using the part_id from this call, results are for eached through to get the part_number using the part_id. This builds a dictionary of parent_part_id, part_number, and quantity. This dictionary is added to a list of dictionaries, which is processed using the ExcelFile class. This is then saved to the users documents file, using a specified file_name when the user runs:
py rollup.py <file_name>