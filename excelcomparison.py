import openpyxl
from openpyxl.styles import PatternFill

data_file1 = openpyxl.load_workbook("Target_actual_data.xlsx")
data_file2 = openpyxl.load_workbook("Target_expected_Data.xlsx")

fill_style = PatternFill(start_color="FDD835", end_color="FDD835",fill_type="solid")
null_fill=PatternFill(start_color="FF0000", end_color="FF0000" , fill_type="solid")
data_sheet1 = data_file1['Actual']
data_sheet2 = data_file2['Expected']
data_sheet3 = data_file1['null_scenario']

print(f'{data_sheet1} found')
print(f'{data_sheet2} found')

#code for record count comparison
def count_records(filename):
    line_count=0
    for reccount in filename:
        line_count += 1
    return line_count

first_file_record_count=count_records(data_sheet1)
second_file_record_count=count_records(data_sheet2)
#print(f'"record count of first file is {first_file_record_count}"')
#print(f'"record count of second file is {second_file_record_count}"')

if first_file_record_count==second_file_record_count:
    print(f'{data_sheet1} record count is  {first_file_record_count}  and {data_sheet2}  record count is  {second_file_record_count}  is matches')
else:
    print("records count mismatch")


# code for rows comparison
for row in data_sheet1.iter_rows():
    for cell in row:
        current_cell_value=cell.value
        cell_location=cell.coordinate

        if current_cell_value!=data_sheet2[cell_location].value:
            cell.fill=fill_style
            print(data_sheet2[cell_location])
    #print(len(row))
#data_file1.save("compared_file.xlsx")


#code for checking null values in a column
def null_value(dataSheet):

    for null_row in dataSheet.iter_rows():
        # print(null_row[0].value)
        # null_row +=1
        for cell1 in null_row:
            current_cell_value1=cell1.value

            if current_cell_value1 is None:
                cell1.fill=null_fill
                # print(cell1.value)
            # print(f'null value found at {current_cell_value1}')
        # return null_row

# null_records=null_value(data_sheet1)
null_value(data_sheet3)
data_file1.save("Test_result.xlsx")