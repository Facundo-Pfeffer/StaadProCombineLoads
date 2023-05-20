from read_excel_and_combine import ReadAndTransformExcelCombinations, PutCombinationsInFile


input_excel_file = "ELU.xlsx"
input_sheet = "ELS"
STAAD_format_file_name = "base_staad_file.txt"


output_file = "ELS.txt"


combinations_dict_list = ReadAndTransformExcelCombinations(input_excel_file_name=input_excel_file, input_excel_sheet_name=input_sheet)

insert_values_in_file = PutCombinationsInFile(
    input_file_name=STAAD_format_file_name,
    output_file_name=output_file,
    text_to_append=combinations_dict_list.combinations_txt
)
insert_values_in_file.write_file_output()

