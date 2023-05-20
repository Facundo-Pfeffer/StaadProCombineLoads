from excel_manager import ExcelManager
from string import ascii_uppercase


class ReadAndTransformExcelCombinations:
    combination_to_staad_id_dict = {
        "D": 1,
        "L1": 2,
        "L2": 3,
        "L3": 4,
        "L4": 5,
        "Lr1": 6,
        "Lr2": 7,
        "Lr3": 8,
        "W1": 9,
        "W2": 10,
        "W3": 11,
        "W4": 12,
        "T": 13,
    }

    append_format = "LOAD COMB {combination_number} COMBINATION LOAD CASE {combination_number}\n" \
                    "{combinations_str}"

    def __init__(self, input_excel_file_name, input_excel_sheet_name):
        self.excel_wb = ExcelManager(file_name=input_excel_file_name, sheet_name=input_excel_sheet_name)
        self.combinations_txt = self.get_combinations()

    def get_combinations(self):
        result = ""
        for row_number in range(3, 300):

            combination_id = self.excel_wb.get_value("A", row_number)
            if combination_id is None:
                break
            if isinstance(combination_id, str) and not(combination_id.isdigit()):  # Is a merged cell
                continue

            elu_result = ""
            for column_letter in ascii_uppercase[2:15]:  # C to N
                combination_symbol = self.excel_wb.get_value(column_letter, 2)
                combination_staad_number = self.combination_to_staad_id_dict[combination_symbol]
                combination_factor = self.excel_wb.get_value(column_letter, row_number)
                if combination_factor and combination_factor != 0:
                    elu_result = elu_result + f"{combination_staad_number} {combination_factor} "

            result = result + self.append_format.format(
                combination_number=row_number-3+100,
                combinations_str=elu_result
            ) + "\n"


        return result


class PutCombinationsInFile:

    def __init__(self, input_file_name, output_file_name, text_to_append):
        file_input_txt = self.read_txt_file(input_file_name)
        self.file_output_txt = file_input_txt.format(combinations_to_add=text_to_append)

        self.output_file_name = output_file_name
        self.text_to_append = text_to_append

    @staticmethod
    def read_txt_file(file_name):
        with open(file_name, 'r') as f:
            result = f.read()
        return result

    def write_file_output(self):
        with open(self.output_file_name, "w") as f:
            f.write(self.file_output_txt)
