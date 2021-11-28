# ------------------------------------------------ Imports -------------------------------------------------------------
from datetime import datetime, timedelta
import xlsxwriter
import os
import sys
# ------------------------------------------------- Class --------------------------------------------------------------


class EmployeesHours:
    _NAME = "Name"                                                  # One of the table titles.
    _DATE = "Date"                                                  # One of the table titles.
    _ENTRANCE = "Entrance"                                          # One of the table titles.
    _EXIT = "Exit"                                                  # One of the table titles.
    _TOTAL = "Total"                                                # One of the table titles.
    _DATETIME_FORMAT = "%d/%m/%Y %H:%M"                             # DateTime format for DateTime objects.
    _DATE_FORMAT = "%d/%m/%Y"                                       # Date format for DateTime objects.
    _TIME_FORMAT = "%H:%M"                                          # Time format for DateTime objects.
    _GRANTED = "granted"                                            # Valid employee file line postfix
    _INPUT_PATH = None                                              # The path with the input files
    _OUTPUT_PATH = None                                             # The path in which the output file will be saved
                                                                    # created at.
    _OUTPUT_FILE_NAME = "Employees.xlsx"                            # The name of the output file.
    _WORKBOOK = None                                                # The output excel workbook.
    _WORKSHEET = None                                               # The output excel worksheet.

    def __init__(self, input_path=None, output_path=None):
        """
        Constructor.
        :param path: A path in which the employees report will be created at.
        """
        self._EMPLOYEES_DATA = {}
        if input_path is not None:
            self._INPUT_PATH = input_path
        if output_path is not None:
            self._OUTPUT_PATH = output_path

    def _get_row_datetime(self, row):
        """
        Returns a datetime object according to row
        :param row: A string representing a time.
        """
        try:
            row_datetime = datetime.strptime(row[:16], self._DATETIME_FORMAT)
            return row_datetime
        except Exception:
            pass

    def _get_employee_data(self, employee_name, employee_file_data):
        """
        Gets an employee name and a string which contains a single entrance or exit data of the employee, and creates
        a dictionary with employee data, then adds this employee dictionary to the bigger, all employees dictionary.
        :param employee_name:
        :type employee_name:
        :param employee_file_data:
        :type employee_file_data:
        :return:
        :rtype:
        """
        if employee_file_data:
            employee_data = {}
            for row in employee_file_data:
                row = row.replace("\n", " ").replace("\t", " ").rstrip()  # remove redundant \n and \t notes from row.
                if row.endswith(self._GRANTED):                           # entrance\exit data in row should be parsed
                    row_datetime = self._get_row_datetime(row)
                    if row_datetime is not None:
                        row_key = row_datetime.strftime(self._DATE_FORMAT)
                        if employee_data.get(row_key) is None:            #
                            employee_data[row_key] = [row_datetime]
                        else:
                            employee_data[row_key].append(row_datetime)

            for date in employee_data.keys():
                if employee_data[date]:
                    employee_data[date].sort()                   # Handle unsorted data in employee data file
                    date_key = (employee_name, date)
                    self._EMPLOYEES_DATA[date_key] = {
                        self._ENTRANCE: employee_data[date][0].strftime(self._TIME_FORMAT),
                        self._EXIT:     employee_data[date][-1].strftime(self._TIME_FORMAT),
                    }
                    employee_total = employee_data[date][-1] - employee_data[date][0]

                    if employee_total <= timedelta(minutes=45):  # Avoiding negative total time
                        employee_total = timedelta(minutes=0)
                    else:
                        employee_total = employee_total - timedelta(minutes=45)
                    employee_total_split = str(employee_total).split(":")
                    self._EMPLOYEES_DATA[date_key][self._TOTAL] = f"{employee_total_split[0]}:{employee_total_split[1]}"

    def _write_data_to_excel(self):
        """
        Creates an excel workbook and worksheet, then writes to the specified worksheet the titles of the table,
         the employees, creates table borders to it and closes the workbook.
        :return:
        :rtype:
        """
        if len(self._EMPLOYEES_DATA.keys()) > 0 and self._OUTPUT_PATH is not None:
            # Create Excel File, Worksheet and border format:
            self._WORKBOOK = xlsxwriter.Workbook(f"{self._OUTPUT_PATH}/{self._OUTPUT_FILE_NAME}")
            self._WORKSHEET = self._WORKBOOK.add_worksheet()
            reg_border_format = self._WORKBOOK.add_format({"border": 1})
    
            # Write Table titles
            row_counter = 1
            self._write_table_row(
                row_counter, self._NAME, self._DATE, self._ENTRANCE, self._EXIT, self._TOTAL, reg_border_format
            )
            row_counter += 1
            
            # Write employees data
            for key in self._EMPLOYEES_DATA.keys():
                employee_name, date = key
                self._write_table_row(
                    row_counter, employee_name, date, self._EMPLOYEES_DATA[key][self._ENTRANCE],
                    self._EMPLOYEES_DATA[key][self._EXIT], self._EMPLOYEES_DATA[key][self._TOTAL], reg_border_format
                )
                row_counter += 1
    
            # Draw outer border frame
            self._draw_outer_border_frame(1, 1, row_counter - 1, 5)
            self._WORKBOOK.close()

    def _write_table_row(self, row, name, date, entrance, exit, total, border_format):
        """
        Writes a single row in the excel table according to specified border_format in the specified row, the following
        values in increasing order: name, date, entrance, exit, total.
        """
        self._WORKSHEET.write(row, 1, name, border_format)
        self._WORKSHEET.write(row, 2, date, border_format)
        self._WORKSHEET.write(row, 3, entrance, border_format)
        self._WORKSHEET.write(row, 4, exit, border_format)
        self._WORKSHEET.write(row, 5, total, border_format)

    def _draw_border(self, row, col, rows_count, cols_count, direction):
        """
        A helper function which is used to help and draw the outer borders for the employees table. Is being used in
        reduce code duplication.
        """
        self._WORKSHEET.conditional_format(
            row, col, rows_count, cols_count,
            {
                "type": "formula",
                "criteria": "True",
                "format": self._WORKBOOK.add_format({direction: 5, "border_color": "#000000"})
            }
        )

    def _draw_outer_border_frame(self, first_row, first_col, rows_count, cols_count):
        """
        Draws the outer borders of the employees table.
        """
        # Top Left Corner
        self._draw_border(first_row - 1, first_col, first_row - 1, first_col, "bottom")
        self._draw_border(first_row, first_col - 1, first_row, first_col - 1, "right")
        # Top Right Corner
        self._draw_border(
            first_row - 1, first_col + cols_count - 1, first_row - 1, first_col + cols_count - 1, "bottom")
        self._draw_border(first_row, first_col + cols_count, first_row, first_col + cols_count, "left")
        # Bottom Left Corner
        self._draw_border(first_row + rows_count - 1, first_col - 1, first_row + rows_count - 1, first_col - 1, "right")
        self._draw_border(first_row + rows_count, first_col, first_row + rows_count, first_col, "top")
        # Bottom Right Corner
        self._draw_border(
            first_row + rows_count - 1, first_col + cols_count, first_row + rows_count - 1, first_col + cols_count,
            "left")
        self._draw_border(
            first_row + rows_count, first_col + cols_count - 1, first_row + rows_count, first_col + cols_count - 1,
            "top")

        # Top
        self._draw_border(first_row - 1, first_col + 1, first_row - 1, first_col + cols_count - 2, "bottom")
        # Left
        self._draw_border(first_row + 1, first_col - 1, first_row + rows_count - 2, first_col - 1, "right")
        # Bottom
        self._draw_border(
            first_row + rows_count, first_col + 1, first_row + rows_count, first_col + cols_count - 2, "top")
        # Right
        self._draw_border(
            first_row + 1, first_col + cols_count, first_row + rows_count - 2, first_col + cols_count, "left")

    def run(self):
        """
        Runs the employees report program.
        """
        if self._INPUT_PATH is not None:
            for file in os.listdir(self._INPUT_PATH):
                filename = os.fsdecode(file)
                if filename.endswith(".txt"):
                    employee_name = filename.split(".")[0].replace("_", " ")  # Extract employee name.
                    with open(f"{self._INPUT_PATH}/{filename}", "r") as f:
                        self._get_employee_data(employee_name, f.readlines())
            self._write_data_to_excel()


if __name__ == "__main__":
    input_path = None
    output_path = None
    if len(sys.argv) == 3:           # If input and output path is received by user via command line arguments
        input_path = sys.argv[1]
        output_path = sys.argv[2]
    employees_report_creator = EmployeesHours(input_path, output_path)
    employees_report_creator.run()
