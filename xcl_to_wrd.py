import openpyxl as op
from docx import Document


def get_data(file_name, min_row, max_row):
   """
   функция принимает имя файла, номера строк которые нужно обработать,
   и возвращает список
   :param file_name: передается имя файла (с расширением) для прочтения
   :param min_row: строка с которой начинается чтение
   :param max_row: строка на которой чтение файла заканчивается
   :return:
   """
   data_list = []
   sheets = op.load_workbook(filename=file_name).active
   for rows in sheets.iter_rows(min_row=min_row, max_row=max_row, min_col=2, max_col=5):
       x = []
       for cell in rows:
           x.append(cell.value)
       data_list.append(x)
   return data_list


# задаем имя файлов в качвычках и нужные нам строки min_row и max_row
source_list = get_data('source.xlsx', 5, 25)
document = Document()
table_2 = document.add_table(rows=1, cols=3)

for pc_name, hdd_number, hdd_serial, pc_number in source_list:
   row_cells_2 = table_2.add_row().cells
   row_cells_2[0].merge(row_cells_2[-1])
   row_cells_2[0].text = f'АРМ с именем «{str(pc_name)}» в составе:'
   row_cells_2 = table_2.add_row().cells
   row_cells_2[0].text = str('Системный блок')
   row_cells_2[1].text = str('-')
   if pc_number is None:
      row_cells_2[2].text = str(f'-')
   else:
      row_cells_2[2].text = str(f'{pc_number}')
   row_cells_2 = table_2.add_row().cells
   if hdd_number is None:
      row_cells_2[0].text = str('-')
   else:
      row_cells_2[0].text = str(f'ЖМД № {hdd_number}')
   row_cells_2[1].text = str('-')
   if hdd_serial is None:
      row_cells_2[2].text = str('-')
   else:
      row_cells_2[2].text = str(hdd_serial)
   row_cells_2 = table_2.add_row().cells
   row_cells_2[0].merge(row_cells_2[-1])
   row_cells_2[0].text = f'Периферийное оборудование: монитор, клавиатура, манипулятор мышь, принтер,акустические колонки, web-камера '


document.save('check.docx')
