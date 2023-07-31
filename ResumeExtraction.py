import os
import xlrd
import xlwt
from xlwt import XFStyle

def extract_parts_from_xls_files():
    # 获取当前目录
    current_directory = os.getcwd()

    # 创建新的Excel工作簿
    summary_workbook = xlwt.Workbook(encoding='utf-8')
    summary_sheet = summary_workbook.add_sheet('Summary')

    # 设置日期格式
    date_style = XFStyle()
    date_format = 'YYYY-MM-DD'
    date_style.num_format_str = date_format

    # 设置文本格式
    text_style = XFStyle()
    text_style.num_format_str = '@'

    # 写入表头
    summary_sheet.write(0, 0, '姓名')
    summary_sheet.write(0, 1, '出生年月')
    summary_sheet.write(0, 2, '性别')
    summary_sheet.write(0, 3, '籍贯')
    summary_sheet.write(0, 4, '学历')
    summary_sheet.write(0, 5, '学位和专业技术职务')
    summary_sheet.write(0, 6, '参加工作时间')
    summary_sheet.write(0, 7, '现单位及职务')
    summary_sheet.write(0, 8, '联系电话')
    summary_sheet.write(0, 9, '本人身份证号')
    summary_sheet.write(0, 10, '家庭地址')
    summary_sheet.write(0, 11, 'Excel文件名')

    row_index = 1

    # 遍历当前目录下的文件
    for filename in os.listdir(current_directory):
        if filename.endswith('.xls') and not (filename.startswith('.') or filename.startswith('~')):
            file_path = os.path.join(current_directory, filename)

            try:
                # 打开XLS文件
                workbook = xlrd.open_workbook(file_path)

                # 读取第一个工作表
                sheet = workbook.sheet_by_index(0)

                # 提取每个部分的值并写入汇总表
                name = sheet.cell_value(3, 2) if sheet.nrows > 3 and sheet.ncols > 2 else ''
                birth_date = sheet.cell_value(3, 6)
                if sheet.cell_type(3, 6) == xlrd.XL_CELL_DATE:
                    summary_sheet.write(row_index, 1, birth_date, date_style)
                else:
                    summary_sheet.write(row_index, 1, birth_date, text_style)

                gender = sheet.cell_value(3, 9) if sheet.nrows > 3 and sheet.ncols > 9 else ''
                hometown = sheet.cell_value(4, 2) if sheet.nrows > 4 and sheet.ncols > 2 else ''
                education = sheet.cell_value(6, 2) if sheet.nrows > 6 and sheet.ncols > 2 else ''
                degree_and_position = sheet.cell_value(6, 6) if sheet.nrows > 6 and sheet.ncols > 6 else ''
                work_start_date = sheet.cell_value(5, 6)
                if sheet.cell_type(5, 6) == xlrd.XL_CELL_DATE:
                    summary_sheet.write(row_index, 6, work_start_date, date_style)
                else:
                    summary_sheet.write(row_index, 6, work_start_date, text_style)

                current_position = sheet.cell_value(7, 2) if sheet.nrows > 7 and sheet.ncols > 2 else ''
                contact_number = sheet.cell_value(7, 10) if sheet.nrows > 7 and sheet.ncols > 10 else ''
                id_number = sheet.cell_value(38, 2) if sheet.nrows > 38 and sheet.ncols > 2 else ''
                home_address = sheet.cell_value(39, 2) if sheet.nrows > 39 and sheet.ncols > 2 else ''

                # 写入汇总表
                summary_sheet.write(row_index, 0, name)
                summary_sheet.write(row_index, 2, gender)
                summary_sheet.write(row_index, 3, hometown)
                summary_sheet.write(row_index, 4, education)
                summary_sheet.write(row_index, 5, degree_and_position)
                summary_sheet.write(row_index, 7, current_position)
                summary_sheet.write(row_index, 8, contact_number)
                summary_sheet.write(row_index, 9, id_number)
                summary_sheet.write(row_index, 10, home_address)
                summary_sheet.write(row_index, 11, filename)

                row_index += 1

            except xlrd.XLRDError:
                print(f"无法处理文件: {filename}，该文件可能不是有效的Excel文件或文件格式损坏。")
                continue

    # 保存汇总表到新的Excel文件
    summary_workbook.save('summary.xls')
    print("汇总表已生成：summary.xls")

# 调用函数提取每个部分的值并生成汇总表
extract_parts_from_xls_files()
