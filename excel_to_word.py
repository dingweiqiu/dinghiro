import win32com.client as win32
import os


def excel_to_docx(excel_file_path, docx_file_path):
    try:
        # 创建 Excel 应用程序对象
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        # 打开 Excel 文件
        workbook = excel.Workbooks.Open(os.path.abspath(excel_file_path))
        # 获取第一个工作表
        sheet = workbook.Sheets(1)
        # 获取数据范围
        used_range = sheet.UsedRange
        # 获取数据行数和列数
        rows, columns = used_range.Rows.Count, used_range.Columns.Count

        # 创建 Word 应用程序对象
        word = win32.gencache.EnsureDispatch('Word.Application')
        # 创建一个新的 Word 文档
        doc = word.Documents.Add()
        # 创建一个表格
        table = doc.Tables.Add(doc.Range(), rows, columns)

        # 填充表格数据
        for i in range(1, rows + 1):
            for j in range(1, columns + 1):
                table.Cell(i, j).Range.Text = str(used_range.Cells(i, j).Value)

        # 保存 Word 文档
        doc.SaveAs(os.path.abspath(docx_file_path))
        # 关闭 Word 文档
        doc.Close()
        # 退出 Word 应用程序
        word.Quit()
        # 关闭 Excel 工作簿
        workbook.Close()
        # 退出 Excel 应用程序
        excel.Quit()
        print(f"已成功将 {excel_file_path} 转换为 {docx_file_path}")
    except Exception as e:
        print(f"发生错误：{e}")


if __name__ == "__main__":
    excel_file = 'input.xlsx'
    docx_file = 'output.docx'
    excel_to_docx(excel_file, docx_file)
    