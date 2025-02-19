
import win32com.client as win32
 
def read_word_tables_with_win32(word,file_path):
    
    doc = word.Documents.Open(file_path)
    tables_content = []
 
    for table in doc.Tables:
        rows_content = {}
        for cell in table.Range.Cells:
            row_idx = cell.RowIndex
            cell_text = cell.Range.Text.strip().replace('\r', '').replace('\x07', '')
            if row_idx not in rows_content:
                rows_content[row_idx] = []
            rows_content[row_idx].append(cell_text)
 
        # Join each row's cell contents into a single string
        data = [' '.join(rows_content[row]) for row in sorted(rows_content.keys())]
        tables_content.append(data)
    doc.Close()
   
    return tables_content

word = win32.Dispatch('Word.Application')
tables = read_word_tables_with_win32(word, r'F:\Github\ExportWord\2024年5月份培训考核分析.docx')
tables = read_word_tables_with_win32(word, r'F:\Github\ExportWord\2024年6月份培训考核分析.docx')
for i, table in enumerate(tables):
    print(f'Table {i}:\n', table)
word.Quit()