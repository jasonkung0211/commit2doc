import os
import sys
from docx import *
import commit


def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)


def duplicate_row(table, row, n, c):
    cells = table.add_row().cells
    for i, cell in enumerate(row.cells):
        if '{commit.seq}' in cell.text:
            cells[i].text = cell.text.replace('{commit.seq}', str(n + 1))
            continue
        elif '{commit.module}' in cell.text:
            cells[i].text = cell.text.replace('{commit.module}',
                                              os.path.dirname(os.path.join(c.files[n])).split('/')[0])
            continue
        elif '{commit.file_path}' in cell.text:
            cells[i].text = cell.text.replace('{commit.file_path}', os.path.dirname(os.path.join(c.files[n])))
            continue
        elif '{commit.file_name}' in cell.text:
            cells[i].text = cell.text.replace('{commit.file_name}', os.path.basename(os.path.join(c.files[n])))
            continue
        elif '{commit.mod}' in cell.text:
            cells[i].text = cell.text.replace('{commit.mod}', c.mods[n])
            continue
        else:
            cells[i].text = cell.text


def duplicate_row_times(table, row, n, commit):
    for i in range(n):
        duplicate_row(table, row, i, commit)
        update_progress(i / n)
    remove_row(table, row)


# duplicate_row_when(document, "{row}")
def duplicate_row_when(doc, when, n, commit):
    for table in doc.tables:
        for row in table.rows:
            need_rows = False
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph_text = paragraph.text
                    if when in paragraph_text:
                        need_rows = True
                        paragraph.text = paragraph_text.replace(when, '')
                        break
                if need_rows:
                    break
            if need_rows:
                duplicate_row_times(table, row, n, commit)


def change_working_dir():
    # subprocess.call('cd C:/Users/User\PycharmProjects\commit2doc', shell=True)
    os.chdir(os.getcwd())


def getargv():
    version = "1.0.00"
    help = '......(;==)'
    hash_id = "HEAD"
    if len(sys.argv) >= 2:
        if sys.argv[1].startswith('--'):
            option = sys.argv[1][2:]  # 取出sys.argv[1]的數值但是忽略掉'--'這兩個字元
            if option == 'version':
                print('Version ' + version)
            elif option == 'help':
                print(help)
            sys.exit()
        else:
            hash_id = sys.argv[1]
    return hash_id


def get_dir_list(path):
    lists = []
    for root, subFolders, files in os.walk(path):
        for folder in subFolders:
            lists.append(folder)
        break
    return lists


def update_progress(progress):
    width = 65
    print('\r[{0}{1}] {2}%'.format('#' * int(progress * width), '=' * (width - int(progress * width)),
                                   int(progress * 100)), end='')


# ------------------

c = commit.Commit(getargv())
c.dump()
print(len(c.files))

dir_list = get_dir_list(os.getcwd())

input_file = 'input.docx'

for item in dir_list:
    if "SKB" in item:
        input_file = 'skbtpl.docx'
        break
    if "ESUN" in item:
        input_file = 'esuntpl.docx'
        break
    if "YTBK" in item:
        input_file = 'ytbktpl.docx'
        break

document = Document(os.path.dirname(sys.argv[0]) + '/resource/' + input_file)

if input_file in 'esuntpl.docx':
    document.cell_replace('{commit.project_name}', '玉山應收帳款承購管理系統')
    document.cell_replace('{commit.project_id}', 'ESUNGCL')
if input_file in 'skbtpl.docx':
    document.cell_replace('{commit.project_name}', '新光應收帳款承購管理系統')
    document.cell_replace('{commit.project_id}', 'CSPSDB_PRD11222')
if input_file in 'ytbktpl.docx':
    document.cell_replace('{commit.project_name}', '元大應收帳款承購管理系統')
    document.cell_replace('{commit.project_id}', 'CSPSDEV1')

document.cell_replace('{commit.author_name}', c.author_name.strip('\n'))
document.cell_replace('{commit.author_date}', c.author_date.strip('\n'))
document.cell_replace('{commit.author_email}', ' <' + c.author_email.strip('\n') + '>')
document.cell_replace('{commit.subject}', '     ' + c.subject.strip('\n'))
document.cell_replace('{commit.message}', '     ' + c.message.strip('\n'))
document.cell_replace('{commit.id}', c.id.strip('\n'))

duplicate_row_when(document, "{row}", len(c.files), c)
document.save('./' + 'commit' + '.docx')
update_progress(1)
