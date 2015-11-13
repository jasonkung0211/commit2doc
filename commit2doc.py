import os
import sys
from docx import *
import commit

def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)

def duplicate_row(table, row, n):
    cells = table.add_row().cells
    for i, cell in enumerate(row.cells):
        cells[i].text = cell.text + "#" + str(n) + "#"

def duplicate_row_times(table, row, n):
    for i in range(n):
        duplicate_row(table, row, i)
    remove_row(table, row)

# duplicate_row_when(document, "{row}")
def duplicate_row_when(document, when, n):
    for table in document.tables:
            for row in table.rows:
                needRows = False
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph_text = paragraph.text
                        if when in paragraph_text:
                            needRows = True
                            paragraph.text = paragraph_text.replace(when, '')
                            break
                    if needRows:
                        break
                if needRows:
                    duplicate_row_times(table, row, n)

#change working dir
def changeWD():
    #subprocess.call('cd C:/Users/User\PycharmProjects\commit2doc', shell=True)
    #workpath = os.getcwd()
    #subprocess.call(['cd', workpath], shell=True)
    os.chdir(os.getcwd())

def getargv():
    version = "1.0.00"
    help = '......(;==)'
    hash_id = "HEAD"
    if len(sys.argv) >= 2:
        if sys.argv[1].startswith('--'):
            option = sys.argv[1][2:] # 取出sys.argv[1]的數值但是忽略掉'--'這兩個字元
            if option == 'version':
                print('Version ' + version)
            elif option == 'help':
                print(help)
            sys.exit()
        else:
            hash_id = sys.argv[1]
    return hash_id


def get_dir_list(rootdir):
    lists = []
    for root, subFolders, files in os.walk(rootdir):
        for folder in subFolders:
            lists.append(folder)
        break
    return lists

def update_progress(progress):
    width = 65
    print('\r[{0}{1}] {2}%'.format('#'*int(progress*width), '='*(width-int(progress*width)), int(progress*100)), end='')

#------------------

c = commit.Commit(getargv())
###c.dump()

dir_list = get_dir_list(os.getcwd())

input_name = 'input.docx'
for item in dir_list:
    if "SKB" in item:
        input_name = 'skbtpl.docx'
        break
    if "ESUN" in item:
        input_name = 'esuntpl.docx'
        break
    if "YTBK" in item:
        input_name = 'ytbktpl.docx'
        break

document = Document(os.path.dirname(sys.argv[0]) + '/resource/' + input_name)
duplicate_row_when(document, "{row}", len(c.files))

for i, file in enumerate(c.files):
    # WTF '}'
    # last row will not be replaced if you use WTF '}'
    document.cell_replace('{commit.mod}' + '#' + str(i) + '#', c.mods[i])
    document.cell_replace('{commit.file_path}' + '#' + str(i) + '#', os.path.dirname(os.path.join(c.files[i])))
    document.cell_replace('{commit.author_date}' + '#' + str(i) + '#', c.author_date)
    document.cell_replace('{commit.author_name}' + '#' + str(i) + '#', c.author_name)
    document.cell_replace('{commit.file_name}' + '#' + str(i) + '#', os.path.basename(os.path.join(c.files[i])))
    document.cell_replace('{commit.seq}' + '#' + str(i) + '#', str(i+1))
    document.cell_replace('{commit.module}'+'#' + str(i) + '#', os.path.dirname(os.path.join(c.files[i])).split('/')[0])
    document.cell_replace('#' + str(i) + '#', '')
    update_progress(i/len(c.mods))


document.cell_replace('{commit.author_name}', c.author_name)
document.cell_replace('{commit.author_date}', c.author_date)
document.cell_replace('{commit.author_email}', ' <'+c.author_email + '>')
document.cell_replace('{commit.subject}', '     '+c.subject)
document.cell_replace('{commit.message}', '     '+c.message)
document.cell_replace('{commit.id}', c.id)

if (input_name == 'esuntpl.docx'):
    document.cell_replace('{commit.project_name}', '玉山應收帳款承購管理系統')
    document.cell_replace('{commit.project_id}', 'ESUNGCL')
if (input_name == 'skbtpl.docx'):
    document.cell_replace('{commit.project_name}', '新光應收帳款承購管理系統')
    document.cell_replace('{commit.project_id}', 'CSPSDB_PRD11222')
if (input_name == 'ytbktpl.docx'):
    document.cell_replace('{commit.project_name}', '元大應收帳款承購管理系統')
    document.cell_replace('{commit.project_id}', 'CSPSDEV1')

document.save('./' + 'commit' + '.docx')
update_progress(1)

