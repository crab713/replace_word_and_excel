from docx import Document
from xlrd import open_workbook # xlrd用于读取xld
from openpyxl import load_workbook
from xlutils.filter import process, XLRDReader, XLWTWriter
import os
import shutil
from win32com import client
import win32com
import pythoncom



def process_word(sourse_dir, save_dir):
    word = client.DispatchEx('Word.Application')
    word.Visible = 0
    word.DisplayAlerts = 0

    doc = word.Documents.Open(sourse_dir)
    for old in dic:
        word.Selection.Find.Execute(old, False, False, False, False, False, True, 1, True, dic[old], 2)
    doc.SaveAs(save_dir)
    doc.Close()
    word.Quit()


def process_excel(sourse_dir, save_dir):
    excel = client.DispatchEx('Excel.Application')
    excel.Visible = 0
    excel.DisplayAlerts = 0
    
    xlBook = excel.Workbooks.Open(sourse_dir)
    for old in dic:
        excel.ActiveSheet.UsedRange.Cells.Replace(old, dic[old])

    xlBook.SaveAs(save_dir)
    xlBook.Close()
    excel.Quit()

def run():
    for root, dirs, files in os.walk(SOURSE_FOLDER):
        save_root = SAVE_FOLDER + root.replace(SOURSE_FOLDER, '')
        if not os.path.exists(save_root):
            os.makedirs(save_root)
        try:
            for file in files:
                if file.split('.')[-1] == 'doc' or file.split('.')[-1] == 'docx':
                    process_word(os.path.join(root, file), os.path.join(save_root, file))
                if file.split('.')[-1] == 'xls' or file.split('.')[-1] == 'xlsx':
                    process_excel(os.path.join(root, file), os.path.join(save_root, file))
                else:
                    shutil.copy(os.path.join(root, file), os.path.join(save_root, file))
                print('文件%s完成修改'%os.path.join(root, file))  
        except Exception as e:
            print(str(Exception), end='\t')
            print(str(e))
            print('文件%s处理时出现异常，请手动修改或联系作者'%os.path.join(root, file))

# load args
SOURSE_FOLDER = input('输入需要处理的文件目录：\n')
print('开始录入替换文本，输入start结束输入，开始替换作业')
dic = dict()
old = ''
while old != 'start':
    old = input('输入需要替换的原字符串：')
    if old == 'start':
        break
    new = input('输入替换的新字符串：')
    dic[old] = new

print('开始作业')
SAVE_FOLDER = os.path.join(os.path.dirname(os.path.abspath(SOURSE_FOLDER)), 'save')

run()
print('完成')
os.system('pause')