import xlrd,re,csv
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

def filePath(path):
    if path == 0:    # 获取文件路径
        f_path = filedialog.askopenfilename()
        print('\n获取的文件地址：', f_path)
        return f_path
    elif path == 1:  #获取文件夹路径
        d_path = filedialog.askdirectory()
        print('\n获取的文件地址：', d_path)
        return d_path
    
def xlsProcess(path):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)#选择工作表
    CSV_data = [["课程名称","星期","开始节数","结束节数","老师","地点","周数"]]
    CSV_lession = []
    cell_data =""
    colNum = 2  #从第三列开始
    while colNum <= 8:
        rowNum = 2  #从第三行开始
        while rowNum <=6:
            cell_data = sheet.cell_value(rowNum,colNum)
            #对一个单元格内的数据进行拆分并保存至CSV_data
            #如：习近平新时代中国特色社会主义思想概论/(1-2节)4-5周,7-8周,10-11周,13-15周/南校区 N403/牛鲁玉/无/(2024-2025-1)-BK106013-41/植保2401;植保2402;植保2403;植保2404;植保2405/无/150/0/未安排/无/讲课:34,讨论:14/36/3.0
            #拆为列表["习近平新时代中国特色社会主义思想概论","(1-2节)4-5周,7-8周,10-11周,13-15周","南校区 N403","牛鲁玉"]
            #再转化为：["习近平新时代中国特色社会主义思想概论",f"{colNum-1}","1","2","牛鲁玉","南校区 N403","4-5、7-8、10-11、13-15"]
            print(len(cell_data))
            list_cell = cell_data.split("/")
            print("list:",len(list_cell))
            if len(cell_data) != 0 and len(list_cell) == 1:  #如果当前单元格非空且仅有课程名
                CSV_lession = [list_cell[0],f"{colNum-1}",f"{(rowNum-2)*2+1}",f"{(rowNum-2)*2+2}","","","1-20"] # type: ignore
                CSV_data.append(CSV_lession)
            else:
                #周数
                try:  
                    week1 = list_cell[1].split(")")   #该行遇空单元格会抛出IndexError错误
                    week2 = week1[1].replace("周","")
                    week = week2.replace(",","、")
                    #节数
                    session = re.findall(r"[0-9]{1,2}",week1[0])
                    CSV_lession = [list_cell[0],f"{colNum-1}",session[0],session[1],list_cell[3],list_cell[2],week]
                    CSV_data.append(CSV_lession)
                except IndexError:  #捕捉IndexError错误并pass
                    pass
            rowNum += 1
        colNum += 1
    print(CSV_data)
    return CSV_data

def CSVwriter(data): #打开对话框选择文件夹并保存至output.csv
    with open(f'{filePath(1)}/output.csv', 'w', newline='' ,encoding= 'utf-8') as file:
        writer = csv.writer(file)
        for row in data:
            writer.writerow(row)

path = filePath(0)

if path:
    matchs = re.search(r".课表.",path) #防止意外导入非课表文件
    if matchs == None:
        print("请选择课程表文件！")
    else:
        CSV_data =xlsProcess(path)
        CSVwriter(CSV_data)
        print("Done")
