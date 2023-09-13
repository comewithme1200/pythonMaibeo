import xlrd
import xlwt
from xlutils.copy import copy
import os




def get_file_and_rewrite(startpath):
    for root, dirs, files in os.walk(startpath):
        for dir in dirs:
            for root, dirs, files in os.walk(startpath + "\\" + dir):
                for file in files:
                    file_loc = root + "\\" + file
                    rb = xlrd.open_workbook(file_loc)
                    wb = copy(rb)
                    read_sheet = rb.sheet_by_index(1)
                    write_sheet = wb.get_sheet(0)
                    execute(read_sheet, write_sheet, file_loc, wb, int(dir))

def check(string):
    try:
        float(string)
        return True
    except ValueError:
        return False

def calTimes(percent, total):
    return round((percent/100)*total)

def execute(read_sheet, write_sheet, filename, wb, total):
    print(total)
    questions = []
    for x in range(read_sheet.ncols):
        answer = []
        for y in range(read_sheet.nrows):
            if check(read_sheet.cell_value(y, x)):
                if y==5 and (sum(answer) + calTimes(read_sheet.cell_value(y, x), total)) < total:
                    answer.append(total - sum(answer))
                else:
                    answer.append(calTimes(read_sheet.cell_value(y, x), total))
        if len(answer)>0:
            questions.append(answer)
    # print(questions)
    for i in range(len(questions)):
        for index, number in enumerate(questions[i]):
            for t in range(int(number)):
                if index==0:
                    row = t + 1
                elif index==1:
                    row = t + questions[i][index-1] + 1
                elif index == 2:
                    row = t + questions[i][index - 1] + questions[i][index - 2] + 1
                elif index == 3:
                    row = t + questions[i][index - 1] + questions[i][index - 2] + questions[i][index-3] + 1
                elif index == 4:
                    row = t + questions[i][index - 1] + questions[i][index - 2] + questions[i][index-3] + questions[i][index-4] + 1

                if int(row) <= total:
                    write_sheet.write(int(row), i, index+1)

    wb.save(filename)

if __name__ == '__main__':
    get_file_and_rewrite('C:\\Users\\Admin\\PycharmProjects\\pythonProject\\venv\\SỐ LIỆU')
