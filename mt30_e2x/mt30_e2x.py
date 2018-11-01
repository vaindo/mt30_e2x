import xlrd
import sys
import operator as op
import os
import shutil

#定义需要用到的变量
argv_len = len(sys.argv)
sheet_name = 'Sheet1'
file_path = 'Tra_MT30_Firmware.xls'
refName_col = []
target_lang_col = []
target_col_name = 'TR'
lang_row = []

start_row_num = 0
end_row_num = -1

module_col = []
module_col_name = 'Module'

lang_col_name = "Lang"

dict_k_module_v_lang = {}

flag_classify = True

path_col_name = 'Path'
path_col = []

default_file_name = "strings.xml"
gen_dir = "gen_dir"

dict_k_module_v_path = {}


STR_HEAD = "    <string name=\""
STR_MID  = "\" >"
STR_TAIL = "</string>"
STR_TO_DEL = "S:Firmware:"

FILE_HEAD = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<resources>\n"
FILE_TAIL = "</resources>\n"


# 解析传入的参数
def input_para_analysis():
    global argv_len, sheet_name, file_path, target_col_name, start_row_num, end_row_num, flag_classify

    for i in range(1, argv_len):
        # print('i == %d' %i)
        # print('sys.argv[i]: '+ sys.argv[i])

        if op.eq(sys.argv[i], "sheet_name"):
            if i + 1 <= argv_len - 1:
                sheet_name = sys.argv[i + 1]

        if op.eq(sys.argv[i], "-p"):
            if i + 1 <= argv_len - 1:
                file_path = sys.argv[i + 1]

        if op.eq(sys.argv[i], "-l"):
            if i + 1 <= argv_len - 1:
                target_col_name = sys.argv[i + 1]

        if op.eq(sys.argv[i], "-sn"):
            if i + 1 <= argv_len - 1:
                start_row_num = int(sys.argv[i + 1])

        if op.eq(sys.argv[i], "-en"):
            if i + 1 <= argv_len - 1:
                end_row_num = int(sys.argv[i + 1])

        if op.eq(sys.argv[i], "-fc"):
            if i + 1 <= argv_len - 1:
                if 0 == int(sys.argv[i + 1]):
                    flag_classify = False
                else:
                    flag_classify = True

def read_my_excel():
    global refName_col, STR_TO_DEL, lang_row, target_col_name, target_lang_col, module_col_name, module_col
    global dict_k_module_v_path, path_col_name, dict_k_module_v_lang, path_col

    # 打开Excel
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_name(sheet_name)

    # 获取RefName列 (第一列)
    refName_col = sheet.col_values(0)
    refName_col.pop(0)

    # 编辑refName_col
    for i in range(0, len(refName_col)):
        refName_col[i] = refName_col[i].replace(STR_TO_DEL, '')

    # 获取语言行 (第一行)
    lang_row = sheet.row_values(0)

    # 获取 Lang列的内容,如果有,则使用此内容为目标语言
    for name in lang_row:
        if op.eq(name, lang_col_name):
            target_col_name = sheet.col_values(lang_row.index(lang_col_name))[1]

    # 获取目标语言列
    for col in lang_row:
        if op.eq(col, target_col_name):
            target_lang_col = sheet.col_values(lang_row.index(target_col_name))
    target_lang_col.pop(0)

    # 获取Module列
    for i in lang_row:
        if op.eq(i, module_col_name):
            module_col = sheet.col_values(lang_row.index(module_col_name))
    module_col.pop(0)

    # 获取Path列
    for i in lang_row:
        if op.eq(i, path_col_name):
            path_col = sheet.col_values(lang_row.index(path_col_name))
    path_col.pop(0)

    # 合成字典 Module 和  Path
    for i in range(0, len(module_col)):
        # 如果module id为空,虽然空字符串也可以成为字典的key, 但这里不存储
        key = module_col[i]

        if len(key) == 0 :
            continue
        dict_k_module_v_path[key] = path_col[i]

    # 合成字典 Module 和 language
    for i in range(0, len(module_col)):
        # 如果module id为空,虽然空字符串也可以成为字典的key, 但这里不存储
        key = module_col[i]
        if len(key) == 0 :
            continue

        # 多个模块的内容, 组合成目标字符串保存在字典里
        string = STR_HEAD + refName_col[i] + STR_MID + target_lang_col[i] + STR_TAIL + "\n"

        if key in dict_k_module_v_lang.keys():
            dict_k_module_v_lang[key] = dict_k_module_v_lang[key] + string
        else:
            dict_k_module_v_lang[key] = string

# 写文件
def write_file(path, file_name, content):
    curr_path = os.getcwd()
    if len(path) != 0:
        if not os.path.exists(path):
            os.makedirs(path)
            os.chdir(path)
    file = open(file_name, mode='a', encoding='utf-8')
    # file.write(FILE_HEAD)
    file.write(content)
    # file.write(FILE_TAIL)
    file.close()
    os.chdir(curr_path)

# 清理生成目录
def create_proj_dir():
    if os.path.exists(gen_dir):
        shutil.rmtree(gen_dir)
    os.makedirs(gen_dir)
    os.chdir(gen_dir)

# 保存文件
def save_my_xmls():

    create_proj_dir()

    for key in dict_k_module_v_path.keys():
        if len(dict_k_module_v_path[key]) != 0:
            # 添加文件头和文件尾
            string = FILE_HEAD + dict_k_module_v_lang[key] + FILE_TAIL
            write_file(dict_k_module_v_path[key], default_file_name, string)
        else:
            if os.path.exists(default_file_name) != True:
                # 添加文件头
                write_file("", default_file_name, FILE_HEAD)
            write_file("", default_file_name, dict_k_module_v_lang[key])
    # 添加文件尾
    write_file("", default_file_name, FILE_TAIL)



def main():

    print("This main .............!!!!! start")
    # 解析传入参数
    input_para_analysis()

    #读取Excel内容
    read_my_excel()

    # 保存文件
    save_my_xmls()

    print("This main .............!!!!! end")



# 程序从这里开始
main()
# 到这里结束


#打印变量
# print('Print the Var.====================================================================')
# print('sheet_name: ' + sheet_name)
# print('file_path: ' + file_path)
# print('refName_col: %s' % refName_col)
# print('target_col_name: %s' % target_col_name)
# print('target_lang_col: %s' % target_lang_col)
# print('lang_row: %s' % lang_row)
# print('start_row_num: %d' % start_row_num)
# print('end_row_num: %d' % end_row_num)
# print('module_col: %s' % module_col)
# print('flag_classify: %d' % flag_classify)
# print('dict_k_module_v_lang: %s' % dict_k_module_v_lang)
# print('dict_k_module_v_path: %s' % dict_k_module_v_path)

