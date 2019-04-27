import xlrd
import sys
import os
import re

def write_tcfile(lines, tcfiles, dict):
    for index in range(len(tcfiles)):
        print('   generating tc {}...'.format(tcfiles[index]))
        with open(tcfiles[index],'w') as tc_f:
            for line in lines:
                for key,val in dict.items():
                    if re.search(r'\b{}\b'.format(key),line):
                        if type(val) == list and len(val) > 1:
                            line = re.sub(r'\b{}\b'.format(key),val[index],line)
                        elif type(val) == list and len(val) == 1:
                            line = re.sub(r'\b{}\b'.format(key),val[0],line)
                        else:
                            line = re.sub(r'\b{}\b'.format(key),val,line)
                tc_f.write(line)

def rename_tcfile(tcfile, list_len):
    tcfiles = []
    for i in range(list_len):
        tcfile_list = list(tcfile)
        tcfile_list.insert(tcfile_list.index('.'),'-{}'.format(i))
        newfile = ''.join(tcfile_list)
        tcfiles.append(newfile)
    return tcfiles

def dict_process(dict):
    list_cnt = 0
    list_len = 0
    for key,val in dict.items():
        vals = []
        if re.search('^\[',val) and re.search('\]$',val):
            val = val[1:-1]
            if re.search(',',val):
                list_cnt = list_cnt + 1
                vals = val.split(',')
            elif re.search('-',val):
                list_cnt = list_cnt + 1
                tmp = val.split('-')
                if int(tmp[0]) >= int(tmp[-1])+1:
                    print('num in [a-b] must a <= b!')
                    sys.exit()
                for i in range(int(tmp[0]),int(tmp[-1])+1):
                    vals.append(str(i))
            else:
                vals.append(val)
            list_len = max(list_len, len(vals))
            dict[key] = vals
    if list_cnt > 1:
        print('[a,b,...] and [a-b] must only has one in one tc!')
        sys.exit()
    return dict, list_len

def read_basefie(basefile):
    if not os.path.isfile(basefile):
        print('Basefile {} is not a file!'.format(basefile))
        sys.exit()
    print('   basefile if {}'.format(basefile))
    with open(basefile,'r') as base_f:
        lines = base_f.readlines()
    return lines

def read_sheet(sheet):
    if sheet.ncols < 3 or sheet.nrows < 2:
        print('There is no tc in this sheet!')
        return -1
    if sheet.cell(0,0).ctype != 1: #string
        print('There is no base file in this sheet!')
        return -1
    dict = {}
    basefile = sheet.cell(0,0).value
    lines = read_basefie(basefile)
    for col in range(1,sheet.ncols):
        if col != 1:
            if sheet.cell(0,col).ctype != 1: #string
                print('There is no tc in this sheet cell(0,{})!'.format(col))
                return -1
            tcfile = sheet.cell(0,col).value
        for row in range(1,sheet.nrows):
            if col == 1:
                dict.setdefault(sheet.cell(row,col).value)
            else:
                if sheet.cell(row,col).ctype != 1: #string
                    print('Type of cell({},{}) must be string!'.format(row,col))
                    return -1
                else:
                    dict[sheet.cell(row,1).value] = sheet.cell(row,col).value
        if col != 1:
            dict, list_len = dict_process(dict)
            tcfiles = rename_tcfile(tcfile, list_len)
            write_tcfile(lines, tcfiles, dict)
    return 0

def read_excel(book_name):
    try:
        book = xlrd.open_workbook(book_name)
    except Exception:
        print('Cannot open excel {}!'.format(book_name))
        return
    else:
        print('Opening excel {}!'.format(book_name))
        for sheet_name in book.sheet_names():
            sheet = book.sheet_by_name(sheet_name)
            print('...Opening sheet {}!'.format(sheet_name))
            ret = read_sheet(sheet)
            if ret < 0:
                return

def usage(argv):
    print('Usage:\n\tpython {} <tc_config.xlsx>'.format(argv[0]))

def main(argv):
    if len(argv) != 2:
        usage(argv)
        return
    read_excel(argv[1])

if __name__ == '__main__':
    main(sys.argv)

                        
        
