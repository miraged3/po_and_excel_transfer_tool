import os.path

import xlwings as xw

if __name__ == '__main__':
    print('po转xlsx工具 by mirage')
    location = input("输入po文件路径：")
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()
    sht = wb.sheets["sheet1"]
    current_line = 2
    sht.range('A1').value = "location"
    sht.range('B1').value = "msgctxt"
    sht.range('C1').value = "source"
    sht.range('D1').value = "target"
    f = open(location, 'r', encoding='utf-8')
    lines = f.readlines()
    is_on_location = False
    for line in lines:
        if line.startswith('#: '):
            print('正在处理第' + str(current_line) + '行')
            location = line.split('#: ')[1]
            sht.range(f'A{current_line}').value = location
            is_on_location = True
        elif line.startswith('msgctxt ') and is_on_location:
            msgctxt1 = line.split('msgctxt ')[1]
            msgctxt = msgctxt1[1:len(msgctxt1) - 2]
            if msgctxt.startswith('\''):
                msgctxt = '\'' + msgctxt
            sht.range(f'B{current_line}').value = msgctxt
        elif line.startswith('msgid ') and is_on_location:
            source1 = line.split('msgid ')[1]
            source = source1[1:len(source1) - 2]
            if source.startswith('\''):
                source = '\'' + source
            sht.range(f'C{current_line}').value = source
        elif line.startswith('msgstr ') and is_on_location:
            target1 = line.split('msgstr ')[1]
            target = target1[1:len(target1) - 2]
            if target.startswith('\''):
                target = '\'' + target
            sht.range(f'D{current_line}').value = target
            current_line = current_line + 1
            is_on_location = False
    wb.save(os.path.basename(location).rpartition('.')[0] + '.xlsx')
    wb.close()
    print(f'文件已保存至{os.path.abspath(os.path.basename(location).rpartition(".")[0] + ".xlsx")}')
