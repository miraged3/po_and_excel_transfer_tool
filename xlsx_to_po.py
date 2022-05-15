import os

import xlwings as xw


def write_line(content: str, txt, newline=False, quote=True):
    if content.endswith('\n'):
        content = content[:len(content) - 2]
        if quote:
            txt.write(content + '\"\n')
        else:
            txt.write(content + '\n')
        if newline:
            txt.write('\n')
    elif content.endswith('\\n'):
        if quote:
            txt.write(content + '\"\n')
        else:
            txt.write(content + '\n')
        if newline:
            txt.write('\n')
    else:
        if quote:
            txt.write(content + '\"\n')
        else:
            txt.write(content + '\n')
        if newline:
            txt.write('\n')


def isfloat(x):
    try:
        float(x)
    except ValueError:
        return False
    else:
        return True


if __name__ == '__main__':
    print('xlsx转po工具 by mirage')
    location = input('输入xlsx文件路径：')
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(location)
    sht = xw.sheets[0]
    length = len(sht.used_range.value) - 1
    print('共有' + str(length) + '行')
    new_filename = os.path.basename(location).rpartition('.')[0] + '.po'
    f = open(new_filename, 'a', encoding='utf-8')
    f.write("""msgid ""
    msgstr ""
    "MIME-Version: 1.0\\n"
    "Content-Type: text/plain; charset=UTF-8\\n"
    "Content-Transfer-Encoding: 8bit\\n"

    """)
    for i in range(2, length + 2):
        print('正在处理第' + str(i - 1) + '行')
        if sht.range('A' + str(i)).value is not None:
            location = str(sht.range('A' + str(i)).value)
            write_line('#: ' + location, f, quote=False)
        if sht.range('B' + str(i)).value is not None:
            msgctxt = sht.range('B' + str(i)).value
            if isfloat(msgctxt):
                msgctxt_int = int(float(msgctxt))
                msgctxt = msgctxt_int
            write_line('msgctxt \"' + str(msgctxt), f)
        if sht.range('C' + str(i)).value is not None:
            source = str(sht.range('C' + str(i)).value)
            if isfloat(source):
                source_int = int(float(source))
                source = source_int
            write_line('msgid \"' + str(source), f)
        if sht.range('D' + str(i)).value is not None:
            target = str(sht.range('D' + str(i)).value)
            if isfloat(target):
                target_int = int(float(target))
                target = target_int
            write_line('msgstr \"' + str(target), f, newline=True)
    f.close()
    wb.close()
    print(f'{location}处理完成')
    print(f'文件已保存至：{os.path.abspath(new_filename)}')
