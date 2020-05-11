import re
import time
"""处理从欧路词典导出的单词文本，返回标准文本，使得anki能够导入该文本"""

def strip_extra_line(file):
    print('去除多余换行符...')
    with open(file, 'r', encoding='utf-8') as f:
        contents = f.read()
    pattern = re.compile(r'\n[\n]*\n')
    lines = pattern.findall(contents)
    lines = sorted(lines, reverse=True)
    for each in lines:
        contents = contents.replace(each, '\n')
    return contents

def add_char(text):
    pattern = re.compile(r'\d{1,3}@')
    marks = pattern.findall(text)
    print('添加分割标志...')
    marks.reverse()
    for each in marks:
        text = text.replace(each, '\t')

    # 将文本处理为列表
    text_list = text.split('\t')
    print(len(text_list))
    count = 0
    for i in range(len(text_list)):
        if '@' in text_list[i]:

            text_list[i] = text_list[i].replace('@','\t')
            count += 1
        else:
            print('没找到：@')

    text_list.pop(0)
    return text_list

def final_process(contents):
    words_dict = {}
    for each in contents:
        each = each.split('\t')
        each[1] = each[1].replace('\n','<div>')
        words_dict[each[0]] = each[1]
    words_list = []
    for k, v in words_dict.items():
        line = k + '\t' + v + '\n'
        words_list.append(line)
    return words_list

if __name__ == "__main__":
    file = input('请输入要转化的欧陆词典单词文本路径：')
    text = strip_extra_line(file)
    contents = add_char(text)
    target_contents = final_process(contents)
    print('写入文件...')
    with open('处理后单词文本.txt','w',encoding='utf-8') as f:
        f.writelines(target_contents)
