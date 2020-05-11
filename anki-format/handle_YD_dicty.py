import re

"""处理从有道词典导出的单词文本，返回标准文本，使得anki能够导入该文本
   但是在处理之前必须先用记事本将从有道导出的原始文本另存为utf-8格式，
   有道默认的是Unicode，程序将无法解码
   处理完成后在本地生成一个：有道处理后文本.txt"""

def strip_extra_line(file):
    """去除多余的换行符（仅针对有道词典格式）
       返回一个去除换行符后的文本"""
    print('去除多余换行符...')
    try:
        with open(file, 'r', encoding='utf-8') as f:
            contents = f.read()
    except UnicodeDecodeError:
        print('请先将原始文件转化为 utf-8 的文本文档')
    pattern = re.compile(r'\n{2,3}')
    lines = pattern.findall(contents)
    for each in lines:
        contents = contents.replace(each, '\n')
    return contents

def add_char(text):
    """进一步处理文本，将单词前面的数字去除；
       在单词与释义之间添加分隔符；
       返回一个文本列表，每一个元素为一个单词及其释义"""
    pattern = re.compile(r'\d{1,4},')
    marks = pattern.findall(text)

    print('添加分割标志...')
    marks.reverse()
    for each in marks:
        text = text.replace(each, '\t')

    # 将文本处理为列表
    text_list = text.split('\t')
    pattern2 = re.compile(r'\n')
    text_list.pop(0)
    for each in range(len(text_list)):
        if ']' in text_list[each]:
            text_list[each] = text_list[each].replace(']', ']\t')
        else:
            target = pattern2.search(text_list[each])
            #print(text_list[each],'..........')
            print(target)
            if target:
                num = target.span()[0]
                text_list[each] = text_list[each][:num] +'\t'+ text_list[each][num:]
    print(len(text_list))
    return text_list

def final_process(contents):
    """将contents文本列表转化为anki可以处理的文本格式，并将其返回"""
    for i in range(len(contents)):
        contents[i] = contents[i].replace('\n','<div>').strip() + '\n'
    return contents

if __name__ == '__main__':
    file = input('输入文件名(请务必先将格式转化为utf-8)：')
    text = strip_extra_line(file)
    contents = add_char(text)
    target_contents = final_process(contents)
    print('写入文件...')
    with open('有道处理后文本.txt','w',encoding='utf-8') as f:
        f.writelines(target_contents)
