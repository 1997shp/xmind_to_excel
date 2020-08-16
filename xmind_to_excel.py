from xmindparser import xmind_to_dict
import re
from openpyxl import Workbook





def iter_x(x):
    if isinstance(x, dict):
        for key, value in x.items():
            yield (key, value)
    elif isinstance(x, list):
        for index, value in enumerate(x):
            yield (index, value)





def flat(x):
    for key, value in iter_x(x):
        if isinstance(value, (dict, list)):
            for k, v in flat(value):
                k = f'{key}_{k}'
                yield (k, v)
        else:
            yield (key, value)





list4 = []
def xmind_list(out):
    e = {}
    list1 = []
    for i in out:

        a = {k: v for k,v in flat(i)}

        for b in a:
            pattern = re.compile(r'(?<=_)\d*')
            y = pattern.findall(b)
            y = list(filter(None, y))
            if len(y) != 0:
                y1 = [str(i) for i in y]
                y2 = '_'.join(y1)

                y3 = y2 + '_'

                e[y3] = a[b]
                list1.append(y3)
            else:
                e['title'] = a[b]


    list2 = []
    list3 = []
    for i in list1:
        o = 0
        for a in list1:

            y = re.match(str(i), str(a))
            if y != None:
                o = o + 1


        if o == 1:
            list2.append(i)
        else:
            list3.append(i)




    for i in list2:
        p = e['title']
        list5 = []
        k = len(i)
        j = 2
        list5.append(p)

        while j <= k:

            u = i[0:j]
            for a in list3:
                y = None
                if len(a) == len(u):
                    y = re.match(str(u), str(a))
                    if y != None:

                        list5.append(e[a])
            j = j + 2
        list5.append(e[i])
        list5 = [item for item in list5 if list5.count(item) == 1]

        list4.append(list5)



if __name__ == "__main__":

    out = xmind_to_dict('3.xmind')  # 将此处名称换为想要转换的xmind的名称

    out = out[0]['topic']['topics']


    # 如果只有一个二级标题的情形下，即测试xmind文件中的“需求名”这一层只有一个的情况下，下方for循环可以注释掉 直接xmind_list(out)
    for i in out:
        list9 = []
        list9.append(i)
        xmind_list(list9)


    workbook = Workbook()
    save_file = "写入文件.xlsx"
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    for row in list4:
        worksheet.append(row) # 把每一行append到worksheet中
    workbook.save(filename=save_file) 
