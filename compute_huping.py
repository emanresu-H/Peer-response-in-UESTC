import pandas as pd
import os
import sys

path_now = os.path.dirname(os.path.realpath(sys.argv[0])) #获取当前文件所在路径
path = path_now + "/互评Excel文件"

if __name__ == '__main__':
    # 读取文件名（评价的同学名字）
    filenames = os.listdir(path)
    Name = []  # 评价的同学名字
    file_order = 0
    for filename in filenames:
        if os.path.splitext(filename)[1] == '.xlsx':  # 判断是否为xlsx文件
            # 提取姓名
            name = os.path.splitext(filename)[0]
            Name.append(name)

    read_name = path_now + "/互评Excel文件/" + filenames[0]
    file_init = pd.read_excel(read_name)

    # 初始化一个表的雏形
    file_sxpd = file_init[['姓名', '思想品德素质得分']]
    file_sxpd = file_sxpd.rename(columns={'思想品德素质得分': Name[0]})
    file_sxsm = file_init[['姓名', '身心+审美人文+劳动素质得分']]
    file_sxsm = file_sxsm.rename(columns={'身心+审美人文+劳动素质得分': Name[0]})

    # 读入剩余文件
    file_num = len(filenames)  # 文件数
    for i in range(1, file_num):
        read_name = path_now + "/互评Excel文件/" + filenames[i]
        file_read = pd.read_excel(read_name)  # 读入文件
        file_sxpd[Name[i]] = file_read['思想品德素质得分']
        file_sxsm[Name[i]] = file_read['身心+审美人文+劳动素质得分']

    # 求裁剪平均值
    file_sxpd['average'] = (file_sxpd.sum(axis=1) - file_sxpd.min(axis=1) - file_sxpd.max(axis=1)) / (file_num - 2)
    file_sxsm['average'] = (file_sxsm.sum(axis=1) - file_sxsm.min(axis=1) - file_sxsm.max(axis=1)) / (file_num - 2)

    #排序
    file_sxpd = file_sxpd.sort_values(by="average", ascending=False)
    file_sxsm = file_sxsm.sort_values(by="average", ascending=False)

    print('思想品德素质得分：\n',file_sxpd)
    print('\n身心+审美人文+劳动素质得分：\n',file_sxsm)

    file_sxpd.to_excel('思想品德.xlsx', encoding='utf_8_sig')
    file_sxsm.to_excel('身心+审美人文+劳动.xlsx', encoding='utf_8_sig')
