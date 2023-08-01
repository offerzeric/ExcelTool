import pandas as pd
import  numpy as np
#解决显示不全
pd.set_option('display.max_columns',5000)
pd.set_option('display.width', 5000)
pd.set_option('display.max_colwidth',None)

def read_excel(path, SheetName,tempalte_data,tempalte_SheetName):
    print("处理%s" % SheetName)
    global DATA, TEMPLATE_DATA, max_row
    pd.options.mode.use_inf_as_na = True
    DATA = pd.read_excel(path, sheet_name=SheetName, index_col=None, header=None)
    TEMPLATE_DATA = pd.read_excel(tempalte_data,sheet_name=tempalte_SheetName, index_col=None, header=None)

    max_row = DATA.shape[0]
    print(max_row)
    #在TEMPLATE_DATA DataFrame的基础上增加了填充了NaN值的行，总行数达到max_row。列数保持不变。
    TEMPLATE_DATA = pd.concat([TEMPLATE_DATA, pd.DataFrame([[pd.NA]*TEMPLATE_DATA.shape[1]]*(max_row-2))], ignore_index=True)

    return DATA, TEMPLATE_DATA, max_row

def read_enroute_excel(path):
    global D__data, DB_data, PA_data, EA_data
    D__data = pd.read_excel(path, sheet_name='D_',index_col=None, header=None)
    DB_data = pd.read_excel(path, sheet_name='DB',index_col=None, header=None)
    PA_data = pd.read_excel(path, sheet_name='PA',index_col=None, header=None)
    EA_data = pd.read_excel(path, sheet_name='EA', index_col=None, header=None)
    return D__data, DB_data, PA_data, EA_data
#按已填好的数据自动填充空白单元格
def padding_sameValue(dataframe,clo):
    clo_data = dataframe[0:max_row][clo]
    print(clo_data)
    # 获取非空的索引值，以列表输出。
    notnull_index = []
    for i in range(0,int(max_row)):
        print(clo_data[i],i)
        if pd.isna(clo_data[i]) is False:
            notnull_index.append(i)
    notnull_index.append(max_row)
    print(notnull_index)
    #计算出非空值间的间隔值，以列表输出
    notnull_index_interval =[]
    for j in range(0,len(notnull_index)-1):
        notnull_index_interval.append(notnull_index[j+1]-notnull_index[j])
    print(notnull_index_interval)
    for k in range(0,len(notnull_index_interval)):
        #筛选出无间隔的数据
        if notnull_index_interval[k] == 1:
            continue
        else:
            #将间隔中的数据补充
            for l in range(1,notnull_index_interval[k]):
                print(l)
                clo_data[notnull_index[k]+l]=clo_data[notnull_index[k]]
    print(clo_data)

    print(dataframe[:max_row][clo])
    return dataframe

#PD记录P13的ICAO代码判断
'''
Ref_list: 此列表中的记录号不需要查找其他表格，索引自身数据即可
da_ref_clo: 主数据的参考列的索引
da_target_clo:主数据的目标更改列索引
r_da_target_clo:主数据的目标参考列索引
'''
def PD_P13_ICAO(dataframe, Ref_list, da_ref_clo, da_target_clo, r_da_target_clo ):
    for i in range(max_row):
        if pd.isna(dataframe[12][i]) is False:
            if dataframe[da_ref_clo][i] in Ref_list:
                dataframe[da_target_clo][i] = dataframe[r_da_target_clo][i]
            elif dataframe[da_ref_clo][i] == 'D':
                get_same_ICAO(dataframe, D__data, 12, 7, 13, 9, i)
            elif dataframe[da_ref_clo][i] == 'B':
                get_same_ICAO(dataframe, DB_data, 12, 7, 13, 9, i)
            elif dataframe[da_ref_clo][i] == 'E':
                get_same_ICAO(dataframe, EA_data, 12, 7, 13, 9, i)
            elif dataframe[da_ref_clo][i] == 'A':
                get_same_ICAO(dataframe, PA_data, 12, 4, 13, 5, i)
            else:
                dataframe[da_ref_clo][i] = '人工编码数据有误'
        else:
            continue

#PD记录P26的ICAO代码判断
def PD_P26_ICAO(dataframe, Ref_list, da_ref_clo, da_target_clo, r_da_target_clo ):
    for i in range(max_row):
        if pd.isna(dataframe[25][i]) is False:
            if dataframe[da_ref_clo][i] in Ref_list:
                dataframe[da_target_clo][i] = dataframe[r_da_target_clo][i]
            elif dataframe[da_ref_clo][i] == 'D':
                get_same_ICAO(dataframe, D__data, 25, 7, 26, 9, i)
            elif dataframe[da_ref_clo][i] == 'B':
                get_same_ICAO(dataframe, DB_data, 25, 7, 26, 9, i)
            else:
                dataframe[da_ref_clo][i] = '人工编码数据有误'
        else:
            continue

'''
ref_dataframe: 主数据所需关联的参考数据
da_match_index: 主数据中用于匹配的列索引
r_da_match_index:参考数据中用于匹配的列索引
da_target_index:主数据的目标更改列索引
r_da_target_index：将值传递给主数据的目标更改列的参考列索引
'''
def get_same_ICAO(dataframe, ref_dataframe, da_match_index, r_da_match_index, da_target_index, r_da_target_index, i):
    indices = [j for j, x in enumerate(ref_dataframe[r_da_match_index][:]) if ref_dataframe[r_da_match_index][j] == dataframe[da_match_index][i]]

    if indices:
        dataframe[da_target_index][i] = ref_dataframe[r_da_target_index][indices[0]]
    else:
        dataframe[da_target_index][i] = '未匹配到数据'

#自动填充Sequence_Number
def Sequence_Number(dataframe):
    #机场四码、程序名、过渡名相同时判断重复行的索引值
    df = dataframe[dataframe.duplicated(keep=False,subset=dataframe.columns[[4,7,9]])]

    # 对获取出来相同数值的行索引，用元组组合，并使用列表将元组
    df = df.groupby(dataframe.columns[[4,7,9]].tolist()).apply(lambda x: tuple(x.index)).tolist()
    print(df)
    #生成序列号，并替换原数据
    for i in range(len(df)):
        for j in range(len(df[i])):
            Sequence_Number=str((df[i][j]-df[i][0]+1) * 10).zfill(3)
            print(Sequence_Number)
            dataframe[11][df[i][j]]=Sequence_Number
    print(dataframe)
    return dataframe

# 根据输入的一个字符匹配完整的记录名
def section_sub_code(dataframe, Ref, Sec, Subsec):
    Dic_Sec_Sub = {'D': ['D',' '], 'C': ['P', 'C'], 'G': ['P', 'G'], 'N': ['P', 'N'], 'B': ['D', 'B'], 'I': ['P', 'I'], 'A': ['P', 'A'],
             'E': ['E', 'A'], 'M': ['P', 'M']}
    for i in range(1, max_row):
        print(i)
        # 使用字典键值对匹配
        match_data = [value1 for (key1, value1) in Dic_Sec_Sub.items() if key1 == dataframe[Ref][i]]
        if len(match_data):
            print(match_data)
            dataframe[Sec][i],  dataframe[Subsec][i] = match_data[0][0], match_data[0][1]
        else:
            print(match_data)
            dataframe[Sec][i], dataframe[Subsec][i] = ' ', ' '
# 航路点类型第一位字符判断
def waypoint_code_1(dataframe, Ref, code):
    Dic_way_code1 = {'D': 'V', 'C': 'E', 'G': 'G', 'N': 'N', 'B': 'N', 'I': 'V', 'A': 'A', 'E': 'E'}
    for i in range(1, max_row):
        print(i)
        match_data = [value1 for (key1, value1) in Dic_way_code1.items() if key1 == dataframe[Ref][i]]
        if len(match_data):
            print(match_data)
            dataframe[code][i] = match_data[0]
        else:
            print(match_data)
            dataframe[code][i] = ' '

# 航路点类型第二位字符判断
def waypoint_code_2(dataframe, Ref1, Ref2, code):
    for i in range(1, max_row-1):
        if pd.isna(dataframe[Ref1][i]) is True:
            if int(dataframe[Ref2][i]) < int(dataframe[Ref2][i+1]) :
                dataframe[code][i] = 'E'
            elif int(dataframe[Ref2][i]) > int(dataframe[Ref2][i+1]):
                dataframe[code][i] = ' '
        elif dataframe[Ref1][i] == 'Y':
            if int(dataframe[Ref2][i]) < int(dataframe[Ref2][i+1]) :
                dataframe[code][i] = 'B'
            elif int(dataframe[Ref2][i]) > int(dataframe[Ref2][i+1]):
                dataframe[code][i] = 'Y'
    #最大行的判断
    if pd.isna(dataframe[Ref1][max_row-1]) is True:
        dataframe[code][max_row-1] = 'E'
    elif dataframe[Ref1][max_row-1] == 'Y':
        dataframe[code][max_row-1] = 'B'

# 航路点类型第三位字符判断
def waypoint_code_4(dataframe, Ref, code):
    for i in range(1, max_row):
        if pd.isna(dataframe[Ref][i]) is True:
            dataframe[code][i] = ' '
        elif dataframe[Ref][i] == 'Y':
            dataframe[code][i] = 'H'
        else:
            dataframe[code][i] = '人工编码数据有误'

def Turn_Direction_Valid(dataframe, Ref1, Ref2, target):
    PT_coed = ['AF', 'DF', 'HA', 'HF', 'HM', 'IF', 'PI', 'RF']
    for i in range(1, max_row):
        if pd.isna(dataframe[Ref2][i]) is True:
            dataframe[target][i] = ' '
        elif dataframe[Ref2][i] == 'L' or 'R':
            if dataframe[Ref1][i] in PT_coed:
                dataframe[target][i] = ' '
            else:
                dataframe[target][i] = 'Y'
        else:
            dataframe[target][i] = '转弯方向编码数据有误'



if __name__ == '__main__':
    path = r'D:\8-CodeFile\PycharmProjects\424FastWriteTool\424test.xlsx'
    template_path = r'D:\8-CodeFile\PycharmProjects\424FastWriteTool\424template.xlsx'
    Enroute_path = r'D:\8-CodeFile\PycharmProjects\424FastWriteTool\ZUKD航路部分.xlsx'
    read_excel(path,'Sheet1',template_path,'SID')

    TEMPLATE_DATA.iloc[:, [4, 39, 7, 8, 9, 23, 12, 14, 17, 18, 19, 20, 30, 21, 35, 37, 38, 40, 22, 25, 32,
                           28, 29, 31, 27, 42, 45, 53, 54, 55]] = DATA.iloc[:, :]
    print(TEMPLATE_DATA)

    padding_TEMPLATE_DATA_idnex = [0, 1, 2, 3, 4, 6, 7, 8, 9, 10, 16, 34, 36, 39, 41, 43, 47, 48, 49, 50, 51, 52, 53, 54,
                                   55, 56]
    for j in range(0, len(padding_TEMPLATE_DATA_idnex)):
        padding_sameValue(TEMPLATE_DATA, padding_TEMPLATE_DATA_idnex[j])
    print(TEMPLATE_DATA)

    read_enroute_excel(Enroute_path)

    PD_P13_Ref_list = ['C', 'G', 'N', 'I', 'M']
    PD_P13_ICAO(TEMPLATE_DATA, PD_P13_Ref_list, 14, 13, 5)

    PD_P26_Ref_list = [ 'N', 'I', 'M']
    PD_P26_ICAO(TEMPLATE_DATA, PD_P13_Ref_list, 32, 26, 5)

    Sequence_Number(TEMPLATE_DATA)
    print(TEMPLATE_DATA)
    section_sub_code(TEMPLATE_DATA, 14, 14, 15)
    print('wancheng')
    section_sub_code(TEMPLATE_DATA, 45, 45, 46)
    print(TEMPLATE_DATA)
    waypoint_code_1(TEMPLATE_DATA, 17, 17)
    waypoint_code_4(TEMPLATE_DATA, 20, 20)
    waypoint_code_2(TEMPLATE_DATA, 18, 11, 18)
    Turn_Direction_Valid(TEMPLATE_DATA, 23, 21, 24)
    TEMPLATE_DATA.to_excel('example.xlsx', index=False)
