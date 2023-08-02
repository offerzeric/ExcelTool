import stat
from flask import (Blueprint, request, json, current_app, abort, make_response, send_from_directory)
from openpyxl import Workbook
# from openpyxl import Workbook

from werkzeug.utils import secure_filename
import pandas as pd
import numpy as np
import uuid
import os, sys
import shutil
import math
import collections
import datetime


"""
    解决显示不全
"""
pd.set_option('display.max_columns',5000)
pd.set_option('display.width', 5000)
pd.set_option('display.max_colwidth',None)
#create a bp
bp = Blueprint('code', __name__, url_prefix='/code')

"""
    424编码controller层
"""
"""
    上传源文件接口

"""
@bp.route('/do424Upload',methods=['POST'])
def do_424_Upload():
    if request.method == 'POST':
        current_app.logger.debug("Getting upload files.")
        #获取source_xlsx和result_xlsx的路径
        source_path = os.path.abspath('source_xlsx')
        result_path = os.path.abspath('result_xlsx')
        current_app.logger.debug(source_path);
        current_app.logger.debug(result_path);
        #首先删除源文件文件夹和结果文件夹避免上一次遗留文件的影响
        remove_file_or_dir(source_path);
        remove_file_or_dir(result_path);
        #获取上传的所有excel
        files = request.files.getlist('sourceSheetFiles');
        all_xlsx_addrs = []
        if request.files.get("sourceSheetFiles").filename == '':           
            current_app.logger.debug("没有收到要上传的文件.")
            result = {
                "flag" :0,
                "reason":'没有收到要上传的文件.'
            }
            current_app.logger.debug("上传失败: " + str(result))
            #返回错误情况下的结果
            resp = json.dumps(result);
            res = make_response(resp);
            res.headers["Access-Control-Allow-Origin"] = "*";
            return res;
        for file in files:
            original_filename = (file.filename)
            unique_filename = (original_filename)
            #保存每个excel文件
            file.save(os.path.join(source_path, unique_filename))
            each_xlsx_addr = os.path.realpath(unique_filename)
            current_app.logger.debug(each_xlsx_addr)
            #记录每个文件的路径
            all_xlsx_addrs.append(each_xlsx_addr)
            current_app.logger.debug("该文件已保存: " + unique_filename)

        current_app.logger.debug("完成保存源文件.")
        #上传成功后的返回结果
        result = {
            "flag" : 1,
            "all_xlsx_addrs" : all_xlsx_addrs,
            "reason":'成功上传所有文件.'
        }
        current_app.logger.debug("上传成功: " + str(result))
        #返回上传成功后的结果
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;
    else:
        current_app.logger.debug("非POST请求,接口拒绝.")
        result = {
            "flag" :0,
            "reason":'请使用POST请求.'
        }
        current_app.logger.debug("上传失败: " + str(result))
        #返回失败后的结果
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;


"""
    编码接口
"""
@bp.route('/do424Code',methods=['GET'])
def do_424_Code():
    #获取用到的文件夹路径
    source_path = os.path.abspath('source_xlsx')
    result_path = os.path.abspath('result_xlsx')
    backend_path = os.path.abspath('backend_xlsx')
    #如果source_path下没有上传文件或者模版文件都为空时
    if len (os.listdir (source_path)) == 0 or len(os.listdir(backend_path)) == 0:
        current_app.logger.debug("没有上传文件或者模版为空，无法编码")
        result = {
            "flag" : 0,
            "reason" : "没有上传文件或者模版为空，无法编码.",
        }
        current_app.logger.debug("编码失败: " + str(result))
        #返回错误结果
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;

    #对每个文件进行编码
    current_app.logger.debug("Start coding saved sources.")
    unique_name_files = os.listdir(source_path)
    #用于存储每个文件的debug信息
    result_all_xlsxs = []
    #再次确认如果没有上传源文件，则不进行编码
    if(len(unique_name_files) == 0):
        current_app.logger.debug("source_xlsx文件夹没有上传待编码文件.")
        result = {
            "flag" :0,
            "reason":'source_xlsx文件夹没有上传待编码文件.'
        }
        current_app.logger.debug("编码失败: " + str(result))
        #返回错误结果
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;
    try:
        #编码过程
        for file in unique_name_files:
            current_app.logger.debug(f'正在编码： {file} now.')
            path = source_path+"/"+file;
            template_path = backend_path+'/424template.xlsx'
            Enroute_path = backend_path+'/1_ZBZL航路部分.xlsx'
            #利用顺序字典维护debug信息添加顺序
            result_each_xlsx = collections.OrderedDict()
            result_each_xlsx.setdefault("file", file)
            result_each_xlsx.setdefault(str(datetime.datetime.now()), "开始处理"+file)
            read_excel(path, 'Sheet1', template_path, 'PD',result_each_xlsx)
            read_enroute_excel(Enroute_path)


            TEMPLATE_DATA.iloc[1:, [4, 39, 7, 8, 9, 23, 12, 14, 17, 18, 19, 20, 30, 21, 35, 37, 38, 40, 22, 25, 32,
                               28, 29, 31, 27, 42, 45, 53, 54, 55]] = DATA.iloc[1:, :]
            padding_TEMPLATE_DATA_idnex = [0, 1, 2, 3, 4, 6, 7, 8, 9, 10, 16, 34, 41, 43, 47, 49, 50, 51, 52, 53, 54, 55, 56]
            for j in range(0, len(padding_TEMPLATE_DATA_idnex)):
                padding_sameValue(TEMPLATE_DATA, padding_TEMPLATE_DATA_idnex[j],result_each_xlsx)

            PD_P36_ATC_Indicator(TEMPLATE_DATA, 35, 36, 'A', 'A',result_each_xlsx)
            PD_P35_Altitude_Description(TEMPLATE_DATA, 35, 'A', ' ',result_each_xlsx)
            PD_P48_Speed_Limit_Description(TEMPLATE_DATA, 40, 48, '-',result_each_xlsx)
            PD_P5_ICAO_Code(result_each_xlsx)    
            Theta_Rho(TEMPLATE_DATA, 12, 14, 25, 32, 28, 29, result_each_xlsx)
            waypoint_code_1(TEMPLATE_DATA, 14, 17,result_each_xlsx)
            PD_P13_Ref_list = ['C', 'G', 'N', 'I', 'M']
            PD_P13_ICAO(TEMPLATE_DATA, PD_P13_Ref_list, 14, 13, 5, result_each_xlsx)

            PD_P26_Ref_list = ['N', 'I', 'M']
            PD_P26_ICAO(TEMPLATE_DATA, PD_P26_Ref_list, 32, 26, 5, result_each_xlsx)

            Sequence_Number(TEMPLATE_DATA,result_each_xlsx)
            # print(TEMPLATE_DATA)
            PD_P39_Transition_Altitude(TEMPLATE_DATA, 11, 39, '010', result_each_xlsx)
            section_sub_code(TEMPLATE_DATA, 14, 14, 15,result_each_xlsx)
            # print('wancheng')
            section_sub_code(TEMPLATE_DATA, 45, 45, 46,result_each_xlsx)
            # print(TEMPLATE_DATA)
            

            waypoint_code_4(TEMPLATE_DATA, 20, 20, result_each_xlsx)
           
            waypoint_code_2(TEMPLATE_DATA, 18, 11, 18, result_each_xlsx)
            Turn_Direction_Valid(TEMPLATE_DATA, 23, 21, 24, result_each_xlsx)
        
            PD_char_count = [1, 3, 1, 1, 4, 2, 1, 6, 1, 5, 1, 3, 5, 2, 1, 1, 1, 1, 1, 1, 1, 1, 3, 2, 1, 4, 2, 6, 4, 4, 4, 4, 1,
                             1, 2, 1, 1, 5, 5, 5, 3, 4, 5, 1, 2, 1, 1, 1, 1, 1, 1, 3, 5, 4, 10, 10, 10]
            SupplementarySpace(TEMPLATE_DATA, PD_char_count,result_each_xlsx)
            result_each_xlsx.setdefault(str(datetime.datetime.now()), "结束处理"+file)
          
            #将编码结果输出为excel文件
            wb = Workbook();
            wb.save(filename=result_path + "/output-" + file);
            writer = pd.ExcelWriter(path=result_path + "/output-" + file, engine='openpyxl')
            TEMPLATE_DATA.to_excel(result_path + "/output-" + file, index=False)
            result_all_xlsxs.append(result_each_xlsx)
            #清空这次debug信息为下次存储做准备
            result_each_xlsx = collections.OrderedDict()


        current_app.logger.debug("成功结束所有文件编码。")
        result = {
            "flag" : 1,
            "result_all_xlsxs": result_all_xlsxs
        }
        current_app.logger.debug("编码成功: " + str(result))
        #返回错误结果
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;

    except Exception as e:
        current_app.logger.debug("出现错误，终止编码.")
        #记录错误信息返回给前端debug信息展示
        result_each_xlsx.setdefault(str(datetime.datetime.now()), "出现错误：" + str(e.args))
        result_all_xlsxs.append(result_each_xlsx)
        result = {
            "flag" : 2,
            "result_all_xlsxs": result_all_xlsxs
        }
        current_app.logger.debug("result in failed: " + str(result))
        #返回错误结果
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;



"""
    download file by name
"""
@bp.route('/do424Download',methods=['GET'])
def do_424_Code_Download():
    if (request.args.get("filename") == None):
        current_app.logger.debug("Invalid operation! No filename provided.")
        result = {
            "flag" : 0,
            "reason" : "Invalid operation! No filename provided."
        }
        current_app.logger.debug("result in failed: " + str(result))
        #return wrong status
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;
    filename = request.args.get("filename");
    result_path = os.path.abspath('result_xlsx')
    if not os.path.exists(result_path+"/"+filename):
        current_app.logger.debug("Invalid operation! No this file in server to download.")
        result = {
            "flag" : 0,
            "reason" : "Invalid operation! No this file in server to download."
        }
        current_app.logger.debug("result in failed: " + str(result))
        #return wrong status
        resp = json.dumps(result);
        res = make_response(resp);
        res.headers["Access-Control-Allow-Origin"] = "*";
        return res;
    return send_from_directory(
        directory=result_path, path=filename, as_attachment=True
    )
"""
    delete files in folder
"""
def remove_file_or_dir(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            current_app.logger.error('Failed to delete %s. Reason: %s' % (file_path, e))

"""
    algo part for the 424 coding

"""
"""
    read excel
    argument: source path, SheetName, tempalte_path, tempalte_SheetName
    Return: processed dataFrame...
"""
def read_excel(path, SheetName, tempalte_data, tempalte_SheetName, result_each_xlsx):
    print("开始处理%s" % SheetName)
    result_each_xlsx.setdefault(str(datetime.datetime.now()), "开始处理%s" % SheetName)
    global DATA, TEMPLATE_DATA, max_row
    pd.options.mode.use_inf_as_na = True
    DATA = pd.read_excel(path, sheet_name=SheetName, index_col=None, header=None)
    TEMPLATE_DATA = pd.read_excel(tempalte_data, sheet_name=tempalte_SheetName, index_col=None, header=None)

    max_row = DATA.shape[0]
    print('人工编码模板最大行数为：%s' % max_row)
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '人工编码模板最大行数为：%s' % max_row)
    #在TEMPLATE_DATA DataFrame的基础上增加了填充了NaN值的行，总行数达到max_row。列数保持不变。
    TEMPLATE_DATA = pd.concat([TEMPLATE_DATA, pd.DataFrame([[pd.NA]*TEMPLATE_DATA.shape[1]]*(max_row-2))], ignore_index=True)

    return DATA, TEMPLATE_DATA, max_row

"""
    read_enroute_excel
    argument: enroute_path
    Return: processed dataFrame...
"""
def read_enroute_excel(path):
    global D__data, DB_data, PA_data, EA_data, PN_data, PI_data, PM_data, PG_data, PC_data
    D__data = pd.read_excel(path, sheet_name='D_', index_col=None, header=None)
    DB_data = pd.read_excel(path, sheet_name='DB', index_col=None, header=None)
    PA_data = pd.read_excel(path, sheet_name='PA', index_col=None, header=None)
    EA_data = pd.read_excel(path, sheet_name='EA', index_col=None, header=None)
    PN_data = pd.read_excel(path, sheet_name='PN', index_col=None, header=None)
    PI_data = pd.read_excel(path, sheet_name='PI', index_col=None, header=None)
    PM_data = pd.read_excel(path, sheet_name='PM', index_col=None, header=None)
    PG_data = pd.read_excel(path, sheet_name='PG', index_col=None, header=None)
    PC_data = pd.read_excel(path, sheet_name='PC', index_col=None, header=None)
    # return D__data, DB_data, PA_data, EA_data

"""
    按已填好的数据自动填充空白单元格
    argument: processed dataFrame, column
    Return: processed dataFrame...
"""
def padding_sameValue(dataframe, clo, result_each_xlsx):
    print('*****************人工编码模板转换424模板并填充__%s__*****************' % dataframe[clo][0])
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************人工编码模板转换424模板并填充__%s__*****************' % dataframe[clo][0])
    clo_data = dataframe[0:max_row][clo]
    # print(clo_data)
    # 获取非空的索引值，以列表输出。
    notnull_index = []
    for i in range(0,int(max_row)):
        # print(clo_data[i],i)
        if pd.isna(clo_data[i]) is False:
            notnull_index.append(i)
    notnull_index.append(max_row)
    # print(notnull_index)
    #计算出非空值间的间隔值，以列表输出
    notnull_index_interval =[]
    for j in range(0,len(notnull_index)-1):
        notnull_index_interval.append(notnull_index[j+1]-notnull_index[j])
    # print(notnull_index_interval)
    for k in range(0,len(notnull_index_interval)):
        #筛选出无间隔的数据
        if notnull_index_interval[k] == 1:
            continue
        else:
            #将间隔中的数据补充
            for l in range(1, notnull_index_interval[k]):
                # print(l)
                clo_data.loc[notnull_index[k]+l] = clo_data.loc[notnull_index[k]]
    # print(clo_data)
    # print(dataframe[:max_row][clo])
    return dataframe

"""
此方法需要在PD_P35_Altitude_Description方法前执行，P35的数据为A时，P36填写A，否则为空
"""
def PD_P36_ATC_Indicator(dataframe, da_ref_clo, da_target_clo, ref_data, new_data, result_each_xlsx):
    print('*****************完成转换__%s__*****************' % dataframe[da_target_clo][0])
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************完成转换__%s__*****************' % dataframe[da_target_clo][0])
    P35_A = [i for i in range(max_row) if dataframe[da_ref_clo][i] == ref_data]
    for j in range(len(P35_A)):
        dataframe.loc[P35_A[j], da_target_clo] = new_data
"""
将人工模板填充为“A”的单元格转换为空
"""
def PD_P35_Altitude_Description(dataframe, da_target_clo, old_data, new_data,result_each_xlsx):
    print('*****************完成转换__%s__*****************' % dataframe[da_target_clo][0])
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************完成转换__%s__*****************' % dataframe[da_target_clo][0])

    dataframe[da_target_clo] = dataframe[da_target_clo].replace(old_data, new_data)

"""
#P40有值，填“-”
"""
def PD_P48_Speed_Limit_Description(dataframe, da_ref_clo, da_target_clo,  new_data,result_each_xlsx):
    print('*****************完成转换__%s__*****************' % dataframe[da_target_clo][0])
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************完成转换__%s__*****************' % dataframe[da_target_clo][0])

    P40_notna = [i for i in range(1, max_row) if pd.isna(dataframe[da_ref_clo][i]) is False]
    for j in range(len(P40_notna)):
        dataframe.loc[P40_notna[j], da_target_clo] = new_data
class Vincenty:
    """处理经度度坐标点正算与反算问题
    工具类，不可创建对象"""

    @staticmethod
    def inverse(lat1, lon1, lat2, lon2):
        """给定两个经纬度点求距离与角度
        根据Vincenty算法编写
        Doon 2021/06/09"""
        lon1 = Vincenty.deg2rad(lon1)
        lat1 = Vincenty.deg2rad(lat1)
        lon2 = Vincenty.deg2rad(lon2)
        lat2 = Vincenty.deg2rad(lat2)
        # WGS84
        a = 6378137.0  # meters
        b = 6356752.314245  # meters
        f = 1 / 298.257223563

        L = lon2 - lon1
        tanU1 = (1 - f) * math.tan(lat1)
        cosU1 = 1 / math.sqrt(1 + tanU1 * tanU1)
        sinU1 = tanU1 * cosU1

        tanU2 = (1 - f) * math.tan(lat2)
        cosU2 = 1 / math.sqrt(1 + tanU2 * tanU2)
        sinU2 = tanU2 * cosU2

        lambda_ = L
        cos2Alpha, cos2DeltaM, sinLambda, cosLambda, delta, sinDelta, cosDelta, lambda0 = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
        while True:
            sinLambda = math.sin(lambda_)
            cosLambda = math.cos(lambda_)
            sinSqDelta = (cosU2 * sinLambda) * (cosU2 * sinLambda) + \
                         (cosU1 * sinU2 - sinU1 * cosU2 * cosLambda) * (cosU1 * sinU2 - sinU1 * cosU2 * cosLambda)
            sinDelta = math.sqrt(sinSqDelta)

            cosDelta = sinU1 * sinU2 + cosU1 * cosU2 * cosLambda
            delta = math.atan2(sinDelta, cosDelta)

            sinAlpha = cosU1 * cosU2 * sinLambda / sinDelta

            cos2Alpha = 1 - sinAlpha * sinAlpha

            cos2DeltaM = cosDelta - 2 * sinU1 * sinU2 / cos2Alpha
            c = f / 16 * cos2Alpha * (4 + f * (4 - 3 * cos2Alpha))
            lambda0 = lambda_
            lambda_ = L + (1 - c) * f * sinAlpha * (delta + c * sinDelta * (cos2DeltaM + c * cosDelta * (-1 + 2 * cos2DeltaM * cos2DeltaM)))

            if abs(lambda_ - lambda0) <= 1e-12:
                break

        uSq = cos2Alpha * (a * a - b * b) / (b * b)
        A = 1 + uSq / 16384 * (4096 + uSq * (-768 + uSq * (320 - 175 * uSq)))
        B = uSq / 1024 * (256 + uSq * (-128 + uSq * (74 - 47 * uSq)))
        deltaDelta = B * sinDelta * (cos2DeltaM + B / 4 * (cosDelta * (-1 + 2 * cos2DeltaM * cos2DeltaM)) -
                                     B / 6 * cos2DeltaM * (-3 + 4 * sinDelta * sinDelta) * (-3 + 4 * cos2DeltaM * cos2DeltaM))
        s = b * A * (delta - deltaDelta)
        bearing0 = math.atan2(cosU2 * sinLambda, cosU1 * sinU2 - sinU1 * cosU2 * cosLambda)
        bearing1 = math.atan2(cosU1 * sinLambda, -sinU1 * cosU2 + cosU1 * sinU2 * cosLambda)
        '''
        bearing0: 点1到点2
        bearing1：点2到点1
        '''
        return s, ((bearing0 / math.pi * 180 + 360) % 360), ((bearing1 / math.pi * 180 + 180) % 360)

    @staticmethod
    def forward(lat, lon, bearing, distance, var=0):
        """已知一点经纬度坐标，方位角和距离（m），求另一点经纬度
        （磁方位+磁差）
        根据Vincenty算法编写
        Doon 2021/06/09"""
        lat = Vincenty.deg2rad(lat)
        lon = Vincenty.deg2rad(lon)
        # WGS84
        a = 6378137.0  # meters
        b = 6356752.314245  # meters
        f = 1 / 298.257223563

        alpha1 = Vincenty.deg2rad(bearing + var)

        sinAlpha1 = math.sin(alpha1)
        cosAlpha1 = math.cos(alpha1)
        tanU1 = (1 - f) * math.tan(lat)
        cosU1 = 1 / math.sqrt(1 + tanU1 * tanU1)
        sinU1 = tanU1 * cosU1

        delta1 = math.atan2(tanU1, cosAlpha1)

        sinAlpha = cosU1 * sinAlpha1
        cos2Alpha = 1 - sinAlpha * sinAlpha
        uSq = cos2Alpha * (a * a - b * b) / b / b
        A = 1 + uSq * uSq / 16384 * (4096 + uSq * uSq * (-768 + uSq * (320 - 175 * uSq)))
        B = uSq / 1024 * (256 + uSq * (-128 + uSq * (74 - 47 * uSq)))

        delta = distance / b / A
        delta0, sinDelta, cosDelta, cos2DeltaM = 0.0, 0.0, 0.0, 0.0
        while True:
            cos2DeltaM = math.cos(2 * delta1 + delta)
            sinDelta = math.sin(delta)
            cosDelta = math.cos(delta)
            deltaDelta = B * sinDelta * (cos2DeltaM + B / 4 * (cosDelta * (-1 + 2 * cos2DeltaM * cos2DeltaM)) -
                                         B / 6 * cos2DeltaM * (-3 + 4 * sinDelta * sinDelta) * (-3 + 4 * cos2DeltaM * cos2DeltaM))
            delta0 = delta
            delta = distance / b / A + deltaDelta

            if abs(delta - delta0) <= 1e-12:
                break
        tempx = sinU1 * sinDelta - cosU1 * cosDelta * cosAlpha1
        lat2 = math.atan2(sinU1 * cosDelta + cosU1 * sinDelta * cosAlpha1, (1 - f) * math.sqrt(sinAlpha * sinAlpha + tempx * tempx))
        lambada = math.atan2(sinDelta * sinAlpha1, cosU1 * cosDelta - sinU1 * sinDelta * cosAlpha1)
        C = f / 16 * cos2Alpha * (4 + f * (4 - 3 * cos2Alpha))
        L = lambada - (1 - C) * f * sinAlpha * (delta + C * sinDelta * (cos2DeltaM + C * cosDelta * (-1 + 2 * cos2DeltaM * cos2DeltaM)))
        lambada2 = lon + L
        alpha2 = math.atan2(sinAlpha, -tempx)

        return lat2 / math.pi * 180, lambada2 / math.pi * 180, (alpha2 / math.pi * 180 + 180) % 360

    @staticmethod
    def forwardXY(lat, lon, x, y, var=0):
        """已知一点经纬度坐标和以该点为坐标原点建立的坐标系中的一点坐标（x,y）
        该坐标系以磁北为Y轴，当磁差不为0时，可自动进行坐标轴旋转坐标转换
        （Forward算法计算过程中方位为真方位）
        求该点的经纬度"""
        bearing, distance = Vincenty.cal_azimuth(x, y, var)
        return Vincenty.forward(lat, lon, bearing, distance, var)

    @staticmethod
    def deg2rad(degree):
        """度转弧度"""
        return degree * math.pi / 180

    @staticmethod
    def rotate_axis(x, y, var):
        """坐标轴旋转方位变化"""
        rad = Vincenty.deg2rad(var)
        xx = x * math.cos(rad) + y * math.sin(rad)
        yy = y * math.cos(rad) - x * math.sin(rad)
        return xx, yy

    @staticmethod
    def rad2deg(rad):
        return rad * 180 / math.pi

    @staticmethod
    def cal_azimuth(x, y, var=0):
        """可将X,Y所在坐标系旋转一定角度后，计算X,Y相对于新坐标系原点的方位与距离"""
        xx, yy = Vincenty.rotate_axis(x, y, var)
        bearing = 0.0
        distance = math.sqrt(xx * xx + yy * yy)
        if xx >= 0:
            bearing = Vincenty.rad2deg(math.atan2(xx, yy))
        else:
            bearing = 360 + Vincenty.rad2deg(math.atan2(xx, yy))
        return bearing, distance



"""
424格式经纬度转换为度
"""
def transfrom_LaLo(data,result_each_xlsx):

    if data[0] == 'N' or data[0] == 'S':
        # print('N\S%s' %data[0] )
        # print(data[1:3],data[3:5],data[5:9])
        print('转换纬度格式：%s' % data)
        result_each_xlsx.setdefault(str(datetime.datetime.now()), '转换纬度格式：%s' % data)

        transform_data = int(data[1:3]) + int(data[3:5])/60 + int(data[5:9])/360000
        # print(transform_data)
    elif data[0] == 'E' or data[0] == 'W':
        # print('E\W%s' %data[0] )
        print('转换经度格式：%s' % data)
        result_each_xlsx.setdefault(str(datetime.datetime.now()),'转换经度格式：%s' % data)

        transform_data = int(data[1:4]) + int(data[4:6])/60 + int(data[6:10])/360000
        # print(transform_data)
    else:
        result_each_xlsx.setdefault(str(datetime.datetime.now()),"经纬度人工编码数据错误")
        print("经纬度人工编码数据错误")
    return transform_data

'''

#根据人工编码的记录标号匹配实际记录所读取后的dataframe，并返回对应的经纬度与标识符的索引号

match_data：匹配字典后的dataframe
Lo_index： 该dataframe下需要索引的纬度索引列
La_index： 该dataframe下需要索引的经度索引列
ID_index：  用于匹配相同标识符的索引列
'''
def record_data(dataframe, i ,Ref):
    Dic_record = {'D': [D__data, 13, 14, 7], 'C': [PC_data, 15, 16, 7], 'G': [PG_data, 13, 14, 7], 'N': [PN_data, 13, 14, 7],
                  'B': [DB_data, 13, 14, 7],'I': [PI_data, 13, 14, 7], 'A': [PA_data, 15, 16, 4],'E': [EA_data, 15, 16, 7], 'M': [PM_data, 13, 14, 7]}
    match_data, Lo_index, La_index, ID_index = [value1 for (key1, value1) in Dic_record.items() if key1 == dataframe[Ref][i]][0]

    return match_data, Lo_index, La_index, ID_index

"""

#查找定位点在此机场下的磁差值,并基于真方位返回磁方位

"""
def find_PA_Magnetic_Variation(dataframe, ICAO_index, bearing):

    for i in range(1, max_row):
        for j in range(PA_data.shape[0]):
            if dataframe[ICAO_index][i] == PA_data[4][j]:
                # print(dataframe[ICAO_index][i])
                # print(PA_data[17][j])
                if PA_data[17][j][0] == 'W':
                    magnetic_course = bearing + int(PA_data[17][j][1:5]) / 10
                elif PA_data[17][j][0] == 'E':
                    magnetic_course = bearing - int(PA_data[17][j][1:5]) / 10
                elif PA_data[17][j][0] == 'T':
                    magnetic_course = bearing
    return magnetic_course




'''
# 计算推荐导航台的Theta、Rho值

Fix_Id_index: 定位点标识符列索引号
Fix_Ref_Code: 定位点人工编码的记录标号列索引号
Rec_Nav_index: 推荐导航台列索引号
Rec_Nav_Ref_Code: 推荐导航台人工编码的记录标号列索引号
Theta_index: 方位角列索引号
Rho_index:  测距弧列索引号
'''
def Theta_Rho(dataframe, Fix_Id_index, Fix_Ref_Code, Rec_Nav_index, Rec_Nav_Ref_Code, Theta_index, Rho_index, result_each_xlsx):
    print('*****************推荐导航台Theta、Rho值计算*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()),"*****************推荐导航台Theta、Rho值计算*****************")
   
    for i in range(1,max_row):
        if pd.isna(dataframe[Rec_Nav_index][i]) or pd.isna(dataframe[Fix_Id_index][i]) is True:
            continue
        else:
            print("处理定位点%s" % dataframe[Fix_Id_index][i])
            result_each_xlsx.setdefault(str(datetime.datetime.now()),"处理定位点%s" % dataframe[Fix_Id_index][i])

            Fix_dataframe, Fix_La_index, Fix_Lo_index, Fix_dataframe_RefIndex = record_data(dataframe, i, Fix_Ref_Code)
            # print(Fix_dataframe, Fix_La_index, Fix_Lo_index, Fix_dataframe_RefIndex )
            Rec_Nav_datafarme, Rec_Nav_La_index, Rec_Nav_Lo_index, Rec_Nav_dataframe_RefIndex = record_data(dataframe, i, Rec_Nav_Ref_Code)
            # print(Rec_Nav_datafarme, Rec_Nav_La_index, Rec_Nav_Lo_index, Rec_Nav_dataframe_RefIndex)
            for k in range(Fix_dataframe.shape[0]):
                if dataframe[Fix_Id_index][i] == Fix_dataframe[Fix_dataframe_RefIndex][k]:
                    Fix_La = transfrom_LaLo(Fix_dataframe[Fix_La_index][k],result_each_xlsx)
                    Fix_Lo = transfrom_LaLo(Fix_dataframe[Fix_Lo_index][k],result_each_xlsx)

                    print('定位点标识符经纬度为：%s,%s'%(Fix_La ,Fix_Lo))
                    result_each_xlsx.setdefault(str(datetime.datetime.now()),'定位点标识符经纬度为：%s,%s'%(Fix_La ,Fix_Lo))

                    find_Fix_dataframe = True
                    break
                else:
                    # print("df2未找到df3匹配项")
                    find_Fix_dataframe = False
            print("处理推荐导航台%s" % dataframe[Rec_Nav_index][i])
            result_each_xlsx.setdefault(str(datetime.datetime.now()),"处理推荐导航台%s" % dataframe[Rec_Nav_index][i])

            for l in range(Rec_Nav_datafarme.shape[0]):
                if dataframe[Rec_Nav_index][i] == Rec_Nav_datafarme[Rec_Nav_dataframe_RefIndex][l]:
                    Rec_Nav_La = transfrom_LaLo(Rec_Nav_datafarme[Rec_Nav_La_index][l], result_each_xlsx)
                    Rec_Nav_Lo = transfrom_LaLo(Rec_Nav_datafarme[Rec_Nav_Lo_index][l], result_each_xlsx)
                    print('推荐导航台经纬度为：%s,%s'%(Rec_Nav_La ,Rec_Nav_Lo))
                    result_each_xlsx.setdefault(str(datetime.datetime.now()), '推荐导航台经纬度为：%s,%s'%(Rec_Nav_La ,Rec_Nav_Lo))

                    find_Rec_Nav_datafarme = True
                    break
                else:
                    # print("df2未找到df匹配项")
                    find_Rec_Nav_datafarme = False
            if find_Fix_dataframe is False:
                print("%s未找到匹配的数据" % dataframe[Fix_Id_index][i])
                result_each_xlsx.setdefault(str(datetime.datetime.now()), "%s未找到匹配的数据" % dataframe[Fix_Id_index][i])


            if find_Rec_Nav_datafarme is False:
                print("%s未找到匹配的数据" % dataframe[Rec_Nav_index][i])
                result_each_xlsx.setdefault(str(datetime.datetime.now()), "%s未找到匹配的数据" % dataframe[Rec_Nav_index][i])

            if find_Fix_dataframe & find_Rec_Nav_datafarme is True:
                distance, bearing_a, bearing_b = Vincenty.inverse(Fix_La, Fix_Lo, Rec_Nav_La, Rec_Nav_Lo)

                bearing_b = find_PA_Magnetic_Variation(dataframe, 4, bearing_b)
                formatted_424bearing = "{:.1f}".format(bearing_b).replace(".", "").zfill(4)
                formatted_424distance = "{:.1f}".format(distance/1852).replace(".", "").zfill(4)
                dataframe.loc[i, Rho_index], dataframe.loc[i, Theta_index] = formatted_424distance, formatted_424bearing

                print('%s相对于推荐导航台%s的方位角和距离为:%s,%s' %(dataframe[Fix_Id_index][i], dataframe[Rec_Nav_index][i],
                                                     dataframe.loc[i,Theta_index],dataframe.loc[i, Rho_index]))
                result_each_xlsx.setdefault(str(datetime.datetime.now()), '%s相对于推荐导航台%s的方位角和距离为：%s，%s' %(dataframe[Fix_Id_index][i], dataframe[Rec_Nav_index][i],
                                                     dataframe.loc[i,Theta_index],dataframe.loc[i, Rho_index]))

   


'''
PD记录P13的ICAO代码判断

Ref_list: 此列表中的记录号不需要查找其他表格，索引自身数据即可
da_ref_clo: 主数据的参考列的索引
da_target_clo:主数据的目标更改列索引
r_da_target_clo:主数据的目标参考列索引
'''
def PD_P13_ICAO(dataframe, Ref_list, da_ref_clo, da_target_clo, r_da_target_clo,result_each_xlsx):
    print('*****************PD记录P13的ICAO代码判断r*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************PD记录P13的ICAO代码判断r*****************')
   
    for i in range(1, max_row):
        if pd.isna(dataframe[12][i]) is False:
            if dataframe[da_ref_clo][i] in Ref_list:
                dataframe.loc[i, da_target_clo] = dataframe.loc[i, r_da_target_clo]
            elif dataframe[da_ref_clo][i] == 'D':
                get_same_ICAO(dataframe, D__data, 12, 7, 13, 9, i,result_each_xlsx)
            elif dataframe[da_ref_clo][i] == 'B':
                get_same_ICAO(dataframe, DB_data, 12, 7, 13, 9, i, result_each_xlsx)
            elif dataframe[da_ref_clo][i] == 'E':
                get_same_ICAO(dataframe, EA_data, 12, 7, 13, 9, i, result_each_xlsx)
            elif dataframe[da_ref_clo][i] == 'A':
                get_same_ICAO(dataframe, PA_data, 12, 4, 13, 5, i, result_each_xlsx)
            else:
                print('第%s行的%s人工编码数据有误' % (i, dataframe[da_ref_clo][i]))
                result_each_xlsx.setdefault(str(datetime.datetime.now()), '第%s行的%s人工编码数据有误' % (i, dataframe[da_ref_clo][i]))

                dataframe.loc[i, da_ref_clo] = '人工编码数据有误'
        else:
            continue

"""
    PD记录P26的ICAO代码判断
"""
def PD_P26_ICAO(dataframe, Ref_list, da_ref_clo, da_target_clo, r_da_target_clo,result_each_xlsx ):
    print('*****************PD记录P26的ICAO代码判断r*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************PD记录P26的ICAO代码判断r*****************')

    for i in range(1, max_row):
        if pd.isna(dataframe[25][i]) is False:
            if dataframe[da_ref_clo][i] in Ref_list:
                dataframe.loc[i, da_target_clo] = dataframe.loc[i, r_da_target_clo]
            elif dataframe[da_ref_clo][i] == 'D':
                get_same_ICAO(dataframe, D__data, 25, 7, 26, 9, i, result_each_xlsx)
            elif dataframe[da_ref_clo][i] == 'B':
                get_same_ICAO(dataframe, DB_data, 25, 7, 26, 9, i, result_each_xlsx)
            else:
                print('第%s行的%s人工编码数据有误' % (i, dataframe[da_ref_clo][i]))
                result_each_xlsx.setdefault(str(datetime.datetime.now()), '第%s行的%s人工编码数据有误' % (i, dataframe[da_ref_clo][i]))

                dataframe.loc[i, da_ref_clo] = '人工编码数据有误'
        else:
            continue

def PD_P5_ICAO_Code(result_each_xlsx):
    print('*****************PD记录P5的ICAO代码判断r*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************PD记录P5的ICAO代码判断r*****************')

    for i in range(1, max_row):
        get_same_ICAO(TEMPLATE_DATA, PA_data, 4, 4, 5, 5, i, result_each_xlsx)


'''
ref_dataframe: 主数据所需关联的参考数据
da_match_index: 主数据中用于匹配的列索引
r_da_match_index:参考数据中用于匹配的列索引
da_target_index:主数据的目标更改列索引
r_da_target_index：将值传递给主数据的目标更改列的参考列索引
'''
def get_same_ICAO(dataframe, ref_dataframe, da_match_index, r_da_match_index, da_target_index, r_da_target_index, i,result_each_xlsx):
    indices = [j for j, x in enumerate(ref_dataframe[r_da_match_index][:]) if ref_dataframe[r_da_match_index][j] == dataframe[da_match_index][i]]

    if indices:
        dataframe.loc[i, da_target_index] = ref_dataframe.loc[indices[0],r_da_target_index]
    else:
        print('%s未匹配到数据'%[da_target_index][i])
        result_each_xlsx.setdefault(str(datetime.datetime.now()), '%s未匹配到数据'%[da_target_index][i])

        dataframe.loc[i, da_target_index] = '未匹配到数据'


"""
    自动填充Sequence_Number
    argument: processed dataFrame...
    Return: processed dataFrame...
"""
def Sequence_Number(dataframe,result_each_xlsx):
    print('*****************开始填充Sequence_Number*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************开始填充Sequence_Number*****************')

    #机场四码、程序名、过渡名相同时判断重复行的索引值
    df = dataframe[dataframe.duplicated(keep=False,subset=dataframe.columns[[4,7,9]])]

    # 对获取出来相同数值的行索引，用元组组合，并使用列表将元组
    df = df.groupby(dataframe.columns[[4,7,9]].tolist()).apply(lambda x: tuple(x.index)).tolist()
    # print(df)
    #生成序列号，并替换原数据
    for i in range(len(df)):
        for j in range(len(df[i])):
            Sequence_Number=str((df[i][j]-df[i][0]+1) * 10).zfill(3)
            # print(Sequence_Number)
            dataframe.loc[df[i][j], 11]=Sequence_Number
    # print(dataframe)
    return dataframe

"""
此方法需要在Sequence_Number后执行，P39第一条记录的内容填充至所有P11为“010”的记录
"""

def PD_P39_Transition_Altitude(dataframe, da_ref_clo, da_target_clo, ref_data,result_each_xlsx):
    print('*****************完成填充__%s__*****************' % dataframe[da_target_clo][0])
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************完成填充__%s__*****************' % dataframe[da_target_clo][0])

    P11_010 = [i for i in range(max_row) if dataframe[da_ref_clo][i] == ref_data]
    for j in range(len(P11_010)):
        dataframe.loc[P11_010[j], da_target_clo] =  dataframe[da_target_clo][1]

"""
根据输入的一个字符匹配完整的记录名
"""
def section_sub_code(dataframe, Ref, Sec, Subsec, result_each_xlsx):
    print('*****************人工编码记录标号转换*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************人工编码记录标号转换*****************')

    Dic_Sec_Sub = {'D': ['D',' '], 'C': ['P', 'C'], 'G': ['P', 'G'], 'N': ['P', 'N'], 'B': ['D', 'B'], 'I': ['P', 'I'], 'A': ['P', 'A'],
             'E': ['E', 'A'], 'M': ['P', 'M']}
    for i in range(1, max_row):
        # print(i)
        # 使用字典键值对匹配
        match_data = [value1 for (key1, value1) in Dic_Sec_Sub.items() if key1 == dataframe[Ref][i]]
        if len(match_data):
            # print(match_data)
            dataframe.loc[i, Sec], dataframe.loc[i, Subsec] = match_data[0][0], match_data[0][1]
        else:
            # print(match_data)
            dataframe.loc[i, Sec], dataframe.loc[i, Subsec] = ' ', ' '    
"""
    航路点类型第一位字符判断
    argument: processed dataFrame...
    Return: processed dataFrame...
"""
def waypoint_code_1(dataframe, Ref, code, result_each_xlsx):
    print('*****************航路点类型第一位字符判断*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************航路点类型第一位字符判断*****************')

    Dic_way_code1 = {'D': 'V', 'C': 'E', 'G': 'G', 'N': 'N', 'B': 'N', 'I': 'V', 'A': 'A', 'E': 'E'}
    for i in range(1, max_row):
        # print(i)
        match_data = [value1 for (key1, value1) in Dic_way_code1.items() if key1 == dataframe[Ref][i]]
        if len(match_data):
            # print(match_data)
            dataframe.loc[i, code] = match_data[0]
        else:
            # print(match_data)
            dataframe.loc[i, code] = ' '



"""
    航路点类型第二位字符判断
    argument: processed dataFrame...
    Return: processed dataFrame...
"""
def waypoint_code_2(dataframe, Ref1, Ref2, code, result_each_xlsx):
    print('*****************航路点类型第二位字符判断*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************航路点类型第二位字符判断*****************')

    for i in range(1, max_row-1):
        if pd.isna(dataframe[Ref1][i]) is True:
            if int(dataframe[Ref2][i]) < int(dataframe[Ref2][i+1]) :
                dataframe.loc[i, code]= ' '
            elif int(dataframe[Ref2][i]) > int(dataframe[Ref2][i+1]):
                dataframe.loc[i, code]= 'E'
        elif dataframe[Ref1][i] == 'Y':
            if int(dataframe[Ref2][i]) < int(dataframe[Ref2][i+1]) :
                dataframe.loc[i, code]= 'Y'
            elif int(dataframe[Ref2][i]) > int(dataframe[Ref2][i+1]):
                dataframe.loc[i, code]= 'B'
    #最大行的判断
    if pd.isna(dataframe[Ref1][max_row-1]) is True:
        dataframe.loc[max_row-1, code] = 'E'
    elif dataframe[Ref1][max_row-1] == 'Y':
        dataframe.loc[max_row-1, code] = 'B'
"""
    航路点类型第三位字符判断
    argument: processed dataFrame...
    Return: processed dataFrame...
"""
def waypoint_code_4(dataframe, Ref, code,result_each_xlsx):
    print('*****************航路点类型第三位字符判断*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************航路点类型第三位字符判断*****************')

    for i in range(1, max_row):
        if pd.isna(dataframe[Ref][i]) is True:
            dataframe.loc[i, code] = ' '
        elif dataframe[Ref][i] == 'Y':
            dataframe.loc[i, code]  = 'H'
        else:
            print('第%s行的%s 人工编码数据有误'%(i, dataframe[code][i]))
            result_each_xlsx.setdefault(str(datetime.datetime.now()), '第%s行的%s 人工编码数据有误'%(i, dataframe[code][i]))

            dataframe.loc[i, code] = '人工编码数据有误'

"""
    转弯方向编码校验
    argument: processed dataFrame...
    Return: processed dataFrame...
"""
def Turn_Direction_Valid(dataframe, Ref1, Ref2, target, result_each_xlsx):
    print('*****************转弯方向判断*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************转弯方向判断*****************')

    PT_coed = ['AF', 'DF', 'HA', 'HF', 'HM', 'IF', 'PI', 'RF']
    for i in range(1, max_row):
        if pd.isna(dataframe[Ref2][i]) is True:
            dataframe.loc[i, target] = ' '
        elif dataframe[Ref2][i] == 'L' or 'R':
            if dataframe[Ref1][i] in PT_coed:
                dataframe.loc[i, target]= ' '
            else:
                dataframe.loc[i, target]= 'Y'
        else:
            print('第%s行的%s 人工编码数据有误' % (i, dataframe[target][i]))
            result_each_xlsx.setdefault(str(datetime.datetime.now()), '第%s行的%s 人工编码数据有误' % (i, dataframe[target][i]))

            dataframe.loc[i, target]= '转弯方向编码数据有误'

"""
按424要求补充各个字段的空格
"""
def SupplementarySpace(dataframe, char_count,result_each_xlsx):
    print('*****************字段空格补充完成*****************')
    result_each_xlsx.setdefault(str(datetime.datetime.now()), '*****************字段空格补充完成*****************')

    dataframe.fillna(' ', inplace=True)
    for i in range(1, max_row):
        for j in range(dataframe.shape[1]):
            dataframe.loc[i,j] = str(dataframe.loc[i,j]).ljust(char_count[j])













