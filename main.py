# import xlwings
import pathlib
import random
import subprocess
import win32ui
import win32con
import pandas
import datetime
import tools
import os
from collections import defaultdict


def open_multiple_files_dialog():
    # 创建文件对话框对象时设置多选样式
    dlg = win32ui.CreateFileDialog(
        1,  # 1表示打开对话框，0表示保存对话框
        None,  # 默认文件扩展名
        None,  # 默认文件名
        win32con.OFN_ALLOWMULTISELECT | win32con.OFN_EXPLORER,  # 启用多选和Explorer样式
        "All Files (*.*)|*.*|"
        "Python Files (*.py)|*.py|"
        "Text Files (*.txt)|*.txt|"
        "Xls Files (*.xls)|*.xls|"
        "Xlsx Files (*.xlsx)|*.xlsx|"
    )
    
    # 显示对话框
    if dlg.DoModal() == win32con.IDOK:
        # 获取选择的文件路径
        paths = dlg.GetPathNames()
        return paths
    else:
        return []  # 用户取消了对话框


if __name__ == "__main__":
    title = '地市,区县,客户名称,设备品牌,设备条码,设备名称,维修日期,返回日期,设备类别,故障原因定位分析,分析人,审核人,报废图片'
    # title = '客户名称,设备品牌,设备条码,设备名称,维修日期,设备类别,故障原因定位分析,分析人,审核人,故障原因,报废图片'
    
    
    del_list=['有线平台送修单号','保修类型：翻新/保内送修','送修厂家','物料型号','物料编码',
              '物资属性','SN码','标记','翻新后物料编码','翻新后物资属性','语音机顶盒标记','蓝牙遥控器串码',
              '蓝牙遥控器类型','光功率值','软件版本','配置参数（是否恢复出厂设置）','老化时间','送修日期',
              '测试日期','翻新返回录入单号','地市确认是否属实','返回时长','地市核对日期','备注']
    feedback = pandas.DataFrame()
    scrap_SN = pandas.DataFrame()
    selected_files = open_multiple_files_dialog()
    
    current_path = pathlib.Path(__file__).parent
    
    img_path = current_path / 'temp/LIST.csv'
    pic = pandas.read_csv(img_path, encoding='utf-8',header=0)
    
    if selected_files:
        # print("选择的文件:")
        for file_path in selected_files:
           scrap_SN = pandas.concat([scrap_SN, pandas.read_excel(file_path, '报废SN')])
           feedback = pandas.concat([feedback, pandas.read_excel(file_path, '翻新设备详细处理流程反馈表')])           
    else:
        print("未选择任何文件")
        exit()
    area = feedback['地市'].unique()
    
    # 绍兴丽水和其他地区报废模板不一致需要分开处理
    if area.shape[0] > 1:
        print('选择的文件包含多个地市，请检查数据！')
        subprocess.run('pause', shell=True)
        exit()
    else:
        if area[0] not in ['丽水', '绍兴']:
            scrap_SN = scrap_SN.reset_index(drop=True)  # 重置索引
        else:
            scrap_SN = scrap_SN.drop(index=[0, 1],inplace=False)    # 删除前两行
            scrap_SN = scrap_SN.reset_index(drop=True)  # 重置索引
            
    #之前模板存在重复表头，需要更改
    if '物资属性.1' in scrap_SN.columns:
        scrap_SN.rename(columns={'物资属性.1': '翻新后物资属性'}, inplace=True)  # 重命名列
    if '不能翻新原因(下拉选项)' in feedback.columns:
            feedback.rename(columns={"不能翻新原因(下拉选项)": "不能翻新原因"}, inplace=True)
    scrap_SN = scrap_SN.drop(columns=del_list)  #删除不需要的列
    feedback = feedback.reset_index(drop=True)  # 重置索引
    
    #检测报废SN和反馈表是否一致，一致后
    if scrap_SN['翻新后SN码'].equals(feedback['SN']):
        result = pandas.concat([scrap_SN, feedback['不能翻新原因']], axis=1)
        result['分析人'] = '姚工森'
        result['审核人'] = '刘连英'
        # result['故障原因'] = result['不能翻新原因']
        result['报废图片'] = ''
        
        for i in range(result['报废图片'].shape[0]):
            result.loc[i, '报废图片'] = pic['picture'][random.randint(0, pic.shape[0]-1)]   
        result.columns = title.split(',')
        
        result.to_excel(current_path / 'list.xlsx', sheet_name='报废SN', startcol=0, startrow=0, index=False)
        tools.excel2pdf(result,current_path)
        district = []
        
        for i in range(0,result.shape[0]):
            district.append(str(result.loc[i,'地市']) + str(result.loc[i,'区县']) + str(result.loc[i,'返回日期'].strftime("%Y%m%d")))
        #print(str(districts.loc[1,'地市']) + str(districts.loc[1,'区县']) + str(districts.loc[1,'返回日期'].strftime("%Y%m%d")))
        districts = pandas.unique(pandas.Series(district))
    
        pdf_files = {dist: [] for dist in districts}
        # print(pdf_files)
        for i in range(0,result.shape[0]):
            temp = str(result.loc[i,'地市']) + str(result.loc[i,'区县']) + str(result.loc[i,'返回日期'].strftime("%Y%m%d"))
            if temp in districts:
                fullpath = str(current_path) + '/pdf/' + str(result.loc[i,'客户名称']) + str(result.loc[i,'设备条码'])+'.pdf'
                pdf_files[temp].append(fullpath)
        key_list = list(pdf_files.keys())
        for l in range(0,len(pdf_files)):
            tools.compress_pdfs(pdf_files[key_list[l]], f'{current_path}/pdf/{key_list[l]}报废鉴定报告.zip')
        tools.delete_pdf_files(current_path/'pdf')#删除pdf文件夹下的pdf文件
        subprocess.run('pause', shell=True)
    else:
        print('翻新后SN码与反馈表SN不匹配，请检查数据！')
        subprocess.run('pause', shell=True)