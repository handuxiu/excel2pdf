import datetime
import os
import zipfile
import jinja2
import pdfkit
import pandas

# district_lishui = ['缙云', '景宁', '莲都', '龙泉', '南城', '青田', '庆元', '松阳', '遂昌', '云和']
# district_shaoxing = ['上虞', '嵊州', '绍兴', '新昌', '诸暨','柯桥','越城']
# district_wenzhou = ['洞头', '乐清', '龙湾', '瑞安', '平阳', '苍南', '文成', '泰顺']
# district_taizhou = ['临海', '温岭', '玉环', '三门', '天台', '仙居', '椒江', '路桥']
# district_jinhua = ['东阳', '金华', '兰溪', '磐安', '浦江', '武义', '义乌', '永康','金东','婺城']
# district_quzhou = ['常山', '江山', '开化', '龙游', '衢州','柯城']
def data2str(data):
    return data.strftime('%Y%m%d')
def compress_pdfs(file_list, output_zip):
  
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in file_list:
            # 检查文件是否存在，避免报错
            if os.path.isfile(file_path):
                # 第二个参数 arcname 可以指定文件在 ZIP 内的路径（此处仅保留文件名，不包含目录）
                zf.write(file_path, arcname=os.path.basename(file_path))
                print(f"已添加: {file_path}")
            else:
                print(f"文件不存在，跳过: {file_path}")
                
# def group_pdfs_by_district(districts, pdf_files, base_dir='',date= data2str(datetime.datetime.now())):
#     """
#     将 PDF 文件名按开头两个字符（区县名称）分组，并可选添加基础路径。

#     :param districts: 区县名称列表（两个字符）
#     :param pdf_files: PDF 文件名列表（仅文件名，不含路径）
#     :param base_dir: 基础路径，如果提供，则与文件名拼接成完整路径存储
#     :return: 字典，键为区县名，值为该区县对应的完整路径列表
#     """
    # district_set = set(districts)
    # result = {district: [] for district in districts}

    # for file in pdf_files:
    #     # 取文件名前两个字符作为区县标识（如果文件名包含路径，建议先用 os.path.basename 提取）
    
    #     match file[:2]:
    #         case x if x in district_lishui:
    #             prefix = '丽水' + file[:2] + date
    #         case x if x in district_shaoxing:
    #             prefix = '绍兴' + file[:2] + date
    #         case x if x in district_wenzhou:
    #             prefix = '温州' + file[:2] + date
    #         case x if x in district_taizhou:
    #             prefix = '台州' + file[:2] + date
    #         case x if x in district_jinhua:
    #             prefix = '金华' + file[:2] + date
    #         case x if x in district_quzhou:
    #             prefix = '衢州' + file[:2] + date

    #     if prefix in district_set:
    #         if base_dir:
    #             # 拼接完整路径
    #             full_path = os.path.join(base_dir, file)
    #         else:
    #             full_path = file
    #         result[prefix].append(full_path)
    #     # 如果需要保留未匹配的文件，可以添加 else 分支
    # return result

def delete_pdf_files(directory):
# 获取指定目录下的所有文件和文件夹
    file_list = os.listdir(directory)
    for file in file_list:
        file_path = os.path.join(directory, file)  # 拼接完整路径
        # 检查是否为文件且扩展名为.pdf
        if os.path.isfile(file_path) and file.endswith('.pdf'):
            os.remove(file_path)  # 删除文件
            
def excel2pdf(list:pandas.DataFrame, path):
    env = jinja2.Environment(loader=jinja2.FileSystemLoader(path / 'temp'))
    template = env.get_template('template.html')
    
    
    options = {
        '--enable-local-file-access': None,  # 允许加载本地图片/CSS
        '--page-size': 'A4',
    }
    for i in range(list.shape[0]):
        html = template.render( 
                            name=list.loc[i,'客户名称'],
                            brand=list.loc[i,'设备品牌'],
                            serial_number=list.loc[i,'设备条码'],
                            device_name=list.loc[i,'设备名称'],
                            repair_date=list.loc[i,'维修日期'],
                            device_type=list.loc[i,'设备类别'],
                            reason=list.loc[i,'故障原因定位分析'],
                            analyst=list.loc[i,'分析人'],
                            reviewer=list.loc[i,'审核人'],
                            date=list.loc[i,'返回日期'],
                            pic=path / 'img' / list.loc[i,'报废图片'])
        pdffullname = str(path) + '/pdf/' + list.loc[i,'客户名称'] + list.loc[i,'设备条码'] + '.pdf'
        print(list.loc[i,'客户名称'] + list.loc[i,'设备条码'] + '.pdf'+'创建成功')
        pdfkit.from_string(html,pdffullname, options=options)
    