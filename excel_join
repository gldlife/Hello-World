"""
encoding:utf-8
author:李建飞
version:
"""
from openpyxl import load_workbook
from openpyxl import Workbook
import os

#读取整个sheet表数据并返回列表
#提供sheet表为参数
def get_sheet_datas(sheet):
    datas=[]
    max=sheet.max_column
    for row in sheet.rows:
        rows=[]
        for item in row:
            rows.append(item.value)
        target_ws.append(rows)
    return target_ws

def show_select():
    print()
    print('程序名称：EXCEL工作表合成器\n作者：李建飞\n版本：v1.0\n发布日期：2019-1-17')
    print('='*40)
    print('1、合并一个工作簿内的多张sheet表；')
    print('2、合并指定文件夹内的所有工作簿；')
    print('3、退出程序')
    print('使用说明：合并的多个工作簿或多张sheet表需要具有相同的表头。')
    print('='*40)

if __name__=='__main__':
    while(True):
        show_select()
        choice=input('请输入你想进行的操作：')
        if(choice=='1'):
            source=input('请输入文件存放位置\n（输入示例：d:/document/grade.xlsx)\n').replace('\\','/')
            if(source==''):
                print('未输入文件存放位置，请重新输入')
                continue
            #输出文件名称
            target_name=input('请输入合成后的文件名称：')
            if(target_name==''):
                target_name='合并文件'
            target_path=target_name+'.xlsx'
            # 加载目标excel文件
            source_wb = load_workbook(source)
            # 创建目标文件对象
            target_wb = Workbook()
            target_ws = target_wb.active
            # 获取工作簿中所有sheet表的名称
            sheetnames = source_wb.sheetnames
            for name in sheetnames:
                get_sheet_datas(source_wb[name])
            # 文件存储
            target_wb.save(target_path)
            print('~合并成功~')
        elif(choice=='2'):
            source_path=input('输入文件存放的路径(路径最后要加上斜杠\）：').replace('\\','/')
            if(source_path==''):
                print('未输入文件夹路径，请重新输入')
                continue
            if(source_path[-1]!='/'):
                source_path=source_path+'/'
            # 输出文件名称
            target_name = input('请输入合成后的文件名称：')
            if (target_name == ''):
                target_name = '合并文件'
            target_path = target_name + '.xlsx'
            book_names=os.listdir(source_path)
            # 创建目标文件对象
            target_wb = Workbook()
            target_ws = target_wb.active
            for name in book_names:
                source_name=os.path.join(source_path,name)
                source_wb=load_workbook(source_name)
                get_sheet_datas(source_wb.active)
            # 文件存储
            target_wb.save(target_path)
            print('~合并成功~')
        elif(choice=='3'):
            exit('~退出程序~')
        else:
            print('序号输入错误！请重新输入~')






