import shutil,os
import pdfplumber
import xlwt

def pdf_xls(path):
    # 定义保存Excel的位置
    workbook = xlwt.Workbook()  #定义workbook
    sheet = workbook.add_sheet('Sheet1')  #添加sheet
    
    #path = r"C:\Users\he\Desktop\桂格\150018606\华润卫国道-桂格全品.PDF"  
    pdf = pdfplumber.open(path)
    print('\n')
    print('开始读取数据')
    print('\n')
    i = 0

    for page in pdf.pages:
        all_text = page.extract_text().split('\n')
        for line_2 in all_text:
            if i == 0:
                line_1 = line_2.split( )
                #print(line_1)     
                for j in range(len(line_1)):
                    sheet.write(i, j,line_1[j])
            else:
                line_1 = line_2.split( )
                if len(line_1) == 1:
                    sheet.write(i, 5,line_1)
                    #print(line_1)
                elif len(line_1) == 2:
                    sheet.write(i,4,line_1[0])
                    sheet.write(i,5,line_1[1])
                else:
                    line_1[6] = line_1[6] + ' ' +line_1[7]
                    del line_1[7]
                    for j in range(len(line_1)):
                        sheet.write(i, j,line_1[j])
                    #print(line_1)
            i += 1

    pdf.close()

    # 保存Excel表
    save_path = path.replace('pdf','xls')
    workbook.save(save_path)
    print('\n')
    print('写入excel成功')
    print('保存位置：')
    print(save_path)
    print('\n')
    






def get_dir(path,fileType):

	#查看当前目录文件列表（包含文件夹）
    allfilelist = os.listdir(path)
	
    for file in allfilelist:
        print(file,'\n')
        filepath = os.path.join(path, file)
        #判断是否是文件夹，如果是则继续遍历，否则打印信息
        if os.path.isdir(filepath):
            allfilelist2 = os.listdir(filepath)
            for file2 in allfilelist2:
                filepath3 = os.path.join(filepath, file2)
                #判断文件是否以.avi结尾
                if filepath3.endswith(fileType):
                    print('找到文件：'+filepath3)
                    pdf_xls(filepath3)               
        else:
            print('不是文件夹，继续查找...')


path = r'C:\Users\he\Desktop\桂格'
get_dir(path,'.pdf')