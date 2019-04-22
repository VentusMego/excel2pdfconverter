#! Python 3
# Vision 0.1
# 将子文件夹中的.doc文件批量转为原路径的.pdf文件
# -*- encoding: utf-8 -*-
import  os, string
from win32com import client
#pip instatll win32com
def doc2pdf(doc_name, pdf_name):
    """
    :word文件转pdf
    :param doc_name word文件名称
    :param pdf_name 转换后pdf文件名称
    """
    try:
        word = client.DispatchEx("Word.Application")
        if os.path.exists(pdf_name): # 是否存在
            os.remove(pdf_name)     # 删除一个文件
        worddoc = word.Documents.Open(doc_name,ReadOnly = 1)
       # print('open')
        worddoc.SaveAs(pdf_name, FileFormat = 17)
       # print('saved')
        worddoc.Close()
        return pdf_name
    except:
        return 1


def excel2pdf(excel_name, pdf_name):
    """
    :excel转pdf 存疑 修改
    """
    try:
        xl = client.Dispatch("Excel.Application")
        if os.path.exists(pdf_name):
            os.remove(pdf_name)
        excels = xl.Woekbooks.Open(excel_name,ReadOnly = 1)
        excels.ExportAsFixedFormat(0, pdf_name)
        excels.Quit()
        return pdf_name
    except:
        return 1


# 以下为主程序
if __name__=='__main__':
    for folderName, subfolders, filenames in os.walk('F:\\14、17231AD设计文本'):#<-- Type folder path here.
        print('现在在处理的文件夹为' + folderName)
        for filename in filenames:
            file_path_now = folderName + filename #绝对路径+文件名+扩展名
            print(file_path_now + '为合成后的文件')
            (filepath, tempfilename) = os.path.split(file_path_now)
            (filenameshort, extension) = os.path.splitext(tempfilename)
            if(extension == '.doc'):
                print(filename)
                doc2pdf(file_path_now  , file_path_now + '.pdf' )
                print(filename + '.pdf 打印成功')
            
            
                
#    doc_name = "f:/test.xls"
#    ftp_name = "f:/test.pdf"
#    doc2pdf(doc_name, ftp_name)
    print ("全然OK")
    
