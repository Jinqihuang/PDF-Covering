import os
import gc
import shutil
import importlib,sys
importlib.reload(sys)

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import *
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed


file_list = []
file_name = []

#遍历文件下所有的文件
def find_dir(dir_name):
    #判断输入路径是否存在
    if os.path.exists(dir_name):
        
        for path_name, dir_list, files_name in os.walk(dir_name):    

            for filelist in files_name:

                file=os.path.splitext(filelist)

                filename,type=file
                #文件类型判断是为PDF文件
                filetype = '.pdf'
               
                if filetype == type:
                    #PDF路径加入列表
                    file_list.append(path_name+'/'+filelist)
                    #PDF名称加入列表
                    file_name.append(filelist)

                else:
                    #提示文件不是PDF文件
                    print('The '+filelist+' not pdf File ')
    #否则新增一个路径 
    else:
        os.mkdir('./sop')
    #指定文档输出路径
    path_output = './output'
    #判断输出路径是否存在
    if os.path.exists(path_output):
        print(path_output+' existence.')
    else:
        os.mkdir(path_output)
    

    


def create_watermark(x,y):
    #新建一个默认空白文件
    c = canvas.Canvas("mark-none.pdf", pagesize = (x,y))
    #移动坐标原点(坐标系左下为(0,0))

    #设置字体
    c.setFont("Helvetica", 0)
    #指定描边的颜色
    c.setStrokeColorRGB(1, 1, 1)
    #指定填充颜色
    c.setFillColorRGB(1, 1, 1)
    #画一个矩形
    c.rect(0, 0, 0, 0, fill=1)
                                                                                      
    #关闭并保存pdf文件
    c.save()
 
def create_watermark_pdf(input_pdf, input_name,output,watermark):
    #指定合拼PDF的文件
    watermark_obj=PdfFileReader(watermark)
    watermark_page=watermark_obj.getPage(0)
    #指定合拼的源文件
    pdf_reader = PdfFileReader(input_pdf)
    pdf_writer = PdfFileWriter()
    #循环每一页，插入空白页面
    a = 0
    for page in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page)
        page.mergePage(watermark_page)
        pdf_writer.addPage(page)
        a += 1
        #获取完成的页数
        rate = float(a)/float(pdf_reader.getNumPages())
        #输入完成的进度
        print('File conversion completed :'+'%.2f%%' % (rate * 100)+'.')
    #完成输出文件
    with open(output, 'wb') as out:
        pdf_writer.write(out)
        print('The '+file_name[input_name]+' file first add watermark successful.')




def parse(DataIO,code_source,code_new):

    #要删除所有文件（只是删除文件 而不是文件夹），所以 我们肯定要遍历这个文件目录 （for  in遍历)
    path = './mark'
    if os.path.exists(path):
        for i in os.listdir(path):
            path_file = os.path.join(path,i)
            if os.path.isfile(path_file):
                os.remove(path_file)
            else:
                for f in os.listdir(path_file):  
                    path_file2 =os.path.join(path_file,f)
                    if os.path.isfile(path_file2):
                        os.remove(path_file2)
    else:
        os.mkdir('./mark')

 
    #用文件对象创建一个PDF文档分析器
    parser = PDFParser(DataIO)
    #创建一个PDF文档
    doc = PDFDocument()
    #分析器和文档相互连接
    parser.set_document(doc)
    doc.set_parser(parser)
    #提供初始化密码，没有默认为空
    doc.initialize()
    #检查文档是否可以转成TXT，如果不可以就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        #创建PDF资源管理器，来管理共享资源
        rsrcmagr = PDFResourceManager()
        #创建一个PDF设备对象
        laparams = LAParams()
        #将资源管理器和设备对象聚合
        device = PDFPageAggregator(rsrcmagr, laparams=laparams)
        #创建一个PDF解释器对象
        interpreter = PDFPageInterpreter(rsrcmagr, device)

        pg = 0
        #循环遍历列表，每次处理一个page内容
        #print(doc.get_pages())
        #获取page列表
        for page in doc.get_pages():
            
            size_x = []
            size_y = []
            size_font = []
            layout_x = 2448
            layout_y = 1584
            try:
                interpreter.process_page(page)
                #接收该页面的LTPage对象
                layout = device.get_result()
                #这里的layout是一个LTPage对象 里面存放着page解析出来的各种对象
                #一般包括LTTextBox，LTFigure，LTImage，LTTextBoxHorizontal等等一些对像
                #想要获取文本就得获取对象的text属性
                layout_x =int(layout.bbox[2])
                layout_y =int(layout.bbox[3])
                
                
                for x in layout:
               
                    #读取所有文本
                    if(isinstance(x, LTTextBoxHorizontal)):
                        for line in x:
                            #读取所有文本列
                            if isinstance(line,LTTextLine):
                                #读取每一文本列
                                result = line.get_text().lower()
                                #设置字符串匹配起始值
                                num_code = 0
                                #判断字符串中有几个匹配的值,循环几次
                                for i in  range(result.count(code_source.lower())):
                                    #关键字匹配，并获取在文本列的第几位
                                    codetwo = result.find(code_source,num_code)
                                    #下次从这个值往后开始匹配
                                    num_code = codetwo+1
                                    i = 0
                                    for char in line:
                                        #读取每一个字符
                                        if isinstance(char,LTChar):
                                            #print(char)
                                            if i == codetwo:
                                                c = char.get_text()
                                                size_x.append(float(char.bbox[0]))
                                                size_y.append(float(char.bbox[1]))
                                                size_font.append(float(char.size))
                                        i += 1
                            
            except:
                    print("The watermark file Create Failed")
                    return 1
            #根据匹配文本的坐标系，生成替换水印
            #默认大小为      
            mark = canvas.Canvas("mark/"+str(pg)+".pdf", pagesize = (layout_x,layout_y))
            
            j = 0
            for i in size_x:

                #移动坐标原点(坐标系左下为(0,0))
                x = float(i)
                y = float(size_y[j])

                #指定描边的颜色
                mark.setStrokeColorRGB(1, 1, 1)
                #指定填充颜色
                mark.setFillColorRGB(1, 1, 1)
                #画一个矩形
                mark.rect(x, y, size_font[j]*2.8, size_font[j]*0.8 ,fill=1)
                j +=1
            z = 0  
            for i in size_x:
                x = float(i)
                y = float(size_y[z])
                #设置字体
                mark.setFont("Helvetica", size_font[z]*0.8)
                    
                #指定填充颜色
                mark.setFillColorRGB(0, 0, 0)
                #设置透明度，1为不透明
                mark.setFillAlpha(1)
                #新建一个文本，输入为全大写
                mark.drawString(x+(size_font[z]*0.1), y+(size_font[z]*0.125), code_new.upper())
                #匹配是否为大文件，选择遮盖logo
                if size_font[z] > 20:
                    #指定描边的颜色
                    mark.setStrokeColorRGB(1, 1, 1)
                    #指定填充颜色
                    mark.setFillColorRGB(1, 1, 1)
                    #画一个矩形
                    mark.rect(x-35, y-2, 35, size_font[z]*0.9,fill=1)
                z +=1
                    
            #画一个空白矩阵
            mark.rect(0, 0, 0, 0, fill=1)
                
            #关闭并保存pdf文件
            mark.save()
            pg += 1

        print('The watermark file Create Successful.')



def create_watermark_pdf_add(input_pdf, input_name, output):
    #指定合拼PDF的源文件
    pdf_reader = PdfFileReader(input_pdf)
    pdf_writer = PdfFileWriter()
    a = 0
    #插入匹配成功生成的水印文件
    for page in range(pdf_reader.getNumPages()):
        
        page = pdf_reader.getPage(page)
        #遍历mark目录下的生存的水印文件
        watermark_obj=PdfFileReader('mark/'+str(a)+'.pdf')
        watermark_page=watermark_obj.getPage(0)
        page.mergePage(watermark_page)
        pdf_writer.addPage(page)
        a += 1
        #print(pdf_reader.getNumPages())
        #获取完成的页数
        rate = float(a)/float(pdf_reader.getNumPages())
        #输入完成的进度
        print('File conversion completed :'+'%.2f%%' % (rate * 100)+'.')
        
        #break
        
    with open(output, 'wb') as out:
        #设置密码
        #pdf_writer.encrypt('pega.1234')                 
        pdf_writer.write(out)
        #输出完成信息
        print('The '+file_name[input_name]+' File second add watermark successful.')



if __name__ == '__main__':
    find_dir('./sop')
    #新建一个空白页面
    create_watermark(2448,1584)

    print('Mark-nono.pdf Create Successful.')

    code_source = input ('请输入要替换的 原始 文字（Enter the source code that needs to be replaced）:')
    code_new = input ('请输入要替换成的 新的 文字（Need to output the new code）:')
    
    i = -1
    for sop in file_list:

        i += 1
        print('The '+file_name[i]+' File Create Work.')

        create_watermark_pdf(input_pdf = sop,input_name = i,output = sop,watermark = 'mark-none.pdf')

        print('********************************One Steps Completed***************************** ')

        #解析本地PDF文本，保存到本地TXT
        with open(sop,'rb') as pdf_html:

            error = parse(pdf_html,code_source,code_new)

        print('********************************Two Steps Completed***************************** ')
        #解析PDF文件如无法解析或者中途有异常则取消此文件的变更
        if error == 1 :
            print('This '+file_name[i]+' file is protected and cannot be read.')
            continue
        #正常覆盖
        else:
            create_watermark_pdf_add(input_pdf = sop,input_name = i,output = sop)

        print('********************************Three Steps Completed*************************** ')      
        #移动PDF文件到新文件下面
        shutil.copyfile(sop, './output/'+file_name[i])
        #删除已经移动PDF文件
        os.remove(sop)
        print('File '+file_name[i]+' all function successful.')
        
        print('********************************Four Steps Completed**************************** ')
        del sop
        gc.collect()

    print('All File Completed.')
        
    
