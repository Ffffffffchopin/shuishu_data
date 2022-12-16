#用pdf2docx模块将pdf文件转换成word方便后续提取表格和图片


#引入Converter类
from pdf2docx import Converter


if __name__=='__main__':
    #pdf文件路径
    pdf_file="16263-shuishu.pdf"
    #创建Converter对象
    cv=Converter(pdf_file)
    #convert方法全部转换
    cv.convert("output.docx",start=0)
    #关闭对象
    cv.close

#转换完成之后手动把word文档output.docx中的对照表表格提取出来成table.docx方便后续处理