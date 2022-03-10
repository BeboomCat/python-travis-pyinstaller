#你的python代码
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128
##from reportlab.graphics.barcode import eanbc, qr, usps
from reportlab.graphics.shapes import Drawing 
from reportlab.lib.units import mm
from reportlab.graphics import renderPDF
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm,inch
import fitz,os,sys,time,xlrd
##import win32print
from PIL import Image
import tkinter as tk
from tkinter import filedialog

def signContent(mytxt):
    signtext.delete(0.0,tk.END)
    signtext.insert('insert',mytxt)

def number_Check(ocx):
    data = ocx.get()
    if not dataisdigit():
        ocx.set('')
        signtext.delete(0.0,tk.END)
        signtext.insert('insert',"输入错误，请检查输入是否为整数...")
        

def chooseFile():
    folder_path = filedialog.askopenfilename() #选择文件
    (filepath,filename) = os.path.split(folder_path)
    (name, suffix) = os.path.splitext(filename)
    if folder_path =="":
        return 0
    elif not folder_path.endswith(('.xlsx','.xls')):
        signtext.delete(0.0,tk.END)
        signtext.insert('insert',"小宝贝~文件选错了哦\n要选择xlsx或者xls格式的excel表呀...")
        return 0
    else:
        file_path_content.set(folder_path)
        signtext.delete(0.0,tk.END)

def openxlsx(excel_index = 0, exception=1):
    excel_path = file_path.get()
    excel = xlrd.open_workbook(excel_path,encoding_override="utf-8")
    sheet = excel.sheet_by_index(excel_index)
    main_list = []
    main_nrows = sheet.nrows
    for num in range(exception,main_nrows):
        main_list.append(sheet.row_values(num))
    
    return main_list

def pdftoimage(choose_dir,name):
    
    pdfDoc = fitz.open(choose_dir)
    for pg in range(pdfDoc.pageCount):
        page = pdfDoc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
        # 此处若是不做设置，默认图片大小为：792X612, dpi=96
        zoom_x = 4.33333333 #(1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 4.33333333
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)
            
##            if not os.path.exists(out_path):#判断存放图片的文件夹是否存在
##                os.makedirs(out_path) # 若图片文件夹不存在就创建

        pix.writeImage(".//条码贴纸//"+str(name)+'.jpg')#将图片写入指定的文件夹内
    

def createBarCodes(c,nnum,name,phone_num,addr,barWidth,barHeight):
    barcode_value = nnum
    if barcode_value.isdigit():
        myWidth = 0.25*mm
        x = 0.5*mm
    else:
        myWidth = 0.23*mm
        x = 0.2*mm
    barcode128 = code128.Code128(barcode_value,barHeight=12*mm,barWidth = myWidth)

    y = 4 * mm
    barcode128.drawOn(c, x, y)
    c.drawString(3*mm,35*mm,"姓名：")
    c.drawString(12*mm,35*mm,name)
    c.drawString(27*mm,35*mm,"性别：")
    if int(nnum[16:17]) % 2 == 0:
        c.drawString(37*mm,35*mm,"女")
    else:
        c.drawString(37*mm,35*mm,"男")
        
    c.drawString(44*mm,35*mm,"年龄：")
    c.drawString(54*mm,35*mm,str(int(time.strftime("%Y", time.localtime())) - int(nnum[6:10])))
    c.drawString(3*mm,28*mm,"电话：")
    try:
        if phone_num.isdigit():
            c.drawString(12*mm,28*mm,phone_num)
    except:
        pass
##        c.drawString(12*mm,28*mm,phone_num)
    c.drawString(3*mm,21*mm,"地址：")
    c.drawString(12*mm,21*mm,addr)
    #showPage函数：保存当前页的canvas
    c.showPage()
    #save函数：保存文件并关闭canvas
    c.save()

def creat_bar():
    if True:
        excel_path = file_path.get()
        if excel_path == "":
            signContent("小宝贝，你还未选择数据文件哦\n选择数据后再试吧~")
            return 0
        pdfmetrics.registerFont(TTFont("方正小标宋简体", "FZXBSJW.TTF"))#SimSun.ttf
    ##    pdfmetrics.registerFont(TTFont("SimSun", "SimSun.ttf"))#SimSun.ttf
        if sticker_width.get()=="" or sticker_height.get()=="" or name_serial.get()==""\
           or id_serial.get()=="" or addr_serial.get()=="" or phone_serial.get()=="":
            signContent("上面的空格都是必填的呀\n选择空白的列填上去也是可以的哦\n填好了再试试吧~")
            return 0
        mydata = openxlsx()
        stickerWidth = int(sticker_width.get())
        stickerHeight = int(sticker_height.get())
        nameCol = int(name_serial.get())
        idCol = int(id_serial.get())
        addrCol = int(addr_serial.get())
        phoneCol = int(phone_serial.get())
        
        for people in mydata:
            p_id = str(people[idCol])
    ##        print("p_id=",p_id)
            c=canvas.Canvas("temp.pdf",pagesize=(stickerWidth*mm,stickerHeight*mm))
            c.setFont("方正小标宋简体", 10,3)
    ##        调用函数生成条形码和二维码，并将canvas对象作为参数传递
            createBarCodes(c,p_id,people[nameCol],people[phoneCol],people[addrCol],stickerWidth,stickerHeight)
            pdftoimage("temp.pdf",people[nameCol]+str(people[idCol]))
        if os.path.exists("temp.pdf"):
            os.remove("temp.pdf")
        signContent("条码生成完毕啦！\n请到【条码贴纸】这个文件夹看看成果吧。")
##    except:
##        signContent("啊哦...好像出错了哦\n请先检查所在列的数据是否正确（很有可能是你的数据有问题哦）\n或者找小工具开发者看看吧。")

def creatFold():
    if not os.path.exists("./条码贴纸"):
        os.makedirs("./条码贴纸")

if __name__ == '__main__':
    creatFold()
    
    window = tk.Tk()
    window.title("【条形码贴纸生成】--倒水卫生院信息开发")
    window.geometry("360x350")
    window.resizable(0,0)
    adjustment_x = 80

    file_choose = tk.Button(window, text= "选择文件",command = chooseFile)
    file_choose.pack()
    file_choose.place(x=5, y=10)
    file_path_content = tk.StringVar()
    file_path = tk.Entry(window,textvariable = file_path_content,state = "disabled", width = 40)
    file_path.place(x=70,y=15)

    
    sticker_tag = tk.Label(window, text="贴纸规格长：",command = None).place(x=0+adjustment_x,y=50)
    sticker_width = tk.Entry(window,textvariable = tk.StringVar(window,value = "60"), width = 5)
    sticker_width.place(x=75+adjustment_x,y=50)
    sticker_tag_1 = tk.Label(window, text="宽：",command = None).place(x=120+adjustment_x,y=50)
    sticker_height = tk.Entry(window, textvariable = tk.StringVar(window,value = "40"), width = 5)
    sticker_height.place(x=150+adjustment_x,y=50)


    name_tag = tk.Label(window, text="名字所在列数：",command = None).place(x=5+adjustment_x,y=90)
    name_serial = tk.Entry(window,textvariable = tk.StringVar(window,value = "0"), width = 5)
##    name_serial.bind('<KeyRelease>', lambda ocx =name_serial: number_Check(ocx))
    name_serial.place(x=90+adjustment_x,y=90)

    id_tag = tk.Label(window, text="身份证所在列数：",command = None).place(x=0+adjustment_x,y=120)
    id_serial = tk.Entry(window,textvariable = tk.StringVar(window,value = "3"), width = 5)
    id_serial.place(x=90+adjustment_x,y=120)

    addr_tag = tk.Label(window, text="地址所在列数：",command = None).place(x=5+adjustment_x,y=150)
    addr_serial = tk.Entry(window,textvariable = tk.StringVar(window,value = "8"), width = 5)
    addr_serial.place(x=90+adjustment_x,y=150)

    phone_tag = tk.Label(window, text="电话所在列数：",command = None).place(x=5+adjustment_x,y=180)
    phone_serial = tk.Entry(window,textvariable = tk.StringVar(window,value = "5"), width = 5)
    phone_serial.place(x=90+adjustment_x,y=180)

    create_barcode = tk.Button(window, text= "生成贴纸",width = 10,height = 1,bd = 5, command = creat_bar)
    create_barcode.pack()
    create_barcode.place(x=60+adjustment_x,y=230)

    signtext = tk.Text(window,width = 45)
    signtext.pack()
    signtext.place(height = 70,x=20,y=270)

    window.mainloop()
 
