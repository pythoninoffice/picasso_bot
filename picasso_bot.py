import cv2
import xlsxwriter
import numpy as np
#import matplotlib.pyplot as plt
import xlwings as xw
import time


def rgb2hex(r,g,b):
    r = max(0,min(r,255))
    g = max(0,min(g,255))
    b = max(0,min(b,255))
    return f'#{r:02x}{g:02x}{b:02x}'

def pic2excel(img, file_name,x=1, y=1):
    if x != 1 and y !=1:
        img_resize = cv2.resize(img, None, fx=x, fy=y)
    else:
        img_resize = img
    if len(img_resize.shape) == 3:
        row,col,layer = img_resize.shape
    else:
        row,col = img_resize.shape
    wb = xlsxwriter.Workbook(file_name)
    ws = wb.add_worksheet('pic')
    r = 1
    for row in img_resize:
        c = 1
        for i in row:
            print(f'pixel at position ({r},{c} is {i})')
            if isinstance(i,np.ndarray):
                color_format = wb.add_format({'bg_color':rgb2hex(i[2],i[1],i[0])})   #b,g,r
            else:
                color_format = wb.add_format({'bg_color':rgb2hex(i,i,i)})
            ws.write(r,c,'',color_format)
            c += 1
        r += 1
    ws.set_column(first_col=1, last_col = col, width = 2)
    wb.close()
    
    
def pic2excel_xlwings(img,direction= None, x=1, y=1):
    if x != 1 and y !=1:
        img_resize = cv2.resize(img, None, fx=x, fy=y)
    else:
        img_resize = img
    if len(img_resize.shape) == 3:
        row,col,layer = img_resize.shape
    else:
        row,col = img_resize.shape
    book = xw.Book()
    sheet = book.sheets[0]
    xw.Range((1,1),(1,col)).column_width = 2
    time.sleep(5)
    if direction == 'vertical':
        c = 1
        for col in img_resize:
            r = 1
            for i in col:
                if isinstance(i,np.ndarray):
                    sheet.cells[r,c].color = (i[2],i[1],i[0])
                else:
                    sheet.cells[r,c].color = (i,i,i)
                r += 1
            c +=1
                
    elif direction == 'horizontal':
        r = 1        
        for row in img_resize:
            c = 1
            for i in row:
                print(f'pixel at {r}, {c} is color {i}')
                if isinstance(i,np.ndarray):
                    sheet.cells[r,c].color = (i[2],i[1],i[0])
                else:
                    sheet.cells[r,c].color = (i,i,i)
                c += 1
            r += 1


img2 = cv2.imread(r'where-to-stream-dragonball-z.jpg', cv2.IMREAD_COLOR)
pic2excel_xlwings(img2, direction = 'vertical',x=1,y=1)
