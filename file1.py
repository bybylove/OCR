from PIL import Image
import sys
import pyocr
import re
import xlwt
import os
newregex = re.compile(r'''企业注册号\s?二\s?(.*?)\s企业名称\s?二\s?(.+)''')
rename = re.compile(r'''有眼公司''')
rename2 = re.compile(r'''眼''')
rename3 = re.compile(r'''丁''')
rename4 = re.compile(r'''\s''')
workbook = xlwt.Workbook(encoding='ascii')
worksheet = workbook.add_sheet('mysheet')
filelist = os.listdir("images/")
if __name__ == '__main__':
    tools = pyocr.get_available_tools()[:]
    if len(tools)==0:
        print("no ocr tool found")
        sys.exit(1)
    else:
        pass
    a = 0
    for k in filelist:
        a+=1
        url = 'images/' +k
        try :
            image=Image.open(url)
            if image.mode == 'RGBA':
                image = image.crop((0,0,600,75))
                w,h = image.size
                for i in range(0,w):
                    for j in range (0,h):
                        if image.getpixel((i,j))[1]>150 and image.getpixel((i,j))[3]>150 :
                            image.putpixel((i,j),(255,255,255,255))
                        elif image.getpixel((i,j))[1]>150 and image.getpixel((i,j))[3]<150:
                            image.putpixel((i,j),(255,255,255,255))
                        elif image.getpixel((i,j))[1]<150 and image.getpixel((i,j))[3]>150:
                            image.putpixel((i,j),(0,0,0,255))
                        else :
                            image.putpixel((i,j),(255,255,255,255))
                result = tools[0].image_to_string(image,lang = 'chi_sim')
                newmatch = newregex.findall(result)
                id = newmatch[0][0]
                name = newmatch[0][1]
                id = rename3.sub('T',id)

                name = rename.sub('有限公司', name)
                name = rename2.sub('服', name)
                name = rename4.sub('',name)
                id = rename4.sub('', id)
                print(id,name)
                worksheet.write(a, 0, id)
                worksheet.write(a, 1, name)
        except:
            continue
    workbook.save('result.xls')

