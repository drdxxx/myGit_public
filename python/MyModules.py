'''
时间：2020.6.15
作者：drd
'''

import os
import re
import imghdr
import qrcode
import requests
import filetype
import zipfile
import win32com.client as win
from requests import RequestException
from PIL import Image

#从指定url获取html页面内容
def get_page(url,headers=None,mode='get'):
    try:
        if mode == 'get':
            response = requests.get(url,headers)
        elif mode == 'post':
            response = requests.post(url,headers)
        if response.status_code == 200:
            return response.text
        else:
            print('请求失败，状态码：%d'%response.status_code)
            return -1
    except RequestException:
        return -1

#对html文本进行正则匹配
def parse_page(reg,html):
    pattern = re.compile(reg, re.S)
    items = re.findall(pattern, html)
    if items != []:
        return items
    else:
        print('没有找到符合正则表达式的数据！')
        return None

#将数据写入文件
def write_to_file(file_path,content,mode='w',encoding='utf-8'):
    with open(file_path, mode,encoding=encoding)as f:
        f.write(content)

#判断指定图片是否损坏 若图片损坏，则删除，否则返回图片类型
def isImageDamage(img_path):
    if imghdr.what(img_path) == None:
        #图片已损坏
        os.remove(img_path)
        print('一张图片损坏，已删除！')
        return None
    else:
        return imghdr.what(img_path)

#从指定的链接地址下载一张图片,下载成功返回1，下载失败返回0；url:图片的链接,img_path：图片保存的位置
def get_oneImage(url,img_path):
    try:
        image = requests.get(url)
        with open(img_path, "wb") as f:
            f.write(image.content)
        isImageDamage(img_path)
        return 1
    except:
        return 0

#图片格式转换为JPG
#source_dir：图片文件夹（所有文件必须是图片）或单张图片的位置
#isDeleteRes：布尔量，是否删除原文件，默认False
def other_to_jpg(source_dir,isDeleteRes=False):
    if os.path.isdir(source_dir):
        flist = os.listdir(source_dir)
        for f in flist:
            img = Image.open(source_dir + '\\' + f,'r')
            if Image.isImageType(img):
                try:
                    img.save((source_dir + '\\' + f)[:-4] + '.jpg')
                    if type(isDeleteRes) == bool:
                        if isDeleteRes == True:os.remove(source_dir + '\\' + f)
                except OSError:
                    os.remove((source_dir + '\\' + f)[:-4] + '.jpg')
                    print('convert_to_jpg error:%s'%f)
    else:
        img = Image.open(source_dir,'r')
        try:
            img.save(source_dir[:-4] + '.jpg')
            if type(isDeleteRes) == bool:
                if isDeleteRes == True: os.remove(source_dir)
        except OSError:
            os.remove(source_dir[:-4] + '.jpg')
            print('convert_to_jpg error！')

#图片格式转换为PNG
#source_dir：图片文件夹（所有文件必须是图片）或单张图片的位置
#isDeleteRes：布尔量，是否删除原文件，默认False
def other_to_png(source_dir,isDeleteRes=False):
    if os.path.isdir(source_dir):
        flist = os.listdir(source_dir)
        for f in flist:
            img = Image.open(source_dir + '\\' + f,'r')
            if Image.isImageType(img):
                try:
                    img.save((source_dir + '\\' + f)[:-4] + '.png')
                    if type(isDeleteRes) == bool:
                        if isDeleteRes == True:os.remove(source_dir + '\\' + f)
                except OSError:
                    os.remove((source_dir + '\\' + f)[:-4] + '.png')
                    print('convert_to_png error:%s'%f)
    else:
        img = Image.open(source_dir,'r')
        try:
            img.save(source_dir[:-4] + '.png')
            if type(isDeleteRes) == bool:
                if isDeleteRes == True: os.remove(source_dir)
        except OSError:
            os.remove(source_dir[:-4] + '.png')
            print('convert_to_png error！')

#生成普通二维码
#target_str:目标数据，可以是url，也可以是文本字符串
#qrcode_path:二维码图片保存的位置
def create_normalQrcode(target_str,qrcode_path):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(target_str)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="blue",back_color="white")
    qr_img.save(qrcode_path)

#生成带图片的二维码
#target_str:目标数据，可以是url，也可以是文本字符串
#img_path:二维码中的图片的位置，注意：图片只能是png格式
#qrcode_path:二维码图片保存的位置
def create_withImgQrcode(target_str,img_path,qrcode_path):
    qr = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=1
    )
    qr.add_data(target_str)
    qr.make(fit=True)
    img = qr.make_image()
    img = img.convert("RGBA")
    icon = Image.open(img_path)
    img_w, img_h = img.size
    factor = 4
    size_w = int(img_w / factor)
    size_h = int(img_h / factor)
    icon_w, icon_h = icon.size
    if icon_w > size_w:
        icon_w = size_w
    if icon_h > size_h:
        icon_h = size_h
    icon = icon.resize((icon_w, icon_h), Image.ANTIALIAS)
    w = int((img_w - icon_w) / 2)
    h = int((img_h - icon_h) / 2)
    img.paste(icon, (w, h), icon)
    img.save(qrcode_path)

#读取指定压缩包中的文件，显示压缩包中的文件名，fname：压缩包路径xxx.zip，fileto：文件解压到哪个位置,默认是当前文件夹
def readZip(fname,fileto='./'):
    zip = zipfile.ZipFile(fname)
    print('%s压缩包中的文件:'%fname)
    flist = zip.namelist()
    for i in range(len(flist)):
        print(i,':',flist[i])
    t = input('是否解压全部文件到指定目录（y/n/q（表示退出））？')
    if t == 'y':
        #解压所有文件到指定路径
        for  f in flist:
            zip.extract(f,fileto)
    elif t == 'q':
        #退出程序
        zip.close()
        return None
    else:
        files = input('请输入要解压的文件路径，以英文逗号隔开：')
        fls = list(files.split(','))
        for f in flist:
            if f in fls:
                zip.extract(f, fileto)
    print('解压完成！')
    zip.close()

#识别文件或文件夹类型（支持文件夹，图像，视频，音频……，不支持wps文件类型：doc(x).xls(x).ppt(x)）
#返回值：‘None’-不能识别的类型，‘dir’-文件夹，其他-可识别类型
def identify_type(source_dir):
    if os.path.isdir(source_dir):
        return 'dir'
    else:
        fobj = filetype.guess(source_dir)
        if fobj is None:return None
        else:return fobj.extension

#txt文件内容去重
#txt_path:txt文件路径
#encoding:文件字符编码，默认为utf-8
def txt_deduplicate(txt_path,encoding='utf-8'):
    if os.path.isfile(txt_path) and txt_path[-4:] == '.txt':
        with open(txt_path, "r",encoding=encoding) as f:
            txt_content = f.readlines()
            for i in range(len(txt_content)):
                for j in range(i+1,len(txt_content)):
                    if txt_content[i] == txt_content[j]:txt_content[j] = ''
        os.remove(txt_path)
        with open(txt_path, "a",encoding=encoding) as f:
            for line in txt_content:
                if line != '':f.write(line)
    else:
        print('找不到指定txt文件！')

#语音播放txt文件的内容
#txt_path:txt文件路径
#line:播放哪一行数据，默认为0，表示全部
#encoding:文件字符编码，默认为utf-8
def txt_to_sound(txt_path,line=0,encoding='utf-8'):
    if os.path.isfile(txt_path) and txt_path[-4:] == '.txt':
        with open(txt_path, "r", encoding=encoding) as f:
            txt_content = f.readlines()
        if line == 0:
            sound = win.Dispatch('SAPI.SpVoice')
            for text in txt_content:
                sound.Speak(text)
        else:
            sound = win.Dispatch('SAPI.SpVoice')
            for i in range(len(txt_content)):
                if i == line - 1:
                    sound.Speak(txt_content[i])
    else:
        print('找不到指定txt文件！')


#if __name__ == '__main__':
  #  txt_deduplicate('ttt.txt')
  #  txt_to_sound(r'ttt.txt',2,'gbk')