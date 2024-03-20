__author__ = 'swallow'
__language__= 'python 3.0'

import logging
import os
import sys
import time

from PIL import Image
from PIL import ImageChops
from PIL import ImageDraw

def compare_images(path_one, path_two, output_path, diff_save_location):

    #比较图片，如果有不同则生成展示不同的图片

    #@参数一: path_one: 第一张图片的路径
    #@参数二: path_two: 第二张图片的路径
    #@参数三: diff_save_location: 不同图的保存路径
    
    image_one = Image.open(path_one)
    image_two = Image.open(path_two)
    try: 
        diff = ImageChops.difference(image_one, image_two)
        if (image_one.size != image_two.size):
            print(image_one.size)
            print(image_two.size)

        if diff.getbbox() is None:
        # 图片间没有任何不同则直接退出
            print("【+】We are the same!")
        else:
            str = path_one.replace(".png", "")
            diff_save_location = output_path + "/" + str+ "_diff.jpg"
            diff.save(diff_save_location)
            print("please check " + diff_save_location)
    except ValueError as e:
        text = ("表示图片大小和box对应的宽度不一致，参考API说明：Pastes another image into this image."
                "The box argument is either a 2-tuple giving the upper left corner, a 4-tuple defining the left, upper, "
                "right, and lower pixel coordinate, or None (same as (0, 0)). If a 4-tuple is given, the size of the pasted "
                "image must match the size of the region.使用2纬的box避免上述问题")
        print("【{0}】{1}".format(e,text))

def IsValidImage(img_path):
    """
    判断文件是否为有效（完整）的图片
    :param img_path:图片路径
    :return:True：有效 False：无效
    """
    bValid = True
    try:
        Image.MAX_IMAGE_PIXELS = None
        Image.open(img_path).verify()
    except Exception as e:
        print('Failed to open img: '+ str(e))
        bValid = False
    return bValid


def transimg(img_path):
    """
    转换图片格式
    :param img_path:图片路径
    :return: True：成功 False：失败
    """
    if IsValidImage(img_path):
        try:
            str = img_path.rsplit(".", 1)
            output_img_path = img_path[img_path.rfind("\\")+2:].replace("/","_") 
            print(output_img_path)
            im = Image.open(img_path)
            im = im.convert("RGB")
            #ImageDraw.Draw(im)
            im.save(output_img_path)
            return output_img_path
        except Exception as e:
            print('Failed to trans img: '+ str(e))
    else:
        return None

if __name__ == '__main__':
    if len(sys.argv) == 5:
        print( 'Start comparing...')

        fold_file_path0 = os.getcwd() + "\\" + sys.argv[2]
        fold_file_path1 = os.getcwd() + "\\" + sys.argv[3]
        
        dir_list0 = os.listdir(fold_file_path0)
        dir_list1 = os.listdir(fold_file_path1)

        if not (os.path.exists(os.getcwd() + "/" + sys.argv[4])):
            os.mkdir(os.getcwd() + "/" + sys.argv[4])


        for i in range (0, len(dir_list0)):
            new_jpg_file0 =  transimg(fold_file_path0 + "/" + dir_list0[i])
            if new_jpg_file0 is not None:
                if IsValidImage(fold_file_path1 + "/" + dir_list0[i]):
                    new_jpg_file1 =  transimg(fold_file_path1 + "/" + dir_list0[i])
                    if new_jpg_file1 is not None:
                        compare_images(new_jpg_file0, new_jpg_file1, sys.argv[4], '')
                    else:
                        print("Fail to get corresponding file of " + str(new_jpg_file0))
                else:
                    print("can't find "+str(dir_list0[i]) + " in " + str(fold_file_path1))
            else:
                print("Fail to get jpg file of " + str(dir_list0[i]))

    else:
        print('Please input command such as inputfold1 inputfold2 outputfold:')
        print('python imageCompare.py fold1 fold2 outputfold')