import pyautogui
import time
import win32gui
import win32con
from cnocr import CnOcr
from PIL import ImageGrab
import random
pyautogui.PAUSE=0.1
pyautogui.FAILSAFE =False

"""
读取窗口坐标---读取窗口位置时候经常会用
import  fuc
hwnd=fuc.get_app_hwnd('MuMU模拟器12')
fuc.maximize_window(hwnd)
fuc.move_window_to_top_left(hwnd)
time.sleep(0.5)
pyautogui.mouseInfo() --会弹出对话框，实时显示鼠标坐标

"""
def maximize_window(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

def restore_window(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

def minimize_window(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

#移动窗口到左上角
def move_window_to_top_left(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)

#移动窗口到指定位置
def move_window(hwnd,t):
    win32gui.MoveWindow(hwnd, t[0], t[1], t[2], t[3], True)

#让窗口最上
def set_window_topmost(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
#取消窗口最上
def set_window_notopmost(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

"""
其中，rect1一个四个元素的元组， 是一个窗口的左上角坐标和右下角坐标，
point是鼠标在屏幕上的坐标，rect2为这个窗口变化后左上角坐标和右下角坐标。
本函数首先得到下point在之前窗口内的相对坐标，然后返回窗口变化后相对坐标不变的情况下，相对于屏幕的坐标

"""
def find_point(rect1, point, rect2):
    # 计算鼠标在rect1定义的矩形区域内的相对坐标
    relative_x = (point[0] - rect1[0]) / (rect1[2] - rect1[0])
    relative_y = (point[1] - rect1[1]) / (rect1[3] - rect1[1])

    # 计算rect2移动的距离和大小变化的比例
    dx = rect2[0] - rect1[0]
    dy = rect2[1] - rect1[1]
    scale_x = (rect2[2] - rect2[0]) / (rect1[2] - rect1[0])
    scale_y = (rect2[3] - rect2[1]) / (rect1[3] - rect1[1])

    # 计算鼠标在rect2定义的矩形区域内的绝对坐标
    absolute_x = int(rect2[0] + relative_x * (rect2[2] - rect2[0]))
    absolute_y = int(rect2[1] + relative_y * (rect2[3] - rect2[1]))

    return absolute_x, absolute_y

#获取应用程序窗口句柄
def get_app_hwnd(name):
    hwnd = win32gui.FindWindow(None,name)  # 替换成你需要点击的窗口标题
    return hwnd

#获取应用程序窗口坐标
def get_app_point(name):
    hwnd=get_app_hwnd(name)
    # 获取窗口的位置（左上角坐标）
    window_rect = win32gui.GetWindowRect(hwnd)
    window_x, window_y, _, _ = window_rect
    return window_x, window_y

def get_app_rect(hwnd):
    window_rect = win32gui.GetWindowRect(hwnd)
    return window_rect


#得到某个点在应用程序窗口的相对坐标
def get_app_relative_point(name, target_x, target_y):
    window_x,window_y=get_app_point(name)
    # 计算目标点的相对坐标
    relative_x = target_x - window_x
    relative_y = target_y - window_y
    return relative_x, relative_y

#点击窗口的某一点
def click_window_point( x, y):
    # 将鼠标移动到指定点
    pyautogui.moveTo(x, y)
    # 模拟鼠标点击
    pyautogui.click()

#通过名称找到app的相对位置,点击app应用程序中的某一点
def click_app_point_fromname(name, relative_x, relative_y):
    x,y=get_app_point(name)
    click_window_point(relative_x+x,relative_y+y)

#通过hwnd,点击app应用程序中的某一点
# def click_app_point_fromhwnd(hwnd, relative_x, relative_y):
#     lParam = win32api.MAKELONG(int(relative_x), int(relative_y))
#     win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
#     time.sleep(0.05)
#     win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, lParam)

#得到某一点的颜色rgb
def get_pixel_color( x, y):
   return  pyautogui.pixel(x,y)



#根据坐标中心点和，x轴 y轴可便宜量，随机产生一个鼠标要点击的位置
def set_random_pos(centerpos,offset):
    if offset[0]==0:
        return centerpos
    random_number = random.randint(-offset[0], offset[0])
    x=centerpos[0]+random_number
    random_number = random.randint(-offset[1], offset[1])
    y = centerpos[1] + random_number
    return x,y

def is_have_img(name,r,confidence_val=0.7,grayscale=True):
    r0=r[0]
    r1=r[1]

    if r0<0:
        r0=0
    if r1<0:
        r1=0
    try:
        coords=pyautogui.locateOnScreen(name,region=(r0,r1,r[2],r[3]),confidence=confidence_val,grayscale=grayscale)
        if coords is not None:
            return True
        else:
            return False
    except Exception as e:
        return False

#找到当前屏幕上的与给定的图片匹配的位置并点击
def click_pic(name,confidence_val=0.7,grayscale=True,offset=(3,3)):

    pyautogui.PAUSE=0.1
    pyautogui.FAILSAFE =True
    try:
        coords=pyautogui.locateOnScreen(name,confidence=confidence_val,grayscale=grayscale)
        if coords is not None:
            x,y=pyautogui.center(coords);
            x, y = set_random_pos((x, y), offset)
            pyautogui.leftClick(x,y)
        else:
            print("没找到: ",name)
    except Exception as e:
        print("没找到: ",name)


#找到当前区域（参数r）上的与给定的图片匹配的位置并按要求多次点击，
# 如果一次寻找不到，会再次寻找
#name 图片路径
#click_pos=1  点击中心    click_pos=2 点击左上角
#click_pos=3 点击右下角
#click_pos=4 点击右上角
#click_num 鼠标点击次数，默认为1
#offset 模拟人工操作，鼠标在x轴和y轴偏移量
#confidence_val 辨别图片的精确度
#grayscale 灰度寻找：把图片和屏幕视为灰色


def click_pic_in_rect(name,r,click_pos=1,click_num=1,offset=(3,3),confidence_val=0.7,grayscale=True):
    # 调整rect范围在正整数
    r0 = r[0]
    if r[0] < -1:
        r0 = 0
    r1 = r[1]
    if r[1] < -1:
        r1 = 0
    temp_r = (r0, r1, r[2], r[3])
    try:
        coords = pyautogui.locateOnScreen(name, region=temp_r, confidence=confidence_val, grayscale=grayscale)

        if coords is not None:
            if click_pos == 1:
                x, y = pyautogui.center(coords);
                x, y = set_random_pos((x, y), offset)
            if click_pos == 2:
                x = coords.left + coords.width
                y = coords.top + coords.height
                x, y = set_random_pos((x, y), offset)
            if click_pos == 3:
                x = coords.left
                y = coords.top
                x, y = set_random_pos((x, y), offset)
            if click_pos == 4:
                x = coords.left+coords.width
                y = coords.top
                x, y = set_random_pos((x, y), offset)
            #print("***", x, y)
            for i in range(0, click_num):
                pyautogui.leftClick(x, y,0.1)

            return True
        else:
            return False
    except Exception as e:
        return False


""" 
本方法根据参数plan指定的工作流程，
去寻找游戏画面的相应的图片或者坐标进行点击

win_rect：要寻找图片的区域
floder：图片文件存放路径
plan：
    imgs 要寻找的图片名称，可能某次寻找会因为存在多种情况
    比如要点击背包这个图像，这个图像空的时候是灰色，满的时候是红色
    会依次寻找多个，寻找到某个立即结束，
    coordinates  坐标,如果图片均为寻找到，就会看这个属性是否设置，如果设置，
    则点击此坐标
    wait_time 多久后开始操作
    ignore  如果失败是否忽略
    retry_num  这个参数是寻找次数
    delay_time retry时间间隔
        
    注意：wait_time是这张图片要多少秒后点击，delay_time是如果寻找失败后且要多次寻找的话之间的时间间隔     
    注意：此后的键值都是需要传递到find_click_pic_in_rect函数的 
    
    click_pos=1  点击中心  click_pos=2 点击右下角 
    click_pos=3 点击左上角
    click_pos=4 点击右上角
    click_num 鼠标点击次数，默认为1
    offset 模拟人工操作，鼠标在x轴和y轴偏移量
    confidence_val 辨别图片的精确度
    grayscale 灰度寻找：把图片和屏幕视为灰色
    
    格式：
    {
      'imgs': ('close1.jpg',),'coordinates':(100,355), 'wait_time': 1, 'ignore': 1,'retry_num' : 3, 'delay_time' : 2,'click_pos': 1, 'click_num': 1, 
      'offset' : (2, 2), 'confidence_val' : 0.7, 'grayscale' : True
    }
find_click_pic_in_rect 默认值：
(name,r,click_pos=1,click_num=1,offset=(3,3),confidence_val=0.7,grayscale=True):
"""
def do_process(win_rect, floder, plan):

    for i in range(0, len(plan)):
        sing_p = plan[i]
        wait_time=1
        if 'wait_time' in sing_p:
            wait_time=sing_p['wait_time']
        time.sleep(wait_time)
        is_suc = False
        path=''
        #每个步骤会包含多张图片
        for n in range(0, len(sing_p['imgs'])):

            # 拼装图片路径
            path = floder + sing_p['imgs'][n]

            # 取出所有的要传递的参数
            para_dict = {key: value for key, value in sing_p.items() if
                         key not in ('imgs', 'coordinates', 'wait_time', 'ignore','retry_num','delay_time')}

            is_suc=False
            retry_num = 3
            if 'retry_num' in sing_p:
                retry_num = sing_p['retry_num']

            while not is_suc and retry_num>0:
                is_suc = click_pic_in_rect(path, win_rect, **para_dict)
                if not is_suc:
                    #print("--------",path)
                    delay_time = 1
                    if 'delay_time' in sing_p:
                        delay_time = sing_p['delay_time']
                    time.sleep(delay_time)
                    retry_num=retry_num-1

            if is_suc:
                break
        if not is_suc and 'coordinates' in sing_p:
            # 点击图片失败后可以选择点击坐标
            x,y=2,2
            if 'offset' in sing_p:
                x,y=sing_p['offset']
            #print("---",sing_p['coordinates'])
            x, y = set_random_pos(sing_p['coordinates'], (x,y))
            #print("---",x,y)
            pyautogui.leftClick(x,y)
            print("使用坐标")
            is_suc=True



        if  not is_suc and not sing_p['ignore']:
            return False,path
    return True,'0'


#截取屏幕上一部分区域
def capture_screen(region):
    # 使用 ImageGrab 模块截取屏幕上指定区域的图片
    image = ImageGrab.grab(region)
    return image


#截取屏幕上一部分区域并保存在指定路径
def capture_screen(region,file_path):
    # 使用 ImageGrab 模块截取屏幕上指定区域的图片
    image = ImageGrab.grab(region)
    image.save(file_path)

#假定图片上只有一行，读取并返回json数据
def read_single_pic(img_fp):
    #ocr = CnOcr()
    #doc-densenet_lite_246-gru_base
    #ocr = CnOcr(rec_model_name='densenet_lite_136-gru', det_model_name='naive_det', cand_alphabet=r'(-;；+=-*,.、;|\')0123456789')
    ocr = CnOcr(rec_model_name='en_PP-OCRv3', det_model_name='naive_det')




    #ocr = CnOcr(cand_alphabet='(0123456789--;；\,\.、;|\')')
    out = ocr.ocr_for_single_line(img_fp)
    #out = ocr.ocr(img_fp)
    #print(out)
    return out

#读取图片上的内容，返回json数据
def read_pic(img_fp):
    ocr = CnOcr(cand_alphabet=r'(0123456789--;；+=-*,.、;|\)')
    out = ocr.ocr(img_fp)
    #print(out)
    return out






