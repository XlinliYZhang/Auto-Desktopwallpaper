import os
import time
import random
import ctypes
import win32api
import win32gui
import win32con
import requests
import threading
import urllib.request
from tkinter import *

Label2 = None
Label4 = None
button3value = None
is_Sync = False
is_Exit = False
is_Lock = False
Text_Path = "Log.txt"
DLL_Path = "./SetWallpaper.dll"
HTTP_URL = "http://bing.plmeizi.com/view/"
Image_Path = "\\image\\"
Image_Name_Head = "<meta name=author content=\""
Image_Name_Tail = "\"><meta name=viewport content=\"width=device-width, initial-scale=1\">"
Image_List_Head = "<ul id=\"images\" class=\"docs-pictures clearfix\">"
Image_List_Tail = "</ul>"
Refresh_Time = 900
Last_Refresh_Time = 0


# 更新屏幕线程
def refresh_thread():
    global Last_Refresh_Time, is_Exit, Label4, is_Lock
    try:
        print("Thread Run")
        while True:
            # 当窗口关闭的时候退出线程
            if is_Exit:
                print("Stop thread")
                sys.exit(0)
            lock_count = 0
            while is_Lock:
                time.sleep(0.1)
                # 防止死锁
                lock_count = lock_count + 1
                if lock_count >= 20:
                    print("Over flow")
                    break
            if lock_count < 20:
                is_Lock = True
                # 定时刷新屏幕
                try:
                    if (Last_Refresh_Time == 0) or (time.time() - Last_Refresh_Time >= Refresh_Time):
                        refresh_wallpaper()
                        # 记录刷新时间
                        Last_Refresh_Time = time.time()
                except:
                    print("")
                # 刷新界面倒计时
                # 如果是0则计算间隔时间
                if Last_Refresh_Time > 0:
                    Last_Time = Refresh_Time - (time.time() - Last_Refresh_Time)
                    Last_Time_str = "%02d:%02d" % (Last_Time / 60, Last_Time % 60)
                else:
                    Last_Time = Refresh_Time
                    Last_Time_str = "%02d:%02d" % (Last_Time / 60, Last_Time % 60)
                Label4["text"] = Last_Time_str
                is_Lock = False
                time.sleep(0.5)
    except:
        if not is_Exit:
            print("Threading Error Stop")
            refresh = threading.Thread(target=refresh_thread)
            refresh.start()
        else:
            print("Threading Normal Stop")


# 设置壁纸函数
# screen_id  屏幕号
# image_path 完整本地壁纸路径
def setWallpaper(screen_id, image_path):
    global is_Sync
    print("Set Screen:", screen_id, end="\tPath:")
    print(image_path)
    # 判断是否共用一张壁纸
    if is_Sync:
        # 用于设置单张壁纸, 无需DLL文件
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, image_path, win32con.SPIF_SENDWININICHANGE)
    else:
        try:
            dll = ctypes.CDLL(DLL_Path)
            dll.SetWallpaper(image_path, screen_id)
        except:
            print("Use dll set wallpaper fail")


# 返回离线壁纸
def offline_image():
    global Image_Path, Text_Path
    image_list = []
    # 读取文件夹目录
    print("Read Image Dir:", Image_Path)
    for x in os.listdir(Image_Path):
        # 判断是不是壁纸
        if x.endswith("jpg"):
            # print("Image File:", x)
            image_list.append(x)
        # else:
        # print("No Image_File:", x)
    # 随机抽取文件
    if len(image_list):
        # print("Image Paht:", image_list)
        image_path = Image_Path + random.choice(image_list)
        print("Offline Image: ", image_path)
        try:
            # 写入更新历史
            file_Handle = open(Text_Path, mode='a+')
            print(time.strftime("%Y-%m-%d %H:%M:%S\t[", time.localtime()), end="", file=file_Handle)
            print(image_path[image_path.find(Image_Path) + len(Image_Path):image_path.find(".jpg")],
                  end="", file=file_Handle)
            print("]", file=file_Handle)
            file_Handle.close()
        except:
            print("Save Log error")
        return image_path
    print("Dir is None")
    return None


# 请求图片函数
# 返回None,则请求失败
# 返回文件目录则请求成功
def request_bing():
    global HTTP_URL, Text_Path, Image_Path, Image_Name_Head, Image_Name_Tail, Image_List_Head, Image_List_Tail
    # 获取随机图片ID
    image_id = random.randint(1, 1846)
    # 请求图片
    try:
        res = requests.get(HTTP_URL + str(image_id), timeout=1.5)
        print("\r\nGET " + HTTP_URL + str(image_id))
    except requests.exceptions.RequestException as e:
        print("Get Fail: ", e)
        return offline_image()
    # 判断请求是否成功
    if res.status_code != 200:
        print("GET Fail, Code: ", res.status_code)
        return offline_image()
    # 判断是否被劫持
    if res.text.find("BING每日壁纸") == -1:
        print("Net work no")
        return offline_image()
    # 打印请求返回值
    # print("Response:", end="")
    # print(res.text)
    # 获取图片名称
    image_name = res.text[(res.text.find(Image_Name_Head) + len(Image_Name_Head)):res.text.find(Image_Name_Tail)]
    # 打印图片名称
    print("image name:", end="")
    print(image_name)
    # 切割全部图片为列表
    image_list = res.text[(res.text.find(Image_List_Head) + len(Image_List_Head)):
                          res.text.find(Image_List_Tail)].split("<li>")
    # 查找列表中图片名字
    for i in image_list:
        if i.find(image_name) != -1:
            # 切割出图片url
            image_url = i[i.find("src=//") + len("src=//"):i.find("-thumb alt")]
            image_url = image_url.replace("bimgs", "bing")
            print("image url: " + image_url)
            # 保存图片
            image_file = Image_Path + str(image_id) + ".jpg"
            try:
                urllib.request.urlretrieve("http://" + image_url, image_file)
                print("image path:", image_file)
            except:
                # 储存失败则使用本地壁纸
                print("Save image fail")
                return offline_image()
            try:
                # 写入更新日志
                file_Handle = open(Text_Path, mode='a+')
                print(time.strftime("%Y-%m-%d %H:%M:%S\t", time.localtime()), end="", file=file_Handle)
                print("[", end="", file=file_Handle)
                print(image_id, end="", file=file_Handle)
                print("]", end="\t", file=file_Handle)
                file_Handle.close()
                file_Handle = open(Text_Path, mode='ab+')
                file_Handle.write(image_name.encode("GBK", "ignore"))
                file_Handle.write("\r\n".encode("GBK"))
                file_Handle.close()
            except:
                print("Save Log error")
            return image_file
            # 返回文件名称
    return offline_image()


# 请求并更换壁纸
def refresh_wallpaper():
    global Last_Refresh_Time, is_Sync
    # 判断屏幕数量
    MonitorNumber = win32api.GetSystemMetrics(win32con.SM_CMONITORS)
    print("Monitor: ", MonitorNumber)
    if is_Sync:
        screen_image = None
        while screen_image == None:
            screen_image = request_bing()
        setWallpaper(0, screen_image)
    else:
        # 为每一个屏幕请求壁纸
        for i in range(MonitorNumber):
            screen_image = None
            while screen_image == None:
                screen_image = request_bing()
            # 设置壁纸
            setWallpaper(i, screen_image)
    print("Refresh complete\r\n")


# 检查需要的资源文件在不在
def check_resources():
    global Image_Path, Text_Path
    # 获取图片路径
    Image_Path = os.getcwd() + Image_Path
    # 判断文件夹是否存在
    Path_Valid = os.path.exists(Image_Path.rstrip("\\"))
    if Path_Valid:
        print("Path is valid")
    else:
        print("Path is not valid")
        os.makedirs(Image_Path.rstrip("\\"))
    print("Save Image Path:", Image_Path)
    # 用于储存图片介绍
    Text_Path = Image_Path + Text_Path
    print("Save Log Path:", Text_Path)


# 刷新壁纸按钮回调函数
def button1_callback():
    global Last_Refresh_Time, is_Lock
    print("Button1 Callback\r\n")
    Lock_Time = 0
    while is_Lock:
        time.sleep(0.1)
        Lock_Time = Lock_Time + 1
        if Lock_Time >= 20:
            print("Button1 Overflow")
            return None
    is_Lock = True
    Last_Refresh_Time = 0
    is_Lock = False


# 打开目录按钮回调函数
def button2_callback():
    print("Button2 Callback")
    os.system("explorer.exe %s" % Image_Path)


# 复选框回调函数
def button3_callback():
    global button3value, is_Sync
    print("Button3 Callback:", button3value.get())
    is_Sync = button3value.get()


# 时间滑轨回调函数
def slider1_callback(text):
    global Label2, Refresh_Time
    # print("Slider1: ",text)
    try:
        Refresh_Time = float(text)
        Label2["text"] = "刷新间隔:" + text + " S"
    except:
        print("Set time fail")


# 绘制UI
def draw_windows():
    global Refresh_Time, is_Exit, Label2, Label4, button3value
    windows = Tk()
    # 窗口外观相关
    windows.geometry("400x200")
    windows.title("Bing壁纸刷新程序")
    # 阻止调整串口大小
    windows.resizable(0, 0)
    # 滑轨
    slider1value = IntVar()
    button3value = IntVar()
    Slider1 = Scale(windows, from_=30, to=900, orient="horizontal", length=390, width=15,
                    variable=slider1value,
                    command=slider1_callback,
                    tickinterval=200,
                    showvalue=False,
                    font=("宋体", 20))
    slider1value.set(Refresh_Time)
    button3value.set(False)
    Slider1.place(x=5, y=95)
    # 显示监视器数量
    MonitorNumber = win32api.GetSystemMetrics(win32con.SM_CMONITORS)
    Label1 = Label(windows, anchor='n', text="本机屏幕数量: " + str(MonitorNumber), font=("宋体", 10))
    Label1.place(x=5, y=0)
    Label2 = Label(windows, anchor='n', text="刷新间隔:900 S", font=("宋体", 20))
    Label2.place(x=5, y=50)
    Label3 = Label(windows, anchor='n', text="距下次刷新", font=("宋体", 20))
    Label3.place(x=250, y=0)
    Label4 = Label(windows, anchor='n', text="15:00", font=("宋体", 20))
    Label4.place(x=280, y=40)
    # 刷新按钮
    Button1 = Button(windows, command=button1_callback, text="立即更新壁纸", font=("宋体", 20), relief="raised", bd=5)
    Button1.place(x=5, y=145)
    Button2 = Button(windows, command=button2_callback, text="打开壁纸目录", font=("宋体", 20), relief="raised", bd=5)
    Button2.place(x=205, y=145)

    Button3 = Checkbutton(windows, command=button3_callback,
                          onvalue=True, offvalue=False,
                          variable=button3value, text="同步更新", font=("宋体", 10))
    Button3.place(x=5, y=20)
    # 设置窗口图标
    # windows.iconphoto(True, PhotoImage(file="plane.png"))
    # 启动刷新线程
    refresh = threading.Thread(target=refresh_thread)
    refresh.start()
    # 窗口运行
    windows.mainloop()


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    # 检查资源
    check_resources()
    # 绘制窗口
    draw_windows()
    # 通知线程退出
    is_Exit = True
