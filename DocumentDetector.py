# -*- coding: GBK -*-

'''
Document Detector
文档探测器

通途：
    自动拷课件，课件一被打开立即拷走保存*。按日期归档，方便班级电脑管理员整理文件。
    --------------------
    *注：复制文件时本程序并不会询问老师，请在使用本程序时请尊重老师的著作权；

特点：
    1. 自动扫描并复制已打开的课件；
    2. 只复制含特定的扩展名的文件*；
    3. 复制成功托盘图标气泡提醒；
    4. 复制后自动按日期分类，免去手动分类的烦恼；
    5. 文档关闭时再次复制一遍，以保存更改。
    --------------------
    *注：ppt、pptx、doc、docx、xls、xlsx、pdf。

FileBox目录结构（示例）：
    ./FileBox
        |-- 2018.06.28
        |       `-- Calculus.pptx
        |-- 2018.06.29
        |       |-- Irregular Verbs.docx
        |       `-- Recommended Books.pdf
        |-- 2018.07.01
        |       `-- Statistics.xls
        ...

兼容性（已测试）：
    Windows XP/7/8/10 32bit/64bit
    
Python 版本：
    Python 2.7.15
    
使用到的第三方库：
    1. psutil 3.4.2
    2. pywin32 223
    3. py2exe 0.6.9（打包成exe时使用）

已知缺陷：
    1. CPU占用略高，调整SLEEP_TIME可稍缓解；
    2. 当explorer.exe崩溃时，重启explorer后托盘图标消失；
    3. 由于某种诡异的兼容性问题，该程序只能在Windows XP环境下使用py2exe编译打包，或将"bundle_files"设为3。
'''

# 一些设置
SLEEP_TIME = 1; # 每次扫描的时间间隔（整数，秒）
EXTENSIONS = ['ppt', 'pptx', 'doc', 'docx', 'xls', 'xlsx', 'pdf']; # 需要收集的文件扩展名
NOTIFICATION = True; # 是否开启气泡通知

import re;
import os;
import sys;
import time;
import thread;
import psutil;
import shutil;
import datetime;
import win32api;
import win32gui;
import win32con;
import win32tray;
import CreateIcon;

# 如果直接运行python脚本，则进入调试模式
if re.match('.*\.exe$', sys.argv[0]) == None:
    DEBUG = True;
else:
    DEBUG = False;

if not DEBUG:
    # 检查Document Detector是否已经运行
    run = 0;
    for p in psutil.pids():
        try:
            process = psutil.Process(p);
            if process.cmdline()[0] == sys.argv[0]:
                run += 1;
            if run > 1:
                sys.exit(); # 如果已经运行，直接退出。
        except IndexError:
            pass;
        except psutil.Error: # 忽略进程拒绝访问等异常
            pass;


# ==== 准备部分 ====

# 〔全局变量部分〕
# tray_exit，托盘退出时为True。
# copying，正在复制时为True。
# hWnd，为托盘图标窗口句柄。

tray_exit = False;
copying = False;
hWnd = 0;

# 创建FileBox文件夹
mydir = os.path.dirname(sys.argv[0]); # 当前文件夹
boxdir = os.path.join(mydir, 'FileBox');
if not os.path.exists(boxdir):
    os.mkdir(boxdir);

# 获取日期文件夹路径（文件夹不存在，则在FileBox下创建）
def get_cdir(boxdir):
    now = datetime.datetime.now();
    # 格式如2017.01.31
    date_str = datetime.datetime.strftime(now, '%Y.%m.%d');
    cdir = os.path.join(boxdir, date_str);
    if not os.path.exists(cdir):
        os.mkdir(cdir);
    return cdir;


# ==== 菜单部分 ====

# 关于
def about(obj):
    win32api.MessageBox( 0,
                         '文档探测器',
                         '关于',
                         win32con.MB_ICONASTERISK );

# 打开目标路径
def open_box(obj):
    os.startfile(boxdir);

# 添加到开机启动项
def startup(obj):
    if win32api.MessageBox( 0,
                            '添加到开机启动项？',
                            '文档探测器',
                            win32con.MB_YESNO ) == 6:
        # 注册表启动项所在路径
        key_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Run';
        # 打开该键（使用HKCU，仅对当前用户有效）
        key = win32api.RegOpenKey( win32con.HKEY_CURRENT_USER,
                                   key_path,
                                   0,
                                   win32con.KEY_WRITE          );
        # 获取自身所在路径
        mypath = sys.argv[0];
        # 写入注册表
        win32api.RegSetValueEx( key,
                                'Document_Detector',
                                0,
                                win32con.REG_SZ,
                                mypath               );
        # 关闭
        win32api.RegCloseKey(key);


# ==== 托盘部分 ====

# 显示托盘图标
# (使用全局变量: hWnd, tray_exit)
# (托盘退出时: tray_exit=True)
def show_tray():
    # 在临时目录创建托盘图标ico
    icon_path = os.path.join(os.environ['TEMP'],'dd_tray.ico');
    CreateIcon.create(icon_path);
    # 悬停文字和菜单
    hover_text = '文档探测器';
    menu_options = ( ('关于', None, about),
                     ('打开目标路径', None, open_box),
                     ('添加到开机启动项', None, startup) );
    # 退出时函数
    def bye(STI):
        global tray_exit;
        tray_exit = True;
        if DEBUG: print 'tray_exit';
    # 创建托盘图标
    STI = win32tray.SysTrayIcon( icon_path,
                                 hover_text,
                                 menu_options,
                                 on_quit = bye );
    # 获得窗口句柄保存在全局变量hWnd中
    global hWnd;
    hWnd = STI.hwnd;
    # 开始消息循环
    win32gui.PumpMessages();


# ==== 复制部分 ====

# 复制文件
# (使用全局变量: copying)
# (正在复制时: copying=True)
def copy_file(myfile, boxdir):
    global copying;
    copying = True;
    cdir = get_cdir(boxdir); # 获取日期路径
    try:
        # 使用shutil.copy2可将文件带附加信息一同复制
        shutil.copy2(myfile, cdir);
        if NOTIFICATION and not tray_exit:
            # 气泡显示
            win32tray.show_balloon( hWnd,
                                    'Document Detector',
                                    os.path.basename(myfile).decode('gbk') );
    except Exception as err:
        if DEBUG: print err;
    copying = False;


# ==== 核心部分 ====

def main():
    # 子线程处理托盘
    thread.start_new_thread(show_tray,());
    
    time.sleep(1); # 等待托盘准备完毕
    if NOTIFICATION and not tray_exit:
        win32tray.show_balloon( hWnd,
                                'Document Detector',
                                'Program started.' );
    
    
    # 主循环
    # 通过比较files_this_time（本次扫描到的文档文件）与files_last_time（上次扫描到的文档文件），了解到新打开了某个文档，把它拷走。
    
    files_last_time = set(); # 上次扫描到的文档文件
    
    while not tray_exit:
        files_all = set(); # 当前打开的所有文件
        files_this_time = set(); # 当前打开的文档文件（匹配后的）
        for p in psutil.pids():
            try:
                process = psutil.Process(p);
                # 通过两种方法找出打开的文件
                # 1. 命令行
                c = process.cmdline()[1:];
                for cmd in c:
                    if os.path.exists(cmd): # 是个文件
                        files_all.add(cmd);
                # 2. 进程打开的文件
                l = process.open_files();
                for n in range(len(l)):
                    files_all.add(l[n].path);
            except psutil.Error: # 忽略进程拒绝访问等异常
                pass;
        # 正则匹配
        for f in files_all:
            if re.match(r'.*\.(%s)$' % "|".join(EXTENSIONS), f) != None:
                files_this_time.add(f);
        
        files_to_copy = (files_this_time - files_last_time) | (files_last_time - files_this_time);
        
        # 输出一些调试信息
        if DEBUG: print 'files_this_time: %s' % str(files_this_time);
        if DEBUG: print 'files_last_time: %s' % str(files_last_time); 
        if DEBUG: print 'files_to_copy:   %s' % str(files_to_copy); 
        if DEBUG: print;
        
        for document_detected in files_to_copy:
            if DEBUG: print '<<<%s>>>' % document_detected;
            thread.start_new_thread(copy_file,(document_detected, boxdir));
        
        files_last_time = files_this_time;
        
        # 等待下一次扫描，托盘退出就停
        for i in range(SLEEP_TIME):
            if tray_exit: return;
            time.sleep(1);

if __name__ == "__main__":
    main();
    # 退出前等待子线程拷贝完毕
    while copying:
        time.sleep(1);
    # 安全地退出
