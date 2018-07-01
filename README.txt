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
