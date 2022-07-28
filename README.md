# 使用者
Word格式说明

全文字体类型：等线（正文）
全文字体大小：五号
全文行间距：1.5倍

Excel格式说明

第一行 由 security_company	fund_company  reseacher 三列构成
下面写对应的信息


# 开发者：

	建议开发时安装Anaconda（Python环境），利用Anaconda创建Python虚拟环境
					
	安装依赖，所需依赖在requirements.txt中，需要打开cmd窗口，并将工作路径切换到C:\Users\xxx\Documents\CustomWatermark文件夹下，输入    pip install -r requirements.txt    等待安装完毕即可

	主程序入口：src\main_gui.py

# 打包成exe程序


```

$ pyinstaller -D -w src/main_gui.py -p src/main_window.py -p src/addmask_main.py --hidden-import main_window --hidden-import addmask_main

```