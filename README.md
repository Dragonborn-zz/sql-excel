# SQL 查 Excel
将 Excel 的各个工作表导入本地 SQLite，在桌面窗口中直接编写 SQL 查询，结果可导出为 CSV 或 Excel。原文件不会被修改或删除。

这款软件主要面向平时不常用 Excel、但熟悉 SQL 的人群。比如，你突然需要处理几个 Excel 文件，却不知道该用什么函数；好不容易找到函数，又被一堆参数搞得不知道怎么调用——没关系，直接把文件拖进来，用 SQL 像查表一样查数据。需要对比几个 Excel、找出差异数据？把多个 Excel 一起拖进来，多表联查一下，结果就出来了。处理 Excel，就这么简单。

## 环境

- Windows 10 / 11，已安装 [WebView2](https://developer.microsoft.com/microsoft-edge/webview2/)（较新的系统一般自带）
- Python 3.11 及以上

## 启动

双击 `start.bat`。第一次运行会创建 `.venv` 并安装依赖，然后打开窗口。

也可以手动启动：

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python main.py
```

国内安装依赖若超时，可改用阿里云镜像：

```powershell
python -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
```

## 使用

1. 把 `.xlsx`、`.xlsm` 或 `.xls` 拖进左侧区域，或点击该区域选择文件。每个工作表变成一张表。
2. 表名是 `文件名__工作表`。同名时自动变成 `文件名__工作表_2`、`_3`。
3. 第一行非空行当作列名。空列名会变成 `列1`、`列2`。重复列名里第一个保留原名，后面的用下一个空着的 `_2`、`_3`。
4. 在右侧输入 SQL，点「运行」，或按 `Ctrl + Enter`。
5. 单击左侧表名或列名，会把带引号的名字插入当前页签的光标处。单击表名以外的表头空白处，会打开该表自己的查询页签，不改动已写的 SQL。
6. 查询出结果后，可「导出 CSV」或「导出 Excel」。
7. 「清空数据」会删掉已导入的表。原始 Excel 仍在原位置。

导入后的数据在项目目录的 `data\workspace.sqlite`。这个目录只在本机使用，不会提交到仓库。

查询结果最多显示 2000 行；超过时界面会标明已截断。不支持 `ATTACH`、`DETACH`、`load_extension` 和 `VACUUM INTO`。

##软件截图

<img width="2559" height="1334" alt="image" src="https://github.com/user-attachments/assets/f5d65a01-5225-4610-8694-12c3885571ef" />

