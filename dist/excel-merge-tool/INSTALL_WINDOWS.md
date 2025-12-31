# Windows 安装说明

## 系统要求
- Windows 7 或更高版本
- Python 3.7 或更高版本

## 安装步骤

### 1. 安装 Python
如果您的电脑上还没有安装 Python，请按照以下步骤安装：

1. 访问 [Python 官网](https://www.python.org/downloads/windows/) 下载最新版本的 Python 安装包
2. 运行安装程序，勾选 "Add Python to PATH" 选项
3. 点击 "Install Now" 完成安装
4. 安装完成后，打开命令提示符（CMD）并输入 `python --version` 验证安装是否成功

### 2. 下载并解压工具

1. 下载 `excel-merge-tool.zip` 文件
2. 右键点击文件，选择 "提取全部"，将文件解压到您想要的位置

### 3. 安装依赖

1. 打开命令提示符（CMD）
2. 使用 `cd` 命令导航到解压后的文件夹，例如：
   ```
   cd C:\Users\YourName\Downloads\excel-merge-tool
   ```
3. 运行以下命令安装所需依赖：
   ```
   pip install -r requirements.txt
   ```

## 使用方法

### 方法 1：双击运行批处理文件

1. 在解压后的文件夹中找到 `run_excel_merge.bat` 文件
2. 双击运行该文件
3. 按照提示选择文件进行处理

### 方法 2：命令行运行

1. 打开命令提示符（CMD）
2. 导航到解压后的文件夹
3. 运行以下命令启动交互式模式：
   ```
   python excel_merge.py
   ```
4. 按照提示选择文件进行处理

### 方法 3：命令行参数运行

```
python cli.py <order_file> <payment_file> [-o <output_file>]
```

示例：
```
python cli.py ExcelForHandel/order.xlsx ExcelForHandel/payment.xlsx
```

## 注意事项

1. 确保您的 Excel 文件位于 `ExcelForHandel` 目录下，或者在运行命令时提供完整路径
2. 支持的文件格式：.xlsx、.xls、.csv
3. 程序会直接修改原始订单文件，建议在运行前备份重要数据
4. 如果遇到编码问题，尝试将 CSV 文件保存为 UTF-8 编码

## 故障排除

### 常见错误

1. **Python 未找到**
   - 确保 Python 已安装且已添加到 PATH
   - 尝试使用 `python3` 代替 `python` 命令

2. **依赖安装失败**
   - 确保您的网络连接正常
   - 尝试使用 `pip3` 代替 `pip` 命令
   - 尝试更新 pip：`python -m pip install --upgrade pip`

3. **文件未找到**
   - 确保文件路径正确
   - 确保文件位于 `ExcelForHandel` 目录下

### 联系支持

如果您遇到其他问题，请查看项目的 GitHub 仓库：
https://github.com/cnLeoWux/excel-merge