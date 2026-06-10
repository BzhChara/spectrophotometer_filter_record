# 分光光度滤光片记录

Windows 图形界面工具，用于通过串口读取分光光度模块返回的 24 路强度数据，完成空气记录或 10% / 20% / 30% 滤光片稳定值记录，并生成 Excel 文件。

## 功能概览

- PySide6 图形界面，一页显示 24 个通道。
- 支持波长：410 / 460 / 520 / 550 / 590 / 630 nm。
- 支持通道组：CH1-CH24、CH1-CH12、CH13-CH24。
- 支持空气记录模式和滤光片测试模式。
- 滤光片模式会先建立空气基底，再按 `当前值 / 空气基底` 判断滤光片比例。
- 原始数据先写入 CSV 缓存，停止后生成原始数据 Excel。
- 滤光片稳定值保存为单独 Excel，并统计最小值、最大值、平均数、CV。
- 右侧显示运行日志，顶部显示写入行数和稳定通道数量。
- 通道状态点：灰色=未建立基底，绿色=可测试，黄色=检测到滤光片并等待稳定，蓝色=已记录稳定值并等待回到空气。

## 项目结构

```text
app.py                         图形界面入口
app_paths.py                   源码/EXE 运行目录识别
config.ini                     默认运行配置
config_loader.py               配置读取、默认值和类型转换
excel_writer.py                输出目录、CSV 缓存、Excel 生成
filter_logic.py                波长、通道组、稳定判定常量
serial_protocol.py             串口参数、命令构造、返回帧解析
spectro_ui/                    PySide6 界面和后台串口线程
assets/                        图标资源
old/read_filter_stable.py      旧命令行脚本归档，不作为当前入口
分光光度滤光片记录.spec        PyInstaller 打包配置
打包指令.bat                   打包脚本
串口连接.txt                   串口线序备注
```

## 环境依赖

- Windows
- Python 3.11 或兼容版本
- 串口设备可被 Windows 识别为 `COMx`

安装依赖：

```powershell
cd "D:\Visual Studio Code Projects\Python\分光光度模块"
pip install -r requirements.txt
```

依赖包：

```text
pyserial
openpyxl
PySide6>=6.7
```

## 启动程序

```powershell
cd "D:\Visual Studio Code Projects\Python\分光光度模块"
python app.py
```

当前入口只使用图形界面。`old/read_filter_stable.py` 是旧命令行脚本归档。

## 配置文件

程序读取运行目录下的 `config.ini`：

- 源码运行时：`D:\Visual Studio Code Projects\Python\分光光度模块\config.ini`
- 打包 EXE 运行时：EXE 所在目录下的 `config.ini`

当前默认配置：

```ini
[settings]
port = COM6
wavelength = 520
channel_group = 2
filter_ratio = 30
output = 串口数据记录.xlsx
stable_output = 滤光片稳定值.xlsx
ratio_tolerance = 5
air_tolerance = 5
no_start = false
keep_light = false
```

配置项说明：

```text
port              默认串口号
wavelength        波长，可选 410/460/520/550/590/630
channel_group     通道组，0=CH1-CH24，1=CH1-CH12，2=CH13-CH24
filter_ratio      0/10/20/30；0=空气记录模式
output            原始数据 Excel 文件名
stable_output     稳定值 Excel 文件名
ratio_tolerance   滤光片比例允许误差，单位为百分点
air_tolerance     回到空气基底的允许误差，单位为百分点
no_start          true=不发送开灯命令，只读取已有数据流
keep_light        true=停止时不发送关灯命令
```

注意：`output` 和 `stable_output` 只使用文件名部分。即使写成带目录的路径，程序也会通过 `os.path.basename` 取文件名，并保存到本次记录目录中。

## 图形界面流程

1. 选择串口、波长、滤光片比例和通道组。
2. 点击 `开始测试`。
3. 如果是滤光片模式，先保持空气状态，等待目标通道基底建立完成。
4. 通道状态变为绿色后插入滤光片。
5. 命中比例后通道变为黄色，连续稳定 10 帧后写入稳定值，通道变为蓝色。
6. 拔出滤光片，回到空气并稳定后，通道重新变为绿色，可再次测试。
7. 点击 `停止`，程序会关闭串口、生成 Excel，并在日志中输出保存路径。

空气记录模式下不会建立空气基底，也不会生成稳定值 Excel，只持续记录原始数据。

## 保存路径

程序每次开始测试都会创建新的记录目录。当前代码的真实保存路径为：

```text
程序运行目录\data\YYYY_MM_DD\record_HHMMSS\
```

示例：

```text
D:\Visual Studio Code Projects\Python\分光光度模块\data\2026_06_10\record_180715\
```

源码运行时，`程序运行目录` 是项目根目录。打包后运行时，`程序运行目录` 是 EXE 所在目录。

### 空气记录模式

当 `filter_ratio = 0` 或界面选择 `空气记录` 时，产生：

```text
data\YYYY_MM_DD\record_HHMMSS\_temp_csv\串口数据记录.csv
data\YYYY_MM_DD\record_HHMMSS\串口数据记录.xlsx
```

### 滤光片测试模式

当 `filter_ratio = 10/20/30` 时，产生：

```text
data\YYYY_MM_DD\record_HHMMSS\_temp_csv\串口数据记录.csv
data\YYYY_MM_DD\record_HHMMSS\串口数据记录.xlsx
data\YYYY_MM_DD\record_HHMMSS\滤光片稳定值.xlsx
```

生成规则：

- 每收到一帧有效数据，先追加写入 `_temp_csv\串口数据记录.csv`。
- 点击 `停止` 后，程序把 CSV 转成 `串口数据记录.xlsx`。
- CSV 缓存会保留，方便排查原始记录。
- 滤光片模式下，只有至少一个通道记录到稳定值时才保留 `滤光片稳定值.xlsx`。
- 如果没有产生有效数据，程序会删除空 Excel、空 CSV 和空记录目录。
- 如果 Excel/WPS 正在打开输出文件，保存可能失败，日志会提示关闭文件或避免资源管理器预览占用。

## 原始数据格式

原始 CSV 和原始 Excel 格式：

```text
时间 | CH1 | CH2 | ... | CH24
```

如果选择 `CH1-CH12` 或 `CH13-CH24`，只保存当前通道组的列，不保存未选中的通道列。

时间列格式为当前时分秒：

```text
15:25:03
```

数值保留 6 位小数，Excel 字体为 `微软雅黑`。

## 稳定值格式

稳定值 Excel 保存每个通道的多次滤光片稳定值：

```text
通道    第1次    第2次    第3次    ...    最小值    最大值    平均数    CV
CH1     数值     数值     数值            数值      数值      数值      数值
CH2     数值     数值                     数值      数值      数值      数值
...
```

底部还会按每一次测试纵向统计当前通道组的最小值、最大值、平均数、CV。

同一个通道多次插入滤光片时，新稳定值会横向追加，不会覆盖之前的稳定值。如果选择 12 通道组，稳定值和统计值只覆盖该通道组。

## 串口协议

串口参数固定：

```text
115200, 8N1
```

启动采集命令由 `build_cmd(1, 0x50, data)` 生成：

```text
A5 5A 01 50 02 HH LL SUM
```

以 520nm 为例，当前发送：

```text
A5 5A 01 50 02 02 01 55
```

停止命令：

```text
A5 5A 01 50 02 00 00 52
```

波长编号：

```text
410nm -> 0 -> 紫色光
460nm -> 1 -> 蓝色光
520nm -> 2 -> 绿色光
550nm -> 3 -> 绿色光
590nm -> 4 -> 橙色光
630nm -> 5 -> 红色光
```

返回帧解析：

```text
帧头: 5A A5
功能码: 0x56 或 0x57
payload: 至少 96 字节
数据: payload 前 96 字节按小端 24 个 float 解析
SUM: 从帧头开始逐字节累加后取低 8 位
```

软件会丢弃 SUM 校验失败的帧。解析后对每个 float 取绝对值。

## 稳定判定

公共稳定条件：

```text
最近 10 帧极差 < 0.005
```

滤光片命中条件：

```text
abs(当前值 / 空气基底 - 目标比例) <= ratio_tolerance / 100
```

例如 30% 滤光片、`ratio_tolerance = 5` 时：

```text
25% ~ 35%
```

回到空气条件：

```text
abs(当前值 / 空气基底 - 1.0) <= air_tolerance / 100
```

例如 `air_tolerance = 5` 时：

```text
95% ~ 105%
```

通道回到空气并稳定后，软件会用新的空气稳定均值更新该通道空气基底。

## 打包

推荐直接运行：

```powershell
.\打包指令.bat
```

脚本会自动查找当前目录下的 `.spec` 并执行：

```powershell
pyinstaller --clean 分光光度滤光片记录.spec
```

也可以手动执行：

```powershell
pyinstaller --clean "分光光度滤光片记录.spec"
```

当前 `.spec` 会把 `assets` 打进程序，并使用：

```text
assets\app_icon_spectro_transparent.ico
```

打包后请把 `config.ini` 放到 EXE 同目录。程序运行时会从 EXE 所在目录读取配置，并把 `data` 输出到 EXE 同目录下。

## 串口线序备注

项目根目录的 `串口连接.txt` 当前记录：

```text
黑-黑
红-黄
绿-紫
```

## Git 注意事项

仓库地址：

```text
git@github.com:BzhChara/spectrophotometer_filter_record.git
```

不要提交运行生成物或打包产物：

```text
__pycache__\
.pytest_cache\
build\
dist\
data\
```

这些目录已写入 `.gitignore`。

## 当前注意事项

- `config.ini` 的 `filter_ratio` 必须是 `0/10/20/30`。
- 如果只接一个 12 通道模块，请选择 `CH1-CH12` 或 `CH13-CH24`，否则全通道模式会等待未连接通道。
- 滤光片模式下，空气基底阶段不会写入原始 CSV；基底完成后才开始写入正式原始数据。
- 稳定值 Excel 只有记录到至少一个稳定滤光片值后才会保留。
- 当前数据是未经过软件校准的原始强度类读数，滤光片比例按 `当前值 / 空气基底` 判断。
