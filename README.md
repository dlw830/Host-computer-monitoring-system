# 🔌 上位机监控系统

一款功能强大的串口数据采集与监控上位机软件，专为工业数据采集、传感器监测、嵌入式系统调试等场景设计。支持自定义协议解析、实时曲线绘制、数据记录与导出等功能。

![Version](https://img.shields.io/badge/version-1.12-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-yellow)

## ✨ 主要功能

### 📡 串口通信
- 支持常用波特率（2400 ~ 921600 bps）
- 可配置数据位、停止位、校验位
- 自动扫描可用串口
- 智能断线恢复功能
- 实时显示收发数据

### 📊 数据协议解析
- **默认协议**：支持固定38字节数据帧
- **自定义协议**：可配置帧头、帧尾、数据长度
- **CRC校验**：支持多种校验方式
  - CRC16-XMODEM
  - CRC16-CCITT
  - CRC16-MODBUS
  - 累加和
  - 异或校验
- **多协议并行**：支持同时使用多种协议解析

### 📈 实时曲线绘制
- 最多支持 **50 条曲线**同时显示
- 自定义曲线配置（起始字节、数据类型、单位等）
- 支持多种数据类型：
  - uint8/int8, uint16/int16, uint32/int32
  - float, double
  - 大小端可选 (LE/BE)
  - 字符串 (ASCII)
  - **位模式**：单独记录某字节的某位状态
- 曲线颜色自定义
- 实时缩放、平移、重置视图
- 自动/手动缩放模式

### 🖥️ 自定义显示窗口
- 最多支持 **50 个显示窗口**
- 实时显示解析后的数据值
- 可配置数据源、单位、小数位数
- 支持数学运算（乘系数、除系数、偏移量）
- 窗口可拖拽、停靠、浮动

### 🔢 位(Bit)状态显示
- 最多支持 **50 个位显示窗口**
- 实时显示字节的每个位状态（0/1）
- 自定义每个位的名称和含义
- 红/绿指示灯直观显示
- 适用于状态监控、标志位查看

### 📝 数据记录与导出
- **Excel 导出**：支持导出为 .xlsx 格式
- **图像导出**：将曲线保存为 PNG 图片
- **原始数据导出**：保存 HEX 格式原始数据
- **自动保存**：每 2 小时自动保存数据
- **长时间记录**：支持 50000+ 数据点（约14小时@1Hz）

### 🕐 时钟同步
- **系统时钟校准**：发送当前时间到设备
- **接收时钟显示**：解析设备返回的时间戳
- 支持自定义时间字段位置和格式

### 🚀 预设命令
- 最多支持 **50 个预设命令**
- 支持文本/HEX 格式发送
- 自动添加 CRC 校验
- **周期发送**：支持定时自动发送（精度到毫秒）
- **数据填充**：动态替换命令中的数据字段
- 可配置发送时间戳

### 📋 数据监视器
- 实时显示收发数据
- HEX/ASCII 双格式显示
- 时间戳标记
- 接收/发送数据计数
- 自动滚动与暂停
- 支持清空、导出

### 🎨 界面特性
- **Apple 毛玻璃风格**：现代化 UI 设计
- **可定制布局**：所有窗口可自由拖拽、停靠
- **布局保存/恢复**：记住你的窗口布局
- **背景图片**：支持自定义背景和透明度
- **多标签管理**：曲线、配置、数据监视等分标签显示
- **启动画面**：美观的启动动画

### 🛠️ 其他功能
- 配置文件导入/导出
- 窗口布局导入/导出
- 日志记录与导出
- 帧间隔时间调节（防止粘包）
- 断线智能恢复
- 模板库（预设命令、协议等）

## 🚀 快速开始

### 环境要求

- Python 3.8 或更高版本
- Windows 操作系统（推荐 Windows 10/11）

### 安装依赖

```bash
pip install -r requirements.txt
```

依赖包包括：
```
PyQt5>=5.15.0
pyserial>=3.5
matplotlib>=3.5.0
pandas>=1.3.0
openpyxl>=3.0.0
```

### 运行程序

```bash
python serial_monitor.py
```

或者直接双击运行 `serial_monitor.py`

### 首次使用

1. **连接设备**：
   - 在左侧"串口配置"区域选择 COM 口
   - 点击"配置参数"设置波特率等参数
   - 点击"打开串口"开始通信

2. **配置曲线**：
   - 点击"配置曲线"按钮
   - 设置曲线名称、起始字节、数据类型等
   - 勾选"启用此曲线"和"记录此曲线"

3. **开始记录**：
   - 点击"开始记录"按钮
   - 数据会实时显示在曲线图上
   - 点击"停止记录"并选择保存路径

## 📦 项目结构

```
serial_monitor/
├── serial_monitor.py       # 主程序
├── serial_config.json      # 配置文件（自动生成）
├── window_layout.json      # 窗口布局（自动生成）
├── README.md              # 说明文档
├── requirements.txt       # 依赖列表
└── data/                  # 数据导出目录（可选）
    ├── excel/            # Excel 导出
    ├── images/           # 图像导出
    └── raw/              # 原始数据
```

## 💡 使用示例

### 1. 默认协议（38字节）

假设设备发送以下数据帧：

```
A8 A8 01 F4 00 00 02 58 01 2C ...（共38字节）... AA AA
```

**曲线配置示例**：

| 曲线名称 | 起始字节 | 字节数 | 数据类型 | 除系数 | 单位 | 说明 |
|---------|---------|--------|----------|--------|------|------|
| CO浓度 | 2 | 2 | uint16 (LE) | 1.0 | ppm | 字节2-3 |
| 平均CO | 4 | 4 | uint32 (LE) | 600.0 | ppm | 字节4-7 |
| 温度 | 8 | 2 | int16 (LE) | 10.0 | °C | 字节8-9 |

解析结果：
- CO浓度 = `0x01F4` = 500 ppm
- 平均CO = `0x00000258` / 600 = 1 ppm
- 温度 = `0x012C` / 10 = 30.0°C

### 2. 自定义协议

**配置协议**：
1. 点击"协议管理"按钮
2. 添加新协议：
   - 协议名称：温湿度协议
   - 帧头：`AA 55`
   - 数据长度：10
   - 帧尾：`55 AA`
   - CRC校验：CRC16-MODBUS
3. 启用该协议

**配置曲线**：
- 在曲线配置中选择"温湿度协议"作为数据来源
- 其余配置同上

### 3. 位显示示例

假设字节17（`0x2F` = 0010 1111）表示设备状态：

**位配置**：
- Bit 0: 电源状态
- Bit 1: 加热状态
- Bit 2: 风扇状态
- Bit 3: 报警状态
- Bit 4: 未使用
- Bit 5: 通讯状态
- Bit 6: 未使用
- Bit 7: 未使用

**显示结果**：
- 电源状态：🟢（1）
- 加热状态：🟢（1）
- 风扇状态：🟢（1）
- 报警状态：🟢（1）
- 通讯状态：🟢（1）

### 4. 预设命令示例

**命令1：查询版本**
```
命令内容：AT+VER\r\n
HEX模式：否
周期发送：否
```

**命令2：设置采样间隔（HEX）**
```
命令内容：A8 A8 01 00 00 AA AA
HEX模式：是
CRC校验：CRC16-MODBUS
周期发送：是（每1.0秒）
```

## 🎨 界面说明

### 主窗口布局

```
┌─────────────────────────────────────────────────────────┐
│  菜单栏：文件 | 窗口 | 帮助                              │
├───────────┬─────────────────────────────┬───────────────┤
│           │                             │               │
│  串口配置  │      实时曲线图                │ 自定义显示窗口 │
│           │                             │               │
│  (左侧)   │      (中心区域)              │   (右侧)       │
│           │                             │               │
│           ├─────────────────────────────┤               │
│           │                             │               │
│           │      数据监视器              │               │
│           │                             │               │
├───────────┴─────────────────────────────┴───────────────┤
│  发送控制 | 预设命令                                      │
├───────────────────────────────────────────────────────┤
│  日志输出                                              │
└───────────────────────────────────────────────────────┘
```

### 窗口说明

- **串口配置**：选择串口、配置参数、打开/关闭串口
- **实时曲线图**：显示实时数据曲线，支持缩放、导出
- **自定义显示窗口**：显示解析后的数值
- **位显示窗口**：显示字节位状态
- **数据监视器**：显示原始收发数据
- **发送控制**：手动发送数据、文件发送
- **预设命令**：快速发送常用命令
- **接收时钟**：显示设备时间
- **日志输出**：系统日志记录

所有窗口均可：
- 拖拽调整位置
- 停靠到边缘或中心
- 浮动显示
- 关闭/显示（通过菜单栏"窗口"）

## 🛠️ 技术栈

- **GUI框架**：PyQt5
- **串口通信**：pyserial
- **数据可视化**：matplotlib
- **数据处理**：pandas, numpy
- **Excel导出**：openpyxl
- **Python版本**：3.8+

## 📝 配置文件

### serial_config.json

保存所有配置信息：

```json
{
    "port": "COM3",
    "baudrate": "115200",
    "databits": "8",
    "stopbits": "1",
    "parity": "None",
    "curves": [...],
    "custom_displays": [...],
    "bit_displays": [...],
    "preset_commands": [...],
    "protocol_configs": [...],
    "clock_config": {...},
    "auto_save_path": "D:/data",
    "background_image": "",
    "background_opacity": 0.3
}
```

### window_layout.json

保存窗口布局状态：

```json
{
    "geometry": "...",
    "windowState": "...",
    "dockStates": {
        "plot_dock": {...},
        "config_dock": {...},
        ...
    }
}
```

## 🔄 更新日志

### v1.12 (2026-01-01)
- ✅ 完整的串口通信功能
- ✅ 自定义协议解析系统
- ✅ 50条曲线同时绘制
- ✅ 50个自定义显示窗口
- ✅ 50个位显示窗口
- ✅ 50个预设命令
- ✅ 数据导出（Excel/图像/原始数据）
- ✅ 智能断线恢复
- ✅ 可定制窗口布局
- ✅ Apple风格UI设计
- ✅ 背景图片支持
- ✅ 完善的日志系统

### 未来计划
- [ ] 支持 TCP/UDP 通信
- [ ] 数据库存储功能
- [ ] 更多数据分析工具（FFT、滤波等）
- [ ] 跨平台支持（macOS, Linux）
- [ ] 插件系统

## ❓ 常见问题

### 1. 串口打开失败
- 检查串口是否被其他程序占用
- 确认设备已正确连接
- 尝试更换 USB 接口
- 检查设备驱动是否正常

### 2. 数据解析错误
- 检查协议配置是否正确（帧头、帧尾、长度）
- 确认 CRC 校验类型
- 查看数据监视器中的原始数据
- 调整"帧间隔时间"参数

### 3. 曲线不显示
- 确认曲线已启用
- 检查起始字节和数据类型配置
- 查看自定义显示窗口是否有数值
- 点击"重置视图"按钮

### 4. 程序启动慢
- 首次启动会加载配置和初始化
- 后续启动会快很多
- 可以清空历史数据减小配置文件

### 5. 导出 Excel 失败
- 确保已安装 openpyxl 库
- 检查磁盘空间是否充足
- 确认导出路径有写入权限
- 关闭已打开的同名 Excel 文件

## 📄 许可证

本项目基于 **MIT License** 开源发布。  
你可以自由地使用、修改和分发本项目，但需保留原始版权声明。

```
MIT License

Copyright (c) 2026 力华亘金

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 💖 支持与赞助（打赏）

如果这个项目对你有帮助，欢迎通过以下方式支持作者的持续维护与改进：

* ⭐ **Star 本项目**（这是最好的支持方式）
* 🍴 **Fork 并参与贡献**
* 💬 提出 Issue / 改进建议
* ☕ **自愿打赏（非强制）**

### 打赏方式

| 平台 | 说明 |
| --- | --- |
| 微信支付 | 扫描下方二维码 |

<img width="300" alt="微信打赏二维码" src="https://github.com/user-attachments/assets/d9a988e0-b9b8-48da-aace-9c2fbe492b55" />

> 打赏完全自愿，不影响项目的任何功能或授权。

---

## 🤝 贡献指南

欢迎提交 **Issue** 和 **Pull Request**！

建议流程：

1. Fork 本仓库
2. 创建新分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

如是较大改动，建议先提交 Issue 讨论。

### 贡献方向

- 🐛 Bug 修复
- ✨ 新功能开发
- 📝 文档完善
- 🎨 UI/UX 改进
- 🌍 国际化支持
- 🧪 测试用例

---

## 📧 联系方式

* **Email**：[1013344248@qq.com](mailto:1013344248@qq.com)
* **GitHub**：[@dlw830](https://github.com/dlw830)
* **项目主页**：[https://github.com/dlw830/serial_monitor](https://github.com/dlw830/serial_monitor)

---

## 🙏 致谢

感谢以下开源项目：

- [PyQt5](https://www.riverbankcomputing.com/software/pyqt/) - 强大的 Python GUI 框架
- [pyserial](https://github.com/pyserial/pyserial) - Python 串口通信库
- [matplotlib](https://matplotlib.org/) - Python 绘图库
- [pandas](https://pandas.pydata.org/) - 数据分析库

---

## 📸 截图展示

### 主界面
<img width="2878" height="1704" alt="image" src="https://github.com/user-attachments/assets/0bb18843-403a-48c3-8f48-a0fef327fbb9" />

---

**Enjoy monitoring!** 🚀  
如果你觉得这个项目有价值，别忘了点个 ⭐

---

**Made with ❤️ by 力华亘金**


