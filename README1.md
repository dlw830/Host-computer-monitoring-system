# 🔧 C语言函数及变量定义和命名工具

一款专为嵌入式C语言开发设计的变量命名和代码生成工具。帮助开发者快速生成符合匈牙利命名法的变量定义，并自动添加详细注释。

## ✨ 主要功能

### 📝 变量定义
- 支持所有C语言基本类型 (int8_t ~ uint64_t, float, double)
- 自动显示类型取值范围
- 智能中英文翻译
- 实时预览生成代码
- 详细的命名解析注释

### 📊 数组定义
- 可视化数组大小配置
- 自动计算内存占用
- 智能存储单位转换 (bytes/KB/MB)
- 数组命名前缀自动添加

### 🏗️ 结构体定义
- 可视化成员管理 (增删改查)
- 自动计算内存对齐和填充
- 显示总内存占用
- 支持成员排序和编辑

### 🔀 联合体定义
- 成员可视化管理
- 自动计算联合体大小（最大成员）
- 共享内存警告提示
- 详细注释生成

### 📋 枚举定义
- 枚举值可视化编辑
- 支持指定值或自动编号
- 成员上下移动功能
- 统计显示（指定值/自动值数量）

### 🔍 变量解析
- 解析现有C代码
- 识别结构体、联合体、枚举定义
- 变量命名规范检查
- 内存布局分析
- 优化建议生成

### 🌐 翻译工具
- 中英文互译
- 本地词库支持
- 翻译历史记录
- 术语浏览功能

### 📚 模板库
- 保存常用代码模板
- 模板分类管理
- 导入/导出功能
- 快速复用代码

### ⚙️ 设置
- 命名规则自定义（全局/静态/结构体/数组前缀）
- 代码生成配置（注释语言、缩进方式）
- 界面主题切换
- 翻译引擎选择

## 🚀 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行程序

```bash
python main.py
```

或者直接双击运行：

```bash
cd c_naming_tool
python main.py
```

## 📦 项目结构

```
c_naming_tool/
├── config/                 # 配置文件
│   ├── type_ranges.json   # 类型范围定义
│   ├── term_database.json # 翻译词库
│   ├── settings.json      # 基本设置
│   ├── templates.json     # 模板库
│   └── app_settings.json  # 应用设置
├── core/                   # 核心模块
│   ├── type_info.py       # 类型信息管理
│   ├── translator.py      # 翻译引擎
│   └── naming.py          # 命名生成器
├── ui/                     # 界面模块
│   ├── main_window.py     # 主窗口
│   ├── variable_panel.py  # 变量定义面板
│   ├── array_panel.py     # 数组定义面板
│   ├── struct_panel.py    # 结构体面板
│   ├── union_panel.py     # 联合体面板
│   ├── enum_panel.py      # 枚举面板
│   ├── parser_panel.py    # 变量解析面板
│   ├── translator_panel.py # 翻译工具面板
│   ├── template_panel.py  # 模板库面板
│   ├── settings_panel.py  # 设置面板
│   └── styles/            # 样式表
│       └── macos_light.qss
├── utils/                  # 工具模块
│   └── code_generator.py  # 代码生成器
├── main.py                # 程序入口
├── requirements.txt       # 依赖列表
└── README.md             # 说明文档
```

## 💡 使用示例

### 1. 变量定义

输入：
- 类型：uint16_t
- 中文名：温度值
- 作用域：全局

输出：
```c
/*
 * 温度值 (Temperature Value)
 * 类型: uint16_t
 * 取值范围: 0 ~ 65535
 * 命名解析:
 *   g: 全局变量 (Global)
 *   u16: uint16_t 无符号16位整数
 *   temperature_value: 温度值
 */
uint16_t gu16_temperature_value;
```

### 2. 数组定义

输入：
- 类型：float
- 中文名：校准表
- 大小：10

输出：
```c
/*
 * 校准表 (Calibration Table)
 * 元素类型: float
 * 元素范围: -3.4e38 ~ 3.4e38 (精度: 6位)
 * 数组大小: 10
 * 总内存: 40 bytes
 * 命名解析:
 *   g: 全局变量 (Global)
 *   fa: float数组
 *   calibration_table: 校准表
 */
float gfa_calibration_table[10];
```

### 3. 结构体定义

输入：
- 结构体名：adc_config
- 成员：
  - uint8_t channel (通道)
  - uint32_t sample_rate (采样率)
  - uint8_t resolution (分辨率)

输出：
```c
/*
 * ADC配置结构体
 * 总大小: 12 bytes (含4字节对齐填充)
 * 内存布局:
 *   偏移 0: channel (1 byte)
 *   偏移 1-3: [填充] (3 bytes)
 *   偏移 4: sample_rate (4 bytes)
 *   偏移 8: resolution (1 byte)
 *   偏移 9-11: [填充] (3 bytes)
 */
typedef struct {
    uint8_t channel;        // 通道 [0-255]
    uint32_t sample_rate;   // 采样率 [0-4294967295]
    uint8_t resolution;     // 分辨率 [0-255]
} adc_config_t;

adc_config_t g_st_adc_config;  // 全局ADC配置
```

### 4. 联合体定义

输入：
- 联合体名：data_union
- 成员：
  - uint32_t word (字数据)
  - uint8_t bytes[4] (字节数组)

输出：
```c
/*
 * 数据联合体
 * 大小: 4 bytes (最大成员: word)
 * 警告: 所有成员共享同一块内存
 */
typedef union {
    uint32_t word;      // 字数据 [0-4294967295]
    uint8_t bytes[4];   // 字节数组 [每个0-255]
} data_union_t;

data_union_t g_un_data_union;  // 全局数据联合体
```

### 5. 枚举定义

输入：
- 枚举名：device_status
- 枚举值：
  - IDLE (0)
  - RUNNING (1)
  - ERROR (自动)

输出：
```c
/*
 * 设备状态枚举
 * 枚举值: 3个
 */
typedef enum {
    IDLE = 0,      // 空闲
    RUNNING = 1,   // 运行中
    ERROR          // 错误 (自动: 2)
} device_status_t;
```

## 🎨 界面特性

- 🍎 **MacOS风格设计**：清爽的浅色主题
- 📱 **响应式布局**：自适应窗口大小
- 🎯 **实时预览**：即时查看生成代码
- 💬 **智能提示**：翻译建议和范围提示
- 🔄 **历史记录**：保存翻译和模板历史

## 🛠️ 技术栈

- **GUI框架**: PyQt6
- **Python版本**: 3.8+
- **样式系统**: QSS (Qt Style Sheets)
- **数据存储**: JSON

## 📝 命名规范

### 作用域前缀
- `g_`: 全局变量 (Global)
- `s_`: 静态变量 (Static)
- `l_`: 局部变量 (Local)

### 类型前缀
- `i8, i16, i32, i64`: 有符号整数
- `u8, u16, u32, u64`: 无符号整数
- `f`: float
- `d`: double
- `st`: 结构体 (struct)
- `un`: 联合体 (union)

### 数组标识
- `a_`: 数组前缀

### 示例
- `gu16_temperature_value`: 全局uint16_t温度值
- `gfa_calibration_table`: 全局float数组校准表
- `g_st_adc_config`: 全局ADC配置结构体

## 🔄 更新日志

### v1.0 (20251228)
- ✅ 完整的9大功能模块
- ✅ 变量、数组、结构体、联合体、枚举定义
- ✅ 变量解析和代码分析
- ✅ 翻译工具和术语管理
- ✅ 模板库和配置系统
- ✅ MacOS风格UI设计
- ✅ 完善的注释生成

## 📄 许可证

MIT License

## 💖 支持与赞助（打赏）

如果这个项目对你有帮助，欢迎通过以下方式支持作者的持续维护与改进：

* ⭐ **Star 本项目**（这是最好的支持方式）
* 🍴 **Fork 并参与贡献**
* 💬 提出 Issue / 改进建议
* ☕ **自愿打赏（非强制）**

### 打赏方式

| 平台              | 说明          |
| --------------- | ----------- |
| 微信         | 扫描下方二维码     |
<img width="902" height="890" alt="image" src="https://github.com/user-attachments/assets/d9a988e0-b9b8-48da-aace-9c2fbe492b55" />


> 打赏完全自愿，不影响项目的任何功能或授权。



## 📄 许可证

本项目基于 **MIT License** 开源发布。
你可以自由地使用、修改和分发本项目，但需保留原始版权声明。

---

## 🤝 贡献指南

欢迎提交 **Issue** 和 **Pull Request**！

建议流程：

1. Fork 本仓库
2. 新建分支进行修改
3. 提交 PR 并简要说明修改内容

如是较大改动，建议先提交 Issue 讨论。

---

## 📧 联系方式

* Email：[1013344248@qq.com](mailto:1013344248@qq.com)
* GitHub：@dlw830

---

**Enjoy coding!** 🚀
如果你觉得这个项目有价值，别忘了点个 ⭐







