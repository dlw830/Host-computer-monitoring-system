# -*- coding: utf-8 -*-
"""
串口上位机 - 主程序
支持数据采集、实时曲线绘制、数据导出等功能
"""

import sys
import json
import os
import time
from datetime import datetime
from collections import deque
import struct

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QGroupBox, QLabel, QComboBox, QPushButton,
                             QTextEdit, QLineEdit, QCheckBox, QFileDialog, QSpinBox,
                             QMessageBox, QGridLayout, QTabWidget, QScrollArea, QDialog,
                             QDialogButtonBox, QDoubleSpinBox, QSplitter, QDockWidget, QMenuBar, QMenu, QAction,
                             QSizePolicy, QListWidget, QSplashScreen, QProgressBar)
from PyQt5.QtCore import QTimer, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QTextCursor, QPixmap, QPainter, QColor, QLinearGradient

import serial
import serial.tools.list_ports

# 配置matplotlib支持中文
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
import pandas as pd


class CRCCalculator:
    """CRC校验计算器"""
    
    @staticmethod
    def calculate_ccitt_crc16(data):
        """计算CCITT-CRC16校验码
        多项式: 0x1021
        初始值: 0xFFFF
        """
        crc = 0xFFFF
        for byte in data:
            crc ^= (byte << 8)
            for _ in range(8):
                if crc & 0x8000:
                    crc = (crc << 1) ^ 0x1021
                else:
                    crc = crc << 1
                crc &= 0xFFFF
        return crc
    
    @staticmethod
    def calculate_modbus_crc16(data):
        """计算Modbus-CRC16校验码
        多项式: 0x8005
        初始值: 0xFFFF
        低字节在前
        """
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc = (crc >> 1) ^ 0xA001
                else:
                    crc = crc >> 1
        return crc
    
    @staticmethod
    def calculate_crc16_xmodem(data):
        """计算CRC16-XMODEM校验码
        多项式: 0x1021
        初始值: 0x0000
        """
        crc = 0x0000
        for byte in data:
            crc ^= (byte << 8)
            for _ in range(8):
                if crc & 0x8000:
                    crc = (crc << 1) ^ 0x1021
                else:
                    crc = crc << 1
                crc &= 0xFFFF
        return crc
    
    @staticmethod
    def calculate_sum_check(data):
        """计算累加和校验(取低8位)"""
        return sum(data) & 0xFF
    
    @staticmethod
    def calculate_xor_check(data):
        """计算异或校验"""
        result = 0
        for byte in data:
            result ^= byte
        return result


class SerialThread(QThread):
    """串口接收线程"""
    data_received = pyqtSignal(bytes)
    connection_lost = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.serial_port = None
        self.running = False
        
    def set_serial(self, serial_port):
        self.serial_port = serial_port
        
    def run(self):
        self.running = True
        while self.running:
            if self.serial_port and self.serial_port.is_open:
                try:
                    if self.serial_port.in_waiting > 0:
                        data = self.serial_port.read(self.serial_port.in_waiting)
                        self.data_received.emit(data)
                except Exception as e:
                    print(f"读取串口数据错误: {e}")
                    self.connection_lost.emit()
                    self.running = False
            self.msleep(10)
    
    def stop(self):
        self.running = False
        self.wait()


class SplashScreenWidget(QSplashScreen):
    """自定义启动画面"""
    
    def __init__(self):
        # 创建一个渐变背景的QPixmap
        pixmap = QPixmap(600, 400)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # 绘制渐变背景
        gradient = QLinearGradient(0, 0, 0, 400)
        gradient.setColorAt(0, QColor(70, 130, 180, 230))  # 钢蓝色
        gradient.setColorAt(1, QColor(100, 149, 237, 230))  # 矢车菊蓝
        painter.setBrush(gradient)
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(0, 0, 600, 400, 20, 20)
        
        # 绘制标题
        painter.setPen(QColor(255, 255, 255))
        title_font = QFont('Microsoft YaHei', 24, QFont.Bold)
        painter.setFont(title_font)
        painter.drawText(0, 80, 600, 50, Qt.AlignCenter, '力华亘金-上位机监控系统')
        
        # 绘制版本信息
        version_font = QFont('Microsoft YaHei', 10)
        painter.setFont(version_font)
        painter.drawText(0, 130, 600, 30, Qt.AlignCenter, 'Version 1.12')
        
        painter.end()
        
        super().__init__(pixmap)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        
        # 创建进度条
        self.progress = QProgressBar(self)
        self.progress.setGeometry(50, 320, 500, 25)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #4682B4;
                border-radius: 12px;
                text-align: center;
                background-color: rgba(255, 255, 255, 100);
                color: white;
                font-weight: bold;
                font-size: 11px;
            }
            QProgressBar::chunk {
                border-radius: 10px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #5cb85c, stop:1 #4cae4c);
            }
        """)
        self.progress.setValue(0)
        
        # 状态标签
        self.status_label = QLabel('正在初始化...', self)
        self.status_label.setGeometry(50, 280, 500, 30)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 12px;
                font-weight: 500;
            }
        """)
    
    def update_progress(self, value, message=''):
        """更新进度"""
        self.progress.setValue(value)
        if message:
            self.status_label.setText(message)
        QApplication.processEvents()


class DataParser:
    """数据解析器 - 解析固定38字节协议"""
    
    FRAME_LENGTH = 38
    FRAME_HEADER = bytes([0xA8, 0xA8])
    FRAME_TAIL = bytes([0xAA, 0xAA])
    
    def __init__(self):
        self.buffer = bytearray()
    
    def add_data(self, data):
        """添加接收到的数据"""
        self.buffer.extend(data)
        
    def parse(self):
        """解析数据帧,返回解析结果列表"""
        results = []
        
        while len(self.buffer) >= self.FRAME_LENGTH:
            # 查找帧头
            header_index = self.buffer.find(self.FRAME_HEADER)
            
            if header_index == -1:
                # 没有找到帧头,清空缓冲区
                self.buffer.clear()
                break
            
            # 删除帧头之前的无效数据
            if header_index > 0:
                self.buffer = self.buffer[header_index:]
            
            # 检查是否有完整帧
            if len(self.buffer) < self.FRAME_LENGTH:
                break
            
            # 提取一帧数据
            frame = self.buffer[:self.FRAME_LENGTH]
            
            # 验证帧尾
            if frame[36:38] != self.FRAME_TAIL:
                # 帧尾不匹配,删除当前帧头,继续查找
                self.buffer = self.buffer[2:]
                continue
            
            # 解析数据
            try:
                # CO浓度 (Byte 2-3, uint16 LE)
                co_concentration = struct.unpack('<H', frame[2:4])[0]
                
                # 平均CO浓度 (Byte 4-7, uint32 LE, 需要除以600)
                co_avg_raw = struct.unpack('<I', frame[4:8])[0]
                co_avg = co_avg_raw / 600.0
                
                # 温度 (Byte 8-9, int16 LE, 单位0.1°C)
                temp_raw = struct.unpack('<h', frame[8:10])[0]
                temperature = temp_raw / 10.0
                
                result = {
                    'timestamp': datetime.now(),
                    'co_concentration': co_concentration,
                    'co_avg': co_avg,
                    'temperature': temperature,
                    'raw_frame': frame
                }
                results.append(result)
                
            except Exception as e:
                print(f"解析数据错误: {e}")
            
            # 删除已处理的帧
            self.buffer = self.buffer[self.FRAME_LENGTH:]
        
        return results


class GenericProtocolParser:
    """通用协议解析器"""
    
    def __init__(self, protocol_config):
        """
        初始化解析器
        protocol_config: 协议配置字典
        """
        self.config = protocol_config
        self.buffer = bytearray()
        self.name = protocol_config.get('name', '未命名协议')
        self.header = bytes(protocol_config.get('header', []))
        self.tail = bytes(protocol_config.get('tail', []))
        self.length = protocol_config.get('length', 0)
        self.crc_type = protocol_config.get('crc_type', '无')
        self.enabled = protocol_config.get('enabled', True)
    
    def add_data(self, data):
        """添加接收到的数据"""
        self.buffer.extend(data)
        # 限制缓冲区大小，防止内存溢出
        if len(self.buffer) > 10240:
            self.buffer = self.buffer[-5120:]
    
    def parse(self):
        """解析数据帧，返回解析结果列表"""
        results = []
        
        if not self.enabled:
            return results
        
        while len(self.buffer) >= self.length:
            # 查找帧头
            if len(self.header) > 0:
                header_index = self.buffer.find(self.header)
                
                if header_index == -1:
                    # 没有找到帧头，保留最后len(header)-1字节，防止帧头被截断
                    if len(self.buffer) > len(self.header):
                        self.buffer = self.buffer[-(len(self.header)-1):]
                    break
                
                # 删除帧头之前的无效数据
                if header_index > 0:
                    self.buffer = self.buffer[header_index:]
            
            # 检查是否有完整帧
            if len(self.buffer) < self.length:
                break
            
            # 提取一帧数据
            frame = bytes(self.buffer[:self.length])
            
            # 验证帧尾
            if len(self.tail) > 0:
                tail_start = self.length - len(self.tail)
                if frame[tail_start:] != self.tail:
                    # 帧尾不匹配，删除当前帧头，继续查找
                    skip_len = max(1, len(self.header))
                    self.buffer = self.buffer[skip_len:]
                    continue
            
            # 验证CRC（若配置了CRC，必须校验通过才解析）
            if not self._verify_crc(frame):
                # CRC校验失败，记录日志并跳过该帧
                frame_hex = ' '.join(f'{b:02X}' for b in frame)
                print(f"[{self.name}] CRC校验失败，丢弃帧: {frame_hex}")
                skip_len = max(1, len(self.header))
                self.buffer = self.buffer[skip_len:]
                continue
            
            # 解析成功
            result = {
                'timestamp': datetime.now(),
                'protocol_name': self.name,
                'raw_frame': frame,
                'frame_hex': ' '.join(f'{b:02X}' for b in frame)
            }
            results.append(result)
            
            # 删除已处理的帧
            self.buffer = self.buffer[self.length:]
        
        return results
    
    def _verify_crc(self, frame):
        """验证CRC校验（若配置了CRC，必须校验通过）"""
        if self.crc_type == '无':
            return True
        
        # CRC位于帧的最后2个字节（或1个字节）
        if self.crc_type in ['CRC16-XMODEM', 'CRC16-CCITT', 'CRC16-MODBUS']:
            if len(frame) < 3:
                return False
            data = frame[:-2]
            
            # 根据不同CRC类型使用不同的字节序
            if self.crc_type == 'CRC16-MODBUS':
                # Modbus CRC使用小端序（低字节在前）
                crc_received = struct.unpack('<H', frame[-2:])[0]
                crc_calculated = CRCCalculator.calculate_modbus_crc16(data)
            elif self.crc_type == 'CRC16-XMODEM':
                # XMODEM使用小端序（低字节在前）
                crc_received = struct.unpack('<H', frame[-2:])[0]
                crc_calculated = CRCCalculator.calculate_crc16_xmodem(data)
            else:  # CRC16-CCITT
                # CCITT使用大端序（高字节在前）
                crc_received = struct.unpack('>H', frame[-2:])[0]
                crc_calculated = CRCCalculator.calculate_ccitt_crc16(data)
            
            # 调试输出（可选）
            if crc_received != crc_calculated:
                print(f"[{self.name}] CRC校验: 接收={crc_received:04X}, 计算={crc_calculated:04X}")
            
            return crc_received == crc_calculated
        
        elif self.crc_type in ['累加和', '异或']:
            if len(frame) < 2:
                return False
            data = frame[:-1]
            check_received = frame[-1]
            
            if self.crc_type == '累加和':
                check_calculated = CRCCalculator.calculate_sum_check(data)
            else:  # 异或
                check_calculated = CRCCalculator.calculate_xor_check(data)
            
            # 调试输出（可选）
            if check_received != check_calculated:
                print(f"[{self.name}] 校验失败: 接收={check_received:02X}, 计算={check_calculated:02X}")
            
            return check_received == check_calculated
        
        return True


class PlotCanvas(FigureCanvas):
    """绘图画布"""
    
    def __init__(self, parent=None, width=8, height=6, dpi=100):
        # 设置透明背景和Apple风格
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.fig.patch.set_facecolor('none')
        self.fig.patch.set_alpha(0.0)
        super().__init__(self.fig)
        self.setParent(parent)
        
        # 创建单个子图
        self.ax = self.fig.add_subplot(111)
        
        # Apple风格设置
        self.ax.set_ylabel('数值', fontsize=10, color='#1d1d1f', fontweight=500)
        self.ax.set_xlabel('时间 (s)', fontsize=10, color='#1d1d1f', fontweight=500)
        self.ax.set_title('数据监测', fontsize=11, pad=10, color='#1d1d1f', fontweight=600)
        self.ax.grid(True, alpha=0.2, linestyle='--', color='#86868B')
        self.ax.set_facecolor((1, 1, 1, 0.7))  # 半透明白色背景
        
        # 设置坐标轴样式
        for spine in ['top', 'right']:
            self.ax.spines[spine].set_color('none')
        for spine in ['left', 'bottom']:
            self.ax.spines[spine].set_color('#d1d1d6')
            self.ax.spines[spine].set_linewidth(1.2)
        
        # 刻度标签颜色
        self.ax.tick_params(colors='#1d1d1f', labelsize=9)
        
        # 启用交互式缩放和平移
        self.ax.set_xlim(auto=True)
        self.ax.set_ylim(auto=True)
        
        self.fig.tight_layout(pad=2.0)
        
        # 用于记录是否使用自动缩放
        self.auto_scale = True
        
        # 数据存储 - 取消点数限制，使用list
        self.time_data = []  # 相对时间(秒)，用于图表显示
        self.timestamp_data = []  # 绝对时间戳，用于Excel导出
        self.recording_start_time = None  # 记录开始时间
        self.last_auto_save_time = None  # 上次自动保存时间
        self.auto_save_interval = 2 * 3600  # 2小时自动保存间隔（秒）
        self.auto_save_path = ''  # 自动保存路径
        
        # 曲线配置（最多50条）
        self.curve_configs = []
        self.curve_data = []  # 每条曲线的数据list
        self.lines = []  # 每条曲线的Line2D对象
        
        # 颜色映射
        self.color_map = {
            '蓝色': '#1f77b4',
            '绿色': '#2ca02c',
            '红色': '#d62728',
            '橙色': '#ff7f0e',
            '紫色': '#9467bd',
            '青色': '#17becf',
            '品红': '#e377c2'
        }
        
        self.start_time = None
        
        # 初始化默认曲线配置（兼容旧版本）
        self.init_default_curves()
    
    def init_default_curves(self):
        """初始化默认的3条曲线配置"""
        default_curves = [
            {
                'name': 'CO浓度',
                'start_byte': 2,
                'byte_count': 2,
                'data_type': 'uint16 (LE)',
                'coefficient': 1.0,
                'divisor': 1.0,
                'offset': 0.0,
                'unit': 'ppm',
                'color': '蓝色',
                'enabled': True,
                'record': True,
                'bit_mode': False,
                'bit_index': 0
            },
            {
                'name': '平均CO浓度',
                'start_byte': 4,
                'byte_count': 4,
                'data_type': 'uint32 (LE)',
                'coefficient': 1.0,
                'divisor': 600.0,
                'offset': 0.0,
                'unit': 'ppm',
                'color': '绿色',
                'enabled': True,
                'record': True,
                'bit_mode': False,
                'bit_index': 0
            },
            {
                'name': '温度',
                'start_byte': 8,
                'byte_count': 2,
                'data_type': 'int16 (LE)',
                'coefficient': 1.0,
                'divisor': 10.0,
                'offset': 0.0,
                'unit': '°C',
                'color': '红色',
                'enabled': True,
                'record': True,
                'bit_mode': False,
                'bit_index': 0
            }
        ]
        self.set_curve_configs(default_curves)
    
    def set_curve_configs(self, configs):
        """设置曲线配置（最多50条）"""
        # 清除旧曲线
        for line in self.lines:
            line.remove()
        self.lines.clear()
        self.curve_data.clear()
        
        # 最多50条曲线
        self.curve_configs = configs[:50]
        
        # 创建新曲线
        for config in self.curve_configs:
            color = self.color_map.get(config.get('color', '蓝色'), '#1f77b4')
            unit = config.get('unit', '')
            name = config.get('name', '曲线')
            label = f"{name} ({unit})" if unit else name
            
            line, = self.ax.plot([], [], color, linewidth=2, label=label)
            line.set_visible(config.get('enabled', True))
            self.lines.append(line)
            
            # 为每条曲线创建数据存储（无限制）
            self.curve_data.append([])
        
        self.update_legend()
        self.draw()
    
    def update_legend(self):
        """更新图例"""
        visible_lines = []
        visible_labels = []
        for i, (line, config) in enumerate(zip(self.lines, self.curve_configs)):
            if config.get('enabled', True) and line.get_visible():
                visible_lines.append(line)
                unit = config.get('unit', '')
                name = config.get('name', f'曲线{i+1}')
                label = f"{name} ({unit})" if unit else name
                visible_labels.append(label)
        
        if visible_lines:
            self.ax.legend(visible_lines, visible_labels, loc='upper right', fontsize=9)
        else:
            self.ax.legend([], [])
    
    def set_curve_visibility(self, index, visible):
        """设置指定曲线的可见性"""
        if 0 <= index < len(self.lines):
            self.lines[index].set_visible(visible)
            if index < len(self.curve_configs):
                self.curve_configs[index]['enabled'] = visible
            self.update_legend()
            self.draw()
    
    def parse_value_from_frame(self, frame, config):
        """从数据帧中解析数值"""
        try:
            start = config['start_byte']
            count = config['byte_count']
            dtype = config['data_type']
            bit_mode = config.get('bit_mode', False)
            bit_index = config.get('bit_index', 0)
            
            if start + count > len(frame):
                return None
            
            # 位模式：提取指定字节的指定位
            if bit_mode:
                if start >= len(frame):
                    return None
                target_byte = frame[start]
                # 提取指定位 (0-7, 0为最低位LSB)
                bit_value = (target_byte >> bit_index) & 0x01
                return float(bit_value)
            
            # 普通模式：按数据类型解析
            data_bytes = frame[start:start+count]
            
            # 根据数据类型解析
            if dtype == 'uint8':
                value = data_bytes[0]
            elif dtype == 'int8':
                value = struct.unpack('b', data_bytes[:1])[0]
            elif dtype == 'uint16 (LE)':
                value = struct.unpack('<H', data_bytes[:2])[0]
            elif dtype == 'uint16 (BE)':
                value = struct.unpack('>H', data_bytes[:2])[0]
            elif dtype == 'int16 (LE)':
                value = struct.unpack('<h', data_bytes[:2])[0]
            elif dtype == 'int16 (BE)':
                value = struct.unpack('>h', data_bytes[:2])[0]
            elif dtype == 'uint32 (LE)':
                value = struct.unpack('<I', data_bytes[:4])[0]
            elif dtype == 'uint32 (BE)':
                value = struct.unpack('>I', data_bytes[:4])[0]
            elif dtype == 'int32 (LE)':
                value = struct.unpack('<i', data_bytes[:4])[0]
            elif dtype == 'int32 (BE)':
                value = struct.unpack('>i', data_bytes[:4])[0]
            elif dtype == 'float (LE)':
                value = struct.unpack('<f', data_bytes[:4])[0]
            elif dtype == 'float (BE)':
                value = struct.unpack('>f', data_bytes[:4])[0]
            elif dtype == 'double (LE)':
                value = struct.unpack('<d', data_bytes[:8])[0]
            elif dtype == 'double (BE)':
                value = struct.unpack('>d', data_bytes[:8])[0]
            else:
                return None
            
            # 应用系数和偏移
            value = (value * config['coefficient'] / config['divisor']) + config['offset']
            return value
            
        except Exception as e:
            print(f"解析曲线数据错误: {e}")
            return None
    
    def add_data(self, data_dict):
        """添加数据点（仅在接收到数据包时调用，不由定时器触发）"""
        if self.start_time is None:
            self.start_time = data_dict['timestamp']
            self.recording_start_time = data_dict['timestamp']
            self.last_auto_save_time = data_dict['timestamp']
        
        elapsed = (data_dict['timestamp'] - self.start_time).total_seconds()
        self.time_data.append(elapsed)
        self.timestamp_data.append(data_dict['timestamp'])  # 存储绝对时间戳
        
        # 从raw_frame解析各曲线的数据
        if 'raw_frame' in data_dict:
            frame = data_dict['raw_frame']
            protocol_name = data_dict.get('protocol_name', None)
            
            for i, config in enumerate(self.curve_configs):
                if i < len(self.curve_data):
                    # 检查协议是否匹配
                    curve_protocol = config.get('protocol', None)
                    if curve_protocol == protocol_name:
                        value = self.parse_value_from_frame(frame, config)
                        if value is not None:
                            self.curve_data[i].append(value)
                        else:
                            self.curve_data[i].append(0.0)
                    else:
                        # 协议不匹配，不更新此曲线（但需要占位以保持索引一致）
                        # 使用上一个值或0
                        if self.curve_data[i]:
                            self.curve_data[i].append(self.curve_data[i][-1])
                        else:
                            self.curve_data[i].append(0.0)
        
        # 限制最大数据点数，防止内存溢出（保留最近50000点，约14小时@1Hz）
        if len(self.time_data) > 50000:
            self.time_data.pop(0)
            self.timestamp_data.pop(0)
            for data_list in self.curve_data:
                if data_list:
                    data_list.pop(0)
        
        # 检查是否需要自动保存（2小时）
        current_time = data_dict['timestamp']
        if self.last_auto_save_time and \
           (current_time - self.last_auto_save_time).total_seconds() >= self.auto_save_interval:
            self.auto_save_and_reset()
            self.last_auto_save_time = current_time
        
    def zoom_x(self, factor):
        """缩放X轴
        factor > 1: 放大(显示更少数据)
        factor < 1: 缩小(显示更多数据)
        """
        self.auto_scale = False
        cur_xlim = self.ax.get_xlim()
        center = (cur_xlim[0] + cur_xlim[1]) / 2
        width = (cur_xlim[1] - cur_xlim[0]) * factor
        self.ax.set_xlim([center - width/2, center + width/2])
        self.draw()
    
    def zoom_y(self, factor):
        """缩放Y轴
        factor > 1: 放大(显示更少范围)
        factor < 1: 缩小(显示更大范围)
        """
        self.auto_scale = False
        cur_ylim = self.ax.get_ylim()
        center = (cur_ylim[0] + cur_ylim[1]) / 2
        height = (cur_ylim[1] - cur_ylim[0]) * factor
        self.ax.set_ylim([center - height/2, center + height/2])
        self.draw()
    
    def reset_view(self):
        """重置视图为自动缩放"""
        self.auto_scale = True
        if len(self.time_data) > 0:
            # 收集所有可见曲线的数据
            all_values = []
            for i, (config, data) in enumerate(zip(self.curve_configs, self.curve_data)):
                if config.get('enabled', True) and len(data) > 0:
                    all_values.extend(data)
            
            if self.time_data and all_values:
                self.ax.set_xlim(min(self.time_data), max(self.time_data))
                self.ax.set_ylim(min(all_values), max(all_values))
        
        self.draw()
    
    def update_plot(self):
        """更新曲线"""
        if len(self.time_data) == 0:
            return
        
        # 更新每条曲线的数据
        for i, (line, config, data) in enumerate(zip(self.lines, self.curve_configs, self.curve_data)):
            if config.get('enabled', True) and len(data) > 0:
                line.set_data(self.time_data, data)
        
        # 在自动缩放模式下更新视图
        if self.auto_scale:
            all_values = []
            for i, (config, data) in enumerate(zip(self.curve_configs, self.curve_data)):
                if config.get('enabled', True) and len(data) > 0:
                    all_values.extend(data)
            
            if self.time_data and all_values:
                self.ax.set_xlim(min(self.time_data), max(self.time_data))
                self.ax.set_ylim(min(all_values), max(all_values))
        
        self.draw()
    
    def auto_save_and_reset(self):
        """自动保存当前数据并重置"""
        if len(self.time_data) == 0:
            return
        
        try:
            # 生成文件名（带时间戳）
            if self.auto_save_path:
                base_path = self.auto_save_path
            else:
                base_path = os.path.join(os.getcwd(), 'auto_saved_data')
            
            # 创建目录
            os.makedirs(base_path, exist_ok=True)
            
            # 生成文件名
            timestamp_str = self.recording_start_time.strftime('%Y%m%d_%H%M%S')
            filename = f'curve_data_{timestamp_str}.xlsx'
            filepath = os.path.join(base_path, filename)
            
            # 导出数据
            df = self.get_data_frame()
            if df is not None:
                df.to_excel(filepath, index=False)
                print(f"自动保存曲线数据到: {filepath}")
            
            # 重置数据（开始新的记录周期）
            self.clear_data()
            
        except Exception as e:
            print(f"自动保存曲线数据失败: {e}")
    
    def set_auto_save_path(self, path):
        """设置自动保存路径"""
        self.auto_save_path = path
    
    def clear_data(self):
        """清空数据"""
        self.time_data.clear()
        self.timestamp_data.clear()
        for data in self.curve_data:
            data.clear()
        self.start_time = None
        self.recording_start_time = None
        
        for line in self.lines:
            line.set_data([], [])
        
        self.draw()
    
    def get_data_frame(self):
        """获取数据DataFrame用于导出（只导出标记为记录的曲线）"""
        if len(self.time_data) == 0:
            return None
        
        # 使用实际时间戳而不是相对秒数
        time_strings = [ts.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3] for ts in self.timestamp_data]
        data_dict = {'时间': time_strings}
        
        for i, (config, data) in enumerate(zip(self.curve_configs, self.curve_data)):
            # 只导出标记为记录的曲线
            record = config.get('record', True)  # 默认为True以兼容旧配置
            if record and len(data) > 0:
                name = config.get('name', f'曲线{i+1}')
                unit = config.get('unit', '')
                column_name = f"{name}({unit})" if unit else name
                data_dict[column_name] = data.copy()
        
        df = pd.DataFrame(data_dict)
        return df


class ProtocolConfigDialog(QDialog):
    """数据协议配置对话框"""
    
    def __init__(self, parent=None, protocol_data=None):
        super().__init__(parent)
        self.setWindowTitle('配置数据协议')
        self.setMinimumWidth(550)
        
        layout = QGridLayout(self)
        
        # 协议名称
        layout.addWidget(QLabel('协议名称:'), 0, 0)
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText('例如: 数据类型1')
        layout.addWidget(self.edit_name, 0, 1, 1, 3)
        
        # 帧头
        layout.addWidget(QLabel('帧头(Hex):'), 1, 0)
        self.edit_header = QLineEdit()
        self.edit_header.setPlaceholderText('例如: A8 A8 (用空格分隔)')
        layout.addWidget(self.edit_header, 1, 1, 1, 3)
        
        # 数据长度
        layout.addWidget(QLabel('数据长度:'), 2, 0)
        self.spin_length = QSpinBox()
        self.spin_length.setRange(1, 1024)
        self.spin_length.setValue(38)
        self.spin_length.setSuffix(' 字节')
        self.spin_length.setToolTip('完整帧的字节数(包含帧头帧尾)')
        layout.addWidget(self.spin_length, 2, 1)
        
        # 帧尾
        layout.addWidget(QLabel('帧尾(Hex):'), 2, 2)
        self.edit_tail = QLineEdit()
        self.edit_tail.setPlaceholderText('例如: AA AA (用空格分隔,无则留空)')
        layout.addWidget(self.edit_tail, 2, 3)
        
        # CRC校验
        layout.addWidget(QLabel('CRC校验:'), 3, 0)
        self.combo_crc = QComboBox()
        self.combo_crc.addItems(['无', 'CRC16-XMODEM', 'CRC16-CCITT', 'CRC16-MODBUS', '累加和', '异或'])
        layout.addWidget(self.combo_crc, 3, 1, 1, 3)
        
        # CRC位置说明
        crc_note = QLabel('注: CRC校验将验证除CRC字节外的所有数据字节')
        crc_note.setStyleSheet('color: #666; font-size: 8pt;')
        layout.addWidget(crc_note, 4, 0, 1, 4)
        
        # 启用状态
        self.check_enabled = QCheckBox('启用此协议')
        self.check_enabled.setChecked(True)
        layout.addWidget(self.check_enabled, 5, 0, 1, 4)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons, 6, 0, 1, 4)
        
        # 加载现有数据
        if protocol_data:
            self.load_data(protocol_data)
    
    def load_data(self, data):
        """加载协议数据到界面"""
        self.edit_name.setText(data.get('name', ''))
        header = data.get('header', [])
        self.edit_header.setText(' '.join(f'{b:02X}' for b in header))
        self.spin_length.setValue(data.get('length', 38))
        tail = data.get('tail', [])
        self.edit_tail.setText(' '.join(f'{b:02X}' for b in tail))
        self.combo_crc.setCurrentText(data.get('crc_type', '无'))
        self.check_enabled.setChecked(data.get('enabled', True))
    
    def get_data(self):
        """获取配置数据"""
        # 解析帧头
        header_text = self.edit_header.text().strip()
        header = []
        if header_text:
            try:
                header = [int(b, 16) for b in header_text.split()]
            except ValueError:
                QMessageBox.warning(self, '错误', '帧头格式错误,请使用十六进制格式(例如: A8 A8)')
                return None
        
        # 解析帧尾
        tail_text = self.edit_tail.text().strip()
        tail = []
        if tail_text:
            try:
                tail = [int(b, 16) for b in tail_text.split()]
            except ValueError:
                QMessageBox.warning(self, '错误', '帧尾格式错误,请使用十六进制格式(例如: AA AA)')
                return None
        
        name = self.edit_name.text().strip()
        if not name:
            QMessageBox.warning(self, '错误', '请输入协议名称')
            return None
        
        return {
            'name': name,
            'header': header,
            'length': self.spin_length.value(),
            'tail': tail,
            'crc_type': self.combo_crc.currentText(),
            'enabled': self.check_enabled.isChecked()
        }


class CustomDisplayDialog(QDialog):
    """自定义显示窗口配置对话框"""
    
    def __init__(self, parent=None, display_data=None, protocol_list=None):
        super().__init__(parent)
        self.setWindowTitle('配置显示窗口')
        self.setMinimumWidth(450)
        self.protocol_list = protocol_list or []
        
        layout = QGridLayout(self)
        
        # 窗口名称
        layout.addWidget(QLabel('窗口名称:'), 0, 0)
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText('例如: CO浓度值')
        layout.addWidget(self.edit_name, 0, 1, 1, 2)
        
        # 协议类型选择
        layout.addWidget(QLabel('数据协议:'), 0, 3)
        self.combo_protocol = QComboBox()
        self.combo_protocol.addItem('默认协议', None)
        for proto in self.protocol_list:
            if proto and proto.get('enabled', True):
                self.combo_protocol.addItem(proto.get('name', '未命名'), proto.get('name'))
        self.combo_protocol.setToolTip('选择数据来源协议')
        layout.addWidget(self.combo_protocol, 0, 4, 1, 2)
        
        # 起始字节
        layout.addWidget(QLabel('起始字节:'), 1, 0)
        self.spin_start_byte = QSpinBox()
        self.spin_start_byte.setRange(0, 100)
        self.spin_start_byte.setValue(2)
        self.spin_start_byte.setToolTip('数据帧中的起始字节位置(0-based)')
        layout.addWidget(self.spin_start_byte, 1, 1)
        
        # 字节数量
        layout.addWidget(QLabel('字节数量:'), 1, 2)
        self.spin_byte_count = QSpinBox()
        self.spin_byte_count.setRange(1, 8)
        self.spin_byte_count.setValue(2)
        self.spin_byte_count.setToolTip('读取的字节数(1-8)')
        layout.addWidget(self.spin_byte_count, 1, 3)
        
        # 数据类型
        layout.addWidget(QLabel('数据类型:'), 2, 0)
        self.combo_data_type = QComboBox()
        self.combo_data_type.addItems(['uint16 (LE)', 'uint16 (BE)', 'int16 (LE)', 'int16 (BE)',
                                       'uint32 (LE)', 'uint32 (BE)', 'int32 (LE)', 'int32 (BE)',
                                       'float (LE)', 'float (BE)', 'uint8', 'int8', 'string (ASCII)'])
        layout.addWidget(self.combo_data_type, 2, 1, 1, 3)
        
        # 乘系数
        layout.addWidget(QLabel('乘系数:'), 3, 0)
        self.spin_coefficient = QDoubleSpinBox()
        self.spin_coefficient.setRange(-1000000, 1000000)
        self.spin_coefficient.setValue(1.0)
        self.spin_coefficient.setSingleStep(0.1)
        self.spin_coefficient.setDecimals(4)
        self.spin_coefficient.setToolTip('乘以此系数')
        layout.addWidget(self.spin_coefficient, 3, 1)
        
        # 除系数
        layout.addWidget(QLabel('除系数:'), 3, 2)
        self.spin_divisor = QDoubleSpinBox()
        self.spin_divisor.setRange(0.0001, 1000000)
        self.spin_divisor.setValue(1.0)
        self.spin_divisor.setSingleStep(0.1)
        self.spin_divisor.setDecimals(4)
        self.spin_divisor.setToolTip('除以此系数')
        layout.addWidget(self.spin_divisor, 3, 3)
        
        # 偏移量
        layout.addWidget(QLabel('偏移量:'), 4, 0)
        self.spin_offset = QDoubleSpinBox()
        self.spin_offset.setRange(-1000000, 1000000)
        self.spin_offset.setValue(0.0)
        self.spin_offset.setSingleStep(0.1)
        self.spin_offset.setDecimals(4)
        self.spin_offset.setToolTip('加上此偏移量')
        layout.addWidget(self.spin_offset, 4, 1)
        
        # 单位
        layout.addWidget(QLabel('单位:'), 5, 0)
        self.edit_unit = QLineEdit()
        self.edit_unit.setPlaceholderText('例如: ppm, °C, %')
        layout.addWidget(self.edit_unit, 5, 1, 1, 3)
        
        # 小数位数
        layout.addWidget(QLabel('小数位数:'), 6, 0)
        self.spin_decimals = QSpinBox()
        self.spin_decimals.setRange(0, 6)
        self.spin_decimals.setValue(2)
        layout.addWidget(self.spin_decimals, 6, 1)
        
        # 启用/禁用
        self.check_enabled = QCheckBox('启用此窗口')
        self.check_enabled.setChecked(True)
        layout.addWidget(self.check_enabled, 7, 0, 1, 4)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons, 8, 0, 1, 4)
        
        # 加载数据
        if display_data:
            self.edit_name.setText(display_data.get('name', ''))
            
            # 设置协议
            protocol = display_data.get('protocol', None)
            for i in range(self.combo_protocol.count()):
                if self.combo_protocol.itemData(i) == protocol:
                    self.combo_protocol.setCurrentIndex(i)
                    break
            
            self.spin_start_byte.setValue(display_data.get('start_byte', 2))
            self.spin_byte_count.setValue(display_data.get('byte_count', 2))
            data_type = display_data.get('data_type', 'uint16 (LE)')
            index = self.combo_data_type.findText(data_type)
            if index >= 0:
                self.combo_data_type.setCurrentIndex(index)
            self.spin_coefficient.setValue(display_data.get('coefficient', 1.0))
            self.spin_divisor.setValue(display_data.get('divisor', 1.0))
            self.spin_offset.setValue(display_data.get('offset', 0.0))
            self.edit_unit.setText(display_data.get('unit', ''))
            self.spin_decimals.setValue(display_data.get('decimals', 2))
            self.check_enabled.setChecked(display_data.get('enabled', True))
    
    def get_data(self):
        return {
            'name': self.edit_name.text(),
            'protocol': self.combo_protocol.currentData(),
            'start_byte': self.spin_start_byte.value(),
            'byte_count': self.spin_byte_count.value(),
            'data_type': self.combo_data_type.currentText(),
            'coefficient': self.spin_coefficient.value(),
            'divisor': self.spin_divisor.value(),
            'offset': self.spin_offset.value(),
            'unit': self.edit_unit.text(),
            'decimals': self.spin_decimals.value(),
            'enabled': self.check_enabled.isChecked()
        }


class CurveConfigDialog(QDialog):
    """曲线配置对话框"""
    
    def __init__(self, parent=None, curve_data=None, protocol_list=None):
        super().__init__(parent)
        self.setWindowTitle('配置曲线')
        self.setMinimumWidth(550)
        self.protocol_list = protocol_list or []
        
        layout = QGridLayout(self)
        
        # 曲线名称
        layout.addWidget(QLabel('曲线名称:'), 0, 0)
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText('例如: CO浓度')
        layout.addWidget(self.edit_name, 0, 1, 1, 2)
        
        # 协议类型选择
        layout.addWidget(QLabel('数据协议:'), 0, 3)
        self.combo_protocol = QComboBox()
        self.combo_protocol.addItem('默认协议', None)
        for proto in self.protocol_list:
            if proto and proto.get('enabled', True):
                self.combo_protocol.addItem(proto.get('name', '未命名'), proto.get('name'))
        self.combo_protocol.setToolTip('选择数据来源协议')
        layout.addWidget(self.combo_protocol, 0, 4, 1, 2)
        
        # 起始字节
        label_start = QLabel('起始字节:')
        label_start.setMinimumWidth(80)
        layout.addWidget(label_start, 1, 0)
        self.spin_start_byte = QSpinBox()
        self.spin_start_byte.setRange(0, 100)
        self.spin_start_byte.setValue(2)
        self.spin_start_byte.setToolTip('数据帧中的起始字节位置（从0开始）')
        layout.addWidget(self.spin_start_byte, 1, 1)
        
        # 字节数
        label_bytes = QLabel('字节数:')
        label_bytes.setMinimumWidth(60)
        layout.addWidget(label_bytes, 1, 2)
        self.spin_byte_count = QSpinBox()
        self.spin_byte_count.setRange(1, 8)
        self.spin_byte_count.setValue(2)
        self.spin_byte_count.setToolTip('读取的字节数量')
        layout.addWidget(self.spin_byte_count, 1, 3)
        
        # 位模式复选框
        layout.addWidget(QLabel('位模式:'), 1, 4)
        self.check_bit_mode = QCheckBox('启用')
        self.check_bit_mode.setToolTip('启用后仅记录指定字节的指定位(0/1)')
        self.check_bit_mode.toggled.connect(self.on_bit_mode_toggled)
        layout.addWidget(self.check_bit_mode, 1, 5)
        
        # 数据类型
        layout.addWidget(QLabel('数据类型:'), 2, 0)
        self.combo_data_type = QComboBox()
        self.combo_data_type.addItems([
            'uint8', 'int8',
            'uint16 (LE)', 'uint16 (BE)', 'int16 (LE)', 'int16 (BE)',
            'uint32 (LE)', 'uint32 (BE)', 'int32 (LE)', 'int32 (BE)',
            'float (LE)', 'float (BE)', 'double (LE)', 'double (BE)'
        ])
        self.combo_data_type.setCurrentText('uint16 (LE)')
        layout.addWidget(self.combo_data_type, 2, 1, 1, 3)
        
        # 位选择器（仅在位模式下显示）
        layout.addWidget(QLabel('目标位:'), 2, 4)
        self.spin_bit_index = QSpinBox()
        self.spin_bit_index.setRange(0, 7)
        self.spin_bit_index.setValue(0)
        self.spin_bit_index.setToolTip('要记录的位索引 (0-7, 0为最低位LSB)')
        self.spin_bit_index.setEnabled(False)
        layout.addWidget(self.spin_bit_index, 2, 5)
        
        # 乘系数
        label_coef = QLabel('乘系数:')
        label_coef.setMinimumWidth(80)
        layout.addWidget(label_coef, 3, 0)
        self.spin_coefficient = QDoubleSpinBox()
        self.spin_coefficient.setRange(0.0001, 1000000)
        self.spin_coefficient.setValue(1.0)
        self.spin_coefficient.setSingleStep(0.1)
        self.spin_coefficient.setDecimals(4)
        self.spin_coefficient.setToolTip('乘以此系数')
        layout.addWidget(self.spin_coefficient, 3, 1)
        
        # 除系数
        label_div = QLabel('除系数:')
        label_div.setMinimumWidth(60)
        layout.addWidget(label_div, 3, 2)
        self.spin_divisor = QDoubleSpinBox()
        self.spin_divisor.setRange(0.0001, 1000000)
        self.spin_divisor.setValue(1.0)
        self.spin_divisor.setSingleStep(0.1)
        self.spin_divisor.setDecimals(4)
        self.spin_divisor.setToolTip('除以此系数')
        layout.addWidget(self.spin_divisor, 3, 3)
        
        # 偏移量
        layout.addWidget(QLabel('偏移量:'), 4, 0)
        self.spin_offset = QDoubleSpinBox()
        self.spin_offset.setRange(-1000000, 1000000)
        self.spin_offset.setValue(0.0)
        self.spin_offset.setSingleStep(0.1)
        self.spin_offset.setDecimals(4)
        self.spin_offset.setToolTip('加上此偏移量')
        layout.addWidget(self.spin_offset, 4, 1)
        
        # 单位
        layout.addWidget(QLabel('单位:'), 4, 2)
        self.edit_unit = QLineEdit()
        self.edit_unit.setPlaceholderText('例如: ppm, °C')
        layout.addWidget(self.edit_unit, 4, 3)
        
        # 颜色选择
        layout.addWidget(QLabel('曲线颜色:'), 5, 0)
        self.combo_color = QComboBox()
        self.combo_color.addItems(['蓝色', '绿色', '红色', '橙色', '紫色', '青色', '品红'])
        layout.addWidget(self.combo_color, 5, 1, 1, 3)
        
        # 启用/禁用
        self.check_enabled = QCheckBox('启用此曲线')
        self.check_enabled.setChecked(True)
        self.check_enabled.setToolTip('启用后曲线会显示在主界面和图表中')
        layout.addWidget(self.check_enabled, 6, 0, 1, 2)
        
        # 记录选项
        self.check_record = QCheckBox('记录此曲线')
        self.check_record.setChecked(True)
        self.check_record.setToolTip('勾选后此曲线数据将被导出到Excel文件')
        layout.addWidget(self.check_record, 6, 2, 1, 2)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons, 7, 0, 1, 4)
        
        # 加载数据
        if curve_data:
            self.edit_name.setText(curve_data.get('name', ''))
            
            # 设置协议
            protocol = curve_data.get('protocol', None)
            for i in range(self.combo_protocol.count()):
                if self.combo_protocol.itemData(i) == protocol:
                    self.combo_protocol.setCurrentIndex(i)
                    break
            
            self.spin_start_byte.setValue(curve_data.get('start_byte', 2))
            self.spin_byte_count.setValue(curve_data.get('byte_count', 2))
            data_type = curve_data.get('data_type', 'uint16 (LE)')
            index = self.combo_data_type.findText(data_type)
            if index >= 0:
                self.combo_data_type.setCurrentIndex(index)
            self.spin_coefficient.setValue(curve_data.get('coefficient', 1.0))
            self.spin_divisor.setValue(curve_data.get('divisor', 1.0))
            self.spin_offset.setValue(curve_data.get('offset', 0.0))
            self.edit_unit.setText(curve_data.get('unit', ''))
            color = curve_data.get('color', '蓝色')
            index = self.combo_color.findText(color)
            if index >= 0:
                self.combo_color.setCurrentIndex(index)
            self.check_enabled.setChecked(curve_data.get('enabled', True))
            
            # 位模式相关配置
            bit_mode = curve_data.get('bit_mode', False)
            self.check_bit_mode.setChecked(bit_mode)
            if 'bit_index' in curve_data:
                self.spin_bit_index.setValue(curve_data.get('bit_index', 0))
            
            # 记录选项
            self.check_record.setChecked(curve_data.get('record', True))
    
    def on_bit_mode_toggled(self, checked):
        """位模式切换时的处理"""
        self.spin_bit_index.setEnabled(checked)
        self.combo_data_type.setEnabled(not checked)
        self.spin_byte_count.setEnabled(not checked)
        if checked:
            # 启用位模式时，强制设置为1字节
            self.spin_byte_count.setValue(1)
    
    def get_data(self):
        return {
            'name': self.edit_name.text(),
            'protocol': self.combo_protocol.currentData(),
            'start_byte': self.spin_start_byte.value(),
            'byte_count': self.spin_byte_count.value(),
            'data_type': self.combo_data_type.currentText(),
            'coefficient': self.spin_coefficient.value(),
            'divisor': self.spin_divisor.value(),
            'offset': self.spin_offset.value(),
            'unit': self.edit_unit.text(),
            'color': self.combo_color.currentText(),
            'enabled': self.check_enabled.isChecked(),
            'record': self.check_record.isChecked(),
            'bit_mode': self.check_bit_mode.isChecked(),
            'bit_index': self.spin_bit_index.value()
        }


class ClockConfigDialog(QDialog):
    """接收时钟配置对话框"""
    
    def __init__(self, parent=None, clock_data=None, protocol_list=None):
        super().__init__(parent)
        self.setWindowTitle('配置接收时钟')
        self.setMinimumWidth(550)
        self.protocol_list = protocol_list or []
        
        layout = QGridLayout(self)
        
        # 协议类型选择
        layout.addWidget(QLabel('数据协议:'), 0, 0)
        self.combo_protocol = QComboBox()
        self.combo_protocol.addItem('默认协议', None)
        for proto in self.protocol_list:
            if proto and proto.get('enabled', True):
                self.combo_protocol.addItem(proto.get('name', '未命名'), proto.get('name'))
        self.combo_protocol.setToolTip('选择数据来源协议')
        layout.addWidget(self.combo_protocol, 0, 1, 1, 5)
        
        # 年配置
        layout.addWidget(QLabel('年份起始字节:'), 1, 0)
        self.spin_year_start = QSpinBox()
        self.spin_year_start.setRange(0, 100)
        self.spin_year_start.setValue(10)
        layout.addWidget(self.spin_year_start, 1, 1)
        
        layout.addWidget(QLabel('字节数:'), 1, 2)
        self.spin_year_count = QSpinBox()
        self.spin_year_count.setRange(1, 4)
        self.spin_year_count.setValue(2)
        layout.addWidget(self.spin_year_count, 1, 3)
        
        layout.addWidget(QLabel('数据类型:'), 1, 4)
        self.combo_year_type = QComboBox()
        self.combo_year_type.addItems(['uint16 (LE)', 'uint16 (BE)', 'uint8'])
        layout.addWidget(self.combo_year_type, 1, 5)
        
        # 月配置
        layout.addWidget(QLabel('月份起始字节:'), 2, 0)
        self.spin_month_start = QSpinBox()
        self.spin_month_start.setRange(0, 100)
        self.spin_month_start.setValue(12)
        layout.addWidget(self.spin_month_start, 2, 1)
        
        layout.addWidget(QLabel('数据类型:'), 2, 2)
        self.combo_month_type = QComboBox()
        self.combo_month_type.addItems(['uint8'])
        layout.addWidget(self.combo_month_type, 2, 3, 1, 3)
        
        # 日配置
        layout.addWidget(QLabel('日期起始字节:'), 3, 0)
        self.spin_day_start = QSpinBox()
        self.spin_day_start.setRange(0, 100)
        self.spin_day_start.setValue(13)
        layout.addWidget(self.spin_day_start, 3, 1)
        
        layout.addWidget(QLabel('数据类型:'), 3, 2)
        self.combo_day_type = QComboBox()
        self.combo_day_type.addItems(['uint8'])
        layout.addWidget(self.combo_day_type, 3, 3, 1, 3)
        
        # 时配置
        layout.addWidget(QLabel('小时起始字节:'), 4, 0)
        self.spin_hour_start = QSpinBox()
        self.spin_hour_start.setRange(0, 100)
        self.spin_hour_start.setValue(14)
        layout.addWidget(self.spin_hour_start, 4, 1)
        
        layout.addWidget(QLabel('数据类型:'), 4, 2)
        self.combo_hour_type = QComboBox()
        self.combo_hour_type.addItems(['uint8'])
        layout.addWidget(self.combo_hour_type, 4, 3, 1, 3)
        
        # 分配置
        layout.addWidget(QLabel('分钟起始字节:'), 5, 0)
        self.spin_minute_start = QSpinBox()
        self.spin_minute_start.setRange(0, 100)
        self.spin_minute_start.setValue(15)
        layout.addWidget(self.spin_minute_start, 5, 1)
        
        layout.addWidget(QLabel('数据类型:'), 5, 2)
        self.combo_minute_type = QComboBox()
        self.combo_minute_type.addItems(['uint8'])
        layout.addWidget(self.combo_minute_type, 5, 3, 1, 3)
        
        # 秒配置
        layout.addWidget(QLabel('秒数起始字节:'), 6, 0)
        self.spin_second_start = QSpinBox()
        self.spin_second_start.setRange(0, 100)
        self.spin_second_start.setValue(16)
        layout.addWidget(self.spin_second_start, 6, 1)
        
        layout.addWidget(QLabel('数据类型:'), 6, 2)
        self.combo_second_type = QComboBox()
        self.combo_second_type.addItems(['uint8'])
        layout.addWidget(self.combo_second_type, 6, 3, 1, 3)
        
        # 启用复选框
        self.check_enabled = QCheckBox('启用接收时钟')
        self.check_enabled.setChecked(True)
        layout.addWidget(self.check_enabled, 7, 0, 1, 6)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons, 8, 0, 1, 6)
        
        # 加载数据
        if clock_data:
            # 设置协议
            protocol = clock_data.get('protocol', None)
            for i in range(self.combo_protocol.count()):
                if self.combo_protocol.itemData(i) == protocol:
                    self.combo_protocol.setCurrentIndex(i)
                    break
            
            self.spin_year_start.setValue(clock_data.get('year_start', 10))
            self.spin_year_count.setValue(clock_data.get('year_count', 2))
            self.combo_year_type.setCurrentText(clock_data.get('year_type', 'uint16 (LE)'))
            self.spin_month_start.setValue(clock_data.get('month_start', 12))
            self.spin_day_start.setValue(clock_data.get('day_start', 13))
            self.spin_hour_start.setValue(clock_data.get('hour_start', 14))
            self.spin_minute_start.setValue(clock_data.get('minute_start', 15))
            self.spin_second_start.setValue(clock_data.get('second_start', 16))
            self.check_enabled.setChecked(clock_data.get('enabled', True))
    
    def get_data(self):
        return {
            'protocol': self.combo_protocol.currentData(),
            'year_start': self.spin_year_start.value(),
            'year_count': self.spin_year_count.value(),
            'year_type': self.combo_year_type.currentText(),
            'month_start': self.spin_month_start.value(),
            'day_start': self.spin_day_start.value(),
            'hour_start': self.spin_hour_start.value(),
            'minute_start': self.spin_minute_start.value(),
            'second_start': self.spin_second_start.value(),
            'enabled': self.check_enabled.isChecked()
        }


class PresetCommandDialog(QDialog):
    """预设命令配置对话框"""
    
    def __init__(self, parent=None, preset_data=None):
        super().__init__(parent)
        self.setWindowTitle('配置预设命令')
        self.setMinimumWidth(550)
        
        layout = QGridLayout(self)
        
        # 命令名称
        layout.addWidget(QLabel('命令名称:'), 0, 0)
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText('例如: 查询版本')
        layout.addWidget(self.edit_name, 0, 1, 1, 2)
        
        # 命令内容
        layout.addWidget(QLabel('命令内容:'), 1, 0)
        self.edit_command = QLineEdit()
        self.edit_command.setPlaceholderText('例如: AT+VER 或 A8 A8 01 02')
        layout.addWidget(self.edit_command, 1, 1, 1, 2)
        
        # HEX模式
        self.check_hex = QCheckBox('HEX格式')
        layout.addWidget(self.check_hex, 2, 0)
        
        # 添加时间戳
        self.check_timestamp = QCheckBox('添加时间戳')
        layout.addWidget(self.check_timestamp, 2, 1)
        
        # CRC校验
        layout.addWidget(QLabel('CRC校验:'), 3, 0)
        self.combo_crc = QComboBox()
        self.combo_crc.addItems(['无', 'CCITT-CRC16', 'Modbus-CRC16', 'CRC16-XMODEM', '累加和', '异或'])
        layout.addWidget(self.combo_crc, 3, 1, 1, 2)
        
        # 周期发送
        self.check_periodic = QCheckBox('周期发送')
        self.check_periodic.toggled.connect(self.on_periodic_toggled)
        layout.addWidget(self.check_periodic, 4, 0)
        
        # 发送周期
        layout.addWidget(QLabel('发送周期(秒):'), 4, 1)
        self.spin_period = QDoubleSpinBox()
        self.spin_period.setRange(0.001, 3600.0)  # 最小1毫秒
        self.spin_period.setValue(1.0)
        self.spin_period.setSingleStep(0.01)  # 单步10毫秒
        self.spin_period.setDecimals(3)  # 精度到毫秒
        self.spin_period.setEnabled(False)
        layout.addWidget(self.spin_period, 4, 2)
        
        # 数据填充功能
        self.check_data_fill = QCheckBox('启用数据填充')
        self.check_data_fill.setToolTip('启用后，发送时会用填充框的数据替换第2-3字节')
        self.check_data_fill.toggled.connect(self.on_data_fill_toggled)
        layout.addWidget(self.check_data_fill, 5, 0)
        
        # 数据范围
        layout.addWidget(QLabel('数据范围:'), 5, 1)
        range_layout = QHBoxLayout()
        self.spin_data_min = QSpinBox()
        self.spin_data_min.setRange(0, 65535)
        self.spin_data_min.setValue(0)
        self.spin_data_min.setEnabled(False)
        self.spin_data_min.setPrefix('最小: ')
        range_layout.addWidget(self.spin_data_min)
        self.spin_data_max = QSpinBox()
        self.spin_data_max.setRange(0, 65535)
        self.spin_data_max.setValue(65535)
        self.spin_data_max.setEnabled(False)
        self.spin_data_max.setPrefix('最大: ')
        range_layout.addWidget(self.spin_data_max)
        layout.addLayout(range_layout, 5, 2)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons, 6, 0, 1, 3)
        
        # 加载数据
        if preset_data:
            self.edit_name.setText(preset_data.get('name', ''))
            self.edit_command.setText(preset_data.get('command', ''))
            self.check_hex.setChecked(preset_data.get('is_hex', False))
            self.check_timestamp.setChecked(preset_data.get('add_timestamp', False))
            crc_type = preset_data.get('crc_type', '无')
            index = self.combo_crc.findText(crc_type)
            if index >= 0:
                self.combo_crc.setCurrentIndex(index)
            self.check_periodic.setChecked(preset_data.get('periodic', False))
            self.spin_period.setValue(preset_data.get('period', 1.0))
            self.check_data_fill.setChecked(preset_data.get('data_fill_enabled', False))
            self.spin_data_min.setValue(preset_data.get('data_min', 0))
            self.spin_data_max.setValue(preset_data.get('data_max', 65535))
    
    def on_periodic_toggled(self, checked):
        self.spin_period.setEnabled(checked)
    
    def on_data_fill_toggled(self, checked):
        self.spin_data_min.setEnabled(checked)
        self.spin_data_max.setEnabled(checked)
    
    def get_data(self):
        return {
            'name': self.edit_name.text(),
            'command': self.edit_command.text(),
            'is_hex': self.check_hex.isChecked(),
            'add_timestamp': self.check_timestamp.isChecked(),
            'crc_type': self.combo_crc.currentText(),
            'periodic': self.check_periodic.isChecked(),
            'period': self.spin_period.value(),
            'data_fill_enabled': self.check_data_fill.isChecked(),
            'data_min': self.spin_data_min.value(),
            'data_max': self.spin_data_max.value()
        }


class BitDisplayDialog(QDialog):
    """位(bit)显示窗口配置对话框"""
    
    def __init__(self, parent=None, bit_display_data=None, protocol_list=None):
        super().__init__(parent)
        self.setWindowTitle('配置位显示窗口')
        self.setMinimumWidth(600)
        self.setMinimumHeight(500)
        self.protocol_list = protocol_list or []
        
        layout = QVBoxLayout(self)
        
        # 基本配置区域
        basic_group = QGroupBox('基本配置')
        basic_layout = QGridLayout()
        
        # 窗口名称
        basic_layout.addWidget(QLabel('窗口名称:'), 0, 0)
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText('例如: 状态指示')
        basic_layout.addWidget(self.edit_name, 0, 1, 1, 2)
        
        # 协议类型选择
        basic_layout.addWidget(QLabel('数据协议:'), 1, 0)
        self.combo_protocol = QComboBox()
        self.combo_protocol.addItem('默认协议', None)
        for proto in self.protocol_list:
            if proto and proto.get('enabled', True):
                self.combo_protocol.addItem(proto.get('name', '未命名'), proto.get('name'))
        self.combo_protocol.setToolTip('选择数据来源协议')
        basic_layout.addWidget(self.combo_protocol, 1, 1, 1, 2)
        
        # 目标字节
        basic_layout.addWidget(QLabel('目标字节:'), 2, 0)
        self.spin_target_byte = QSpinBox()
        self.spin_target_byte.setRange(0, 100)
        self.spin_target_byte.setValue(17)
        self.spin_target_byte.setToolTip('要解析的字节在数据帧中的位置')
        basic_layout.addWidget(self.spin_target_byte, 2, 1, 1, 2)
        
        basic_group.setLayout(basic_layout)
        layout.addWidget(basic_group)
        
        # 位配置区域
        bits_group = QGroupBox('位(Bit)配置 - 设置每个位对应的图标名称')
        bits_layout = QGridLayout()
        bits_layout.setSpacing(6)
        
        self.bit_name_edits = []
        for i in range(8):
            # Bit标签
            bit_label = QLabel(f'Bit {i}:')
            bit_label.setStyleSheet('font-weight: bold;')
            bits_layout.addWidget(bit_label, i, 0)
            
            # 名称输入框
            name_edit = QLineEdit()
            name_edit.setPlaceholderText(f'例如: 开关{i+1}')
            bits_layout.addWidget(name_edit, i, 1)
            
            # 状态预览标签
            status_label = QLabel('●')
            status_label.setStyleSheet('color: #ccc; font-size: 18pt;')
            status_label.setAlignment(Qt.AlignCenter)
            status_label.setMinimumWidth(30)
            bits_layout.addWidget(status_label, i, 2)
            
            self.bit_name_edits.append(name_edit)
        
        bits_group.setLayout(bits_layout)
        layout.addWidget(bits_group)
        
        # 启用复选框
        self.check_enabled = QCheckBox('启用此窗口')
        self.check_enabled.setChecked(True)
        layout.addWidget(self.check_enabled)
        
        # 提示信息
        tip_label = QLabel('提示: Bit 0是最低位(LSB)，Bit 7是最高位(MSB)')
        tip_label.setStyleSheet('color: #666; font-size: 9pt; font-style: italic;')
        layout.addWidget(tip_label)
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        # 加载数据
        if bit_display_data:
            self.edit_name.setText(bit_display_data.get('name', ''))
            
            # 设置协议
            protocol = bit_display_data.get('protocol', None)
            for i in range(self.combo_protocol.count()):
                if self.combo_protocol.itemData(i) == protocol:
                    self.combo_protocol.setCurrentIndex(i)
                    break
            
            self.spin_target_byte.setValue(bit_display_data.get('target_byte', 17))
            bit_names = bit_display_data.get('bit_names', [''] * 8)
            for i, name in enumerate(bit_names[:8]):
                self.bit_name_edits[i].setText(name)
            self.check_enabled.setChecked(bit_display_data.get('enabled', True))
    
    def get_data(self):
        bit_names = [edit.text() for edit in self.bit_name_edits]
        return {
            'name': self.edit_name.text(),
            'protocol': self.combo_protocol.currentData(),
            'target_byte': self.spin_target_byte.value(),
            'bit_names': bit_names,
            'enabled': self.check_enabled.isChecked()
        }


class SerialParamsDialog(QDialog):
    """串口参数配置对话框"""
    
    def __init__(self, parent=None, current_params=None):
        super().__init__(parent)
        self.setWindowTitle('串口参数配置')
        self.setMinimumWidth(350)
        
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        # 参数设置组
        params_group = QGroupBox('串口参数')
        params_layout = QGridLayout()
        params_layout.setSpacing(8)
        
        # 波特率
        params_layout.addWidget(QLabel("波特率:"), 0, 0)
        self.combo_baudrate = QComboBox()
        self.combo_baudrate.setMinimumHeight(28)
        self.combo_baudrate.addItems(['2400', '4800', '9600', '14400', '19200', '38400', 
                                      '57600', '115200', '128000', '256000', '460800', '921600'])
        self.combo_baudrate.setCurrentText('115200')
        params_layout.addWidget(self.combo_baudrate, 0, 1)
        
        # 数据位
        params_layout.addWidget(QLabel("数据位:"), 1, 0)
        self.combo_databits = QComboBox()
        self.combo_databits.setMinimumHeight(28)
        self.combo_databits.addItems(['5', '6', '7', '8'])
        self.combo_databits.setCurrentText('8')
        params_layout.addWidget(self.combo_databits, 1, 1)
        
        # 停止位
        params_layout.addWidget(QLabel("停止位:"), 2, 0)
        self.combo_stopbits = QComboBox()
        self.combo_stopbits.setMinimumHeight(28)
        self.combo_stopbits.addItems(['1', '1.5', '2'])
        self.combo_stopbits.setCurrentText('1')
        params_layout.addWidget(self.combo_stopbits, 2, 1)
        
        # 校验位
        params_layout.addWidget(QLabel("校验位:"), 3, 0)
        self.combo_parity = QComboBox()
        self.combo_parity.setMinimumHeight(28)
        self.combo_parity.addItems(['None', 'Even', 'Odd', 'Mark', 'Space'])
        self.combo_parity.setCurrentText('None')
        params_layout.addWidget(self.combo_parity, 3, 1)
        
        params_group.setLayout(params_layout)
        layout.addWidget(params_group)
        
        # 加载当前参数
        if current_params:
            self.combo_baudrate.setCurrentText(str(current_params.get('baudrate', '115200')))
            self.combo_databits.setCurrentText(str(current_params.get('databits', '8')))
            self.combo_stopbits.setCurrentText(str(current_params.get('stopbits', '1')))
            self.combo_parity.setCurrentText(str(current_params.get('parity', 'None')))
        
        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
    
    def get_params(self):
        """获取配置参数"""
        return {
            'baudrate': self.combo_baudrate.currentText(),
            'databits': self.combo_databits.currentText(),
            'stopbits': self.combo_stopbits.currentText(),
            'parity': self.combo_parity.currentText()
        }


class SerialMonitorApp(QMainWindow):
    """串口上位机主窗口"""
    
    CONFIG_FILE = "serial_config.json"
    LAYOUT_FILE = "window_layout.json"  # 窗口布局配置文件
    
    def __init__(self, splash=None):
        super().__init__()
        self.splash = splash
        self.serial_port = None
        self.serial_thread = None
        self.parser = DataParser()
        self.raw_data_buffer = []
        
        # 预设命令相关
        self.preset_commands = []
        self.preset_buttons = []
        self.preset_timers = []
        self.preset_data_spinboxes = []  # 数据填充输入框
        
        # 自定义显示窗口相关
        self.custom_displays = [None] * 50  # 支持50个显示窗口
        self.custom_display_labels = []
        self.display_widgets = []  # 保存显示窗口的容器
        
        # 位显示窗口相关
        self.bit_displays = [None] * 50  # 支持50个位显示窗口
        self.bit_display_labels = []  # 保存每个位的显示标签
        self.bit_display_widgets = []  # 保存每个位显示窗口的容器
        self.bit_display_name_labels = []  # 保存每个位显示窗口的名称标签
        
        # 时钟相关
        self.clock_config = None  # 接收时钟配置
        self.received_time = None  # 接收到的时间
        
        # 窗口布局恢复标志
        self.is_restoring_layout = False
        
        # 帧间隔时间设置（毫秒）
        self.frame_interval_ms = 100  # 默认100ms
        self.last_rx_time = None  # 上次接收数据时间
        self.rx_buffer = bytearray()  # 接收数据缓冲
        
        # 曲线记录状态
        self.is_curve_recording = False
        
        # 智能恢复功能相关
        self.lost_port_info = None  # 记录断开的串口信息
        self.lost_port_config = None  # 记录断开时的配置
        self.recovery_check_timer = None  # 恢复检测定时器
        self.recovery_start_time = None  # 开始检测时间
        self.port_stable_time = None  # 端口稳定时间
        self.last_port_state = False  # 上次端口状态
        
        # 背景相关
        self.background_image_path = ''  # 背景图片路径
        self.background_opacity = 0.3  # 背景透明度(0.0-1.0)
        
        # 协议管理相关
        self.protocol_configs = []  # 协议配置列表
        self.protocol_parsers = []  # 协议解析器列表
        
        # 【重要】先创建所有定时器，确保在setup_dock_signals()之前就存在
        # 定时器 - 用于更新曲线(1Hz)
        self.plot_timer = QTimer()
        self.plot_timer.timeout.connect(self.update_plot)
        
        # 定时器 - 用于刷新COM口列表
        self.port_refresh_timer = QTimer()
        self.port_refresh_timer.timeout.connect(self.refresh_com_ports)
        
        # 定时器 - 用于延迟保存窗口布局（单次触发）- 必须在setup_dock_signals之前创建
        self.layout_save_timer = QTimer()
        self.layout_save_timer.setSingleShot(True)  # 单次触发
        self.layout_save_timer.timeout.connect(self.save_window_layout)
        
        # 定时器 - 用于刷新接收缓冲区(检查帧超时)
        self.rx_flush_timer = QTimer()
        self.rx_flush_timer.timeout.connect(self.flush_rx_buffer)
        
        self.init_ui()
        
        if self.splash:
            self.splash.update_progress(70, '正在连接窗口信号...')
        
        # 先连接dock窗口信号（在加载配置前，但在layout_save_timer创建后）
        self.setup_dock_signals()
        
        if self.splash:
            self.splash.update_progress(80, '正在加载配置...')
        
        # 加载配置并恢复布局
        self.load_config(restore_layout=True)
        
        # 在加载配置后重新更新背景（此时背景路径已从配置文件加载）
        self.update_background()
        
        if self.splash:
            self.splash.update_progress(90, '正在初始化完成...')
        
        # 启动定时器
        self.port_refresh_timer.start(2000)
        self.rx_flush_timer.start(50)  # 每50ms检查一次
    
    def init_ui(self):
        """初始化界面"""
        if self.splash:
            self.splash.update_progress(10, '正在初始化主窗口...')
        
        self.setWindowTitle('力华亘金-上位机监控系统') # 设置窗口标题
        self.setGeometry(100, 100, 1500, 900)
        
        # 设置窗口图标
        icon_files = ['lihua_logo.png', 'icon.png', 'logo.ico', 'app.ico']
        from PyQt5.QtGui import QIcon
        for icon_file in icon_files:
            if os.path.exists(icon_file):
                self.setWindowIcon(QIcon(icon_file))
                break
        
        # 启用QDockWidget嵌套和分组功能
        self.setDockNestingEnabled(True)
        
        # 设置角落区域分配，优化布局
        self.setCorner(Qt.TopLeftCorner, Qt.LeftDockWidgetArea)
        self.setCorner(Qt.TopRightCorner, Qt.RightDockWidgetArea)
        self.setCorner(Qt.BottomLeftCorner, Qt.LeftDockWidgetArea)
        self.setCorner(Qt.BottomRightCorner, Qt.RightDockWidgetArea)
        
        # 设置背景
        self.update_background()
        
        if self.splash:
            self.splash.update_progress(20, '正在加载样式...')
        
        # 设置样式表 - Apple毛玻璃风格 (白色玻璃质感 - 终极优化版)
        # 注意：QMainWindow 不设置背景，让 QPalette 控制（用于背景图片功能）
        self.setStyleSheet("""
            /* 全局字体与颜色 - 优化抗锯齿 */
            QWidget {
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
                color: #1d1d1f;
                outline: none;
            }
            
            /* ToolTip - 现代深色风格 */
            QToolTip {
                background-color: rgba(30, 30, 30, 0.9);
                color: #fff;
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 6px;
                padding: 6px;
                font-size: 9pt;
            }
            
            QMainWindow {
                /* 背景由 QPalette 控制 */
            }
            
            /* 滚动区域透明，透出背景图 */
            QScrollArea, QScrollArea > QWidget > QWidget {
                background-color: transparent;
                border: none;
            }

            /* QGroupBox - 玻璃拟态 (Glassmorphism) 增强 */
            QGroupBox {
                background-color: rgba(255, 255, 255, 0.6);
                border: 1px solid rgba(255, 255, 255, 0.8);
                border-bottom: 1px solid rgba(0, 0, 0, 0.08);
                border-right: 1px solid rgba(0, 0, 0, 0.04);
                border-radius: 16px;
                margin-top: 12px;
                padding: 20px 12px 12px 12px;
                font-weight: 600;
                font-size: 10pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                background-color: transparent;
                color: #1d1d1f;
                left: 12px;
                top: 2px;
            }
            
            /* QPushButton - 增强立体质感与微交互 */
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                      stop:0 rgba(255, 255, 255, 0.85),
                                      stop:1 rgba(242, 242, 247, 0.75));
                color: #1d1d1f;
                border: 1px solid rgba(0, 0, 0, 0.08);
                border-bottom: 1px solid rgba(0, 0, 0, 0.15); /* 底部更深，模拟立体感 */
                border-radius: 8px;
                padding: 6px 16px;
                font-size: 9pt;
                font-weight: 600;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                      stop:0 rgba(255, 255, 255, 0.95),
                                      stop:1 rgba(250, 250, 255, 0.9));
                border-color: rgba(0, 0, 0, 0.12);
                border-bottom-color: rgba(0, 0, 0, 0.18);
            }
            QPushButton:pressed {
                background-color: rgba(230, 230, 235, 0.9);
                border: 1px solid rgba(0, 0, 0, 0.1);
                border-top: 1px solid rgba(0, 0, 0, 0.15); /* 按下时顶部阴影 */
                padding-top: 7px;
                padding-bottom: 5px;
            }
            QPushButton:disabled {
                background-color: rgba(255, 255, 255, 0.3);
                color: rgba(29, 29, 31, 0.3);
                border: 1px solid rgba(0, 0, 0, 0.05);
            }
            
            /* 输入控件 - 优化光标与边框 */
            QTextEdit, QLineEdit {
                background-color: rgba(255, 255, 255, 0.5);
                border: 1px solid rgba(0, 0, 0, 0.06);
                border-bottom: 1px solid rgba(0, 0, 0, 0.12);
                border-radius: 8px;
                padding: 6px;
                selection-background-color: rgba(0, 122, 255, 0.3);
                selection-color: #000;
            }
            QTextEdit:focus, QLineEdit:focus {
                background-color: rgba(255, 255, 255, 0.95);
                border: 1px solid #007AFF;
                border-bottom: 1px solid #007AFF;
            }
            
            /* 下拉框 - 增加渐变与阴影 */
            QComboBox {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                      stop:0 rgba(255, 255, 255, 0.7),
                                      stop:1 rgba(245, 245, 250, 0.6));
                border: 1px solid rgba(0, 0, 0, 0.06);
                border-bottom: 1px solid rgba(0, 0, 0, 0.12);
                border-radius: 8px;
                padding: 4px 10px;
                min-height: 24px;
            }
            QComboBox:hover {
                background-color: rgba(255, 255, 255, 0.9);
                border-color: rgba(0, 0, 0, 0.1);
            }
            QComboBox::drop-down {
                border: none;
                width: 24px;
            }
            QComboBox::down-arrow {
                image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><path d="M2 3 L5 6 L8 3" fill="none" stroke="#333" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>');
                width: 12px;
                height: 12px;
            }
            QComboBox QAbstractItemView {
                background-color: rgba(255, 255, 255, 0.95);
                border: 1px solid rgba(0, 0, 0, 0.08);
                border-radius: 8px;
                outline: none;
                padding: 4px;
                selection-background-color: #007AFF;
                selection-color: white;
            }
            
            /* 数字输入框 */
            QSpinBox, QDoubleSpinBox {
                background-color: rgba(255, 255, 255, 0.5);
                border: 1px solid rgba(0, 0, 0, 0.06);
                border-bottom: 1px solid rgba(0, 0, 0, 0.12);
                border-radius: 8px;
                padding: 4px;
                padding-right: 15px;
            }
            QSpinBox:focus, QDoubleSpinBox:focus {
                background-color: rgba(255, 255, 255, 0.95);
                border: 1px solid #007AFF;
            }
            QSpinBox::up-button, QDoubleSpinBox::up-button,
            QSpinBox::down-button, QDoubleSpinBox::down-button {
                background-color: rgba(0, 0, 0, 0.03);
                border: none;
                width: 18px;
                margin: 1px;
                border-radius: 4px;
            }
            QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
            QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {
                background-color: rgba(0, 0, 0, 0.1);
            }
            
            QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {
                image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><path d="M2 7 L5 4 L8 7" fill="none" stroke="#000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>');
                width: 10px;
                height: 10px;
            }
            QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {
                image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><path d="M2 3 L5 6 L8 3" fill="none" stroke="#000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>');
                width: 10px;
                height: 10px;
            }
            
            /* 复选框 - 优化选中态 */
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 5px;
                border: 1px solid rgba(0, 0, 0, 0.15);
                background-color: rgba(255, 255, 255, 0.8);
            }
            QCheckBox::indicator:checked {
                background-color: #007AFF;
                border-color: #007AFF;
                image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12"><path d="M2 6 L5 9 L10 3" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>');
            }
            QCheckBox::indicator:hover {
                border-color: #007AFF;
                background-color: #fff;
            }
            
            /* Dock Widget - 标题栏优化 */
            QDockWidget {
                titlebar-close-icon: url(close.png);
                titlebar-normal-icon: url(float.png);
            }
            QDockWidget::title {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                      stop:0 rgba(255, 255, 255, 0.8),
                                      stop:1 rgba(245, 245, 250, 0.8));
                border-bottom: 1px solid rgba(0, 0, 0, 0.08);
                padding: 8px 12px;
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
                font-weight: 600;
            }
            QDockWidget::close-button, QDockWidget::float-button {
                background: transparent;
                border-radius: 4px;
                padding: 2px;
            }
            QDockWidget::close-button:hover, QDockWidget::float-button:hover {
                background: rgba(0, 0, 0, 0.1);
            }
            
            /* 分割器手柄 */
            QSplitter::handle {
                background-color: rgba(0, 0, 0, 0.03);
                margin: 1px;
            }
            QSplitter::handle:hover {
                background-color: rgba(0, 122, 255, 0.2);
            }
            
            /* 滚动条 - iOS 风格悬浮滚动条 */
            QScrollBar:vertical {
                background: transparent;
                width: 10px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background: rgba(0, 0, 0, 0.15);
                border-radius: 5px;
                min-height: 40px;
                margin: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background: rgba(0, 0, 0, 0.35);
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            QScrollBar:horizontal {
                background: transparent;
                height: 10px;
                margin: 0;
            }
            QScrollBar::handle:horizontal {
                background: rgba(0, 0, 0, 0.15);
                border-radius: 5px;
                min-width: 40px;
                margin: 2px;
            }
            QScrollBar::handle:horizontal:hover {
                background: rgba(0, 0, 0, 0.35);
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }

            /* 菜单栏 */
            QMenuBar {
                background-color: rgba(255, 255, 255, 0.8);
                border-bottom: 1px solid rgba(0, 0, 0, 0.05);
            }
            QMenuBar::item {
                background: transparent;
                padding: 6px 12px;
            }
            QMenuBar::item:selected {
                background-color: rgba(0, 0, 0, 0.05);
                border-radius: 4px;
            }
            QMenu {
                background-color: rgba(255, 255, 255, 0.98);
                border: 1px solid rgba(0, 0, 0, 0.08);
                border-radius: 10px;
                padding: 6px;
                border-bottom: 2px solid rgba(0, 0, 0, 0.1); /* 模拟阴影 */
            }
            QMenu::item {
                padding: 6px 24px;
                border-radius: 6px;
            }
            QMenu::item:selected {
                background-color: #007AFF;
                color: white;
            }
            
            /* Tab Widget - 类似 Segmented Control */
            QTabWidget::pane {
                border: 1px solid rgba(255, 255, 255, 0.6);
                border-radius: 12px;
                background-color: rgba(255, 255, 255, 0.4);
            }
            QTabBar::tab {
                background-color: rgba(255, 255, 255, 0.3);
                color: #555;
                padding: 8px 20px;
                border-radius: 6px;
                margin: 4px 2px;
                border: 1px solid transparent;
            }
            QTabBar::tab:selected {
                background-color: #fff;
                color: #007AFF;
                font-weight: bold;
                border: 1px solid rgba(0, 0, 0, 0.05);
                border-bottom: 2px solid rgba(0, 0, 0, 0.05);
            }
            QTabBar::tab:hover:!selected {
                background-color: rgba(255, 255, 255, 0.5);
            }
        """)
        
        # 创建空的中心区域（显示背景）
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建各个dock窗口（包括曲线窗口）
        self.create_dock_windows()
        
        # 初始化串口参数显示标签
        if hasattr(self, 'label_serial_params'):
            self.update_serial_params_label()
        
        # 创建窗口菜单
        self.create_window_menu()
    
    def create_plot_dock(self):
        """创建曲线绘图dock窗口"""
        dock = QDockWidget('实时曲线', self)
        dock.setObjectName('PlotDock')
        dock.setAllowedAreas(Qt.AllDockWidgetAreas)
        # dock.setMinimumWidth(400)  # 移除最小宽度限制
        
        plot_widget = QWidget()
        main_layout = QVBoxLayout(plot_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)
        
        # 创建时钟组件(稍后添加到控制栏)
        clock_splitter = QSplitter(Qt.Horizontal)
        
        # 系统时钟
        system_clock_group = QGroupBox('系统时间')
        # 移除固定尺寸限制，使用布局自适应
        # system_clock_group.setMinimumWidth(340)
        # system_clock_group.setMaximumWidth(340)
        # system_clock_group.setMaximumHeight(200)
        system_clock_layout = QVBoxLayout(system_clock_group)
        system_clock_layout.setContentsMargins(8, 12, 8, 8)
        system_clock_layout.setSpacing(6)
        self.label_system_time = QLabel('--:--:--')
        self.label_system_time.setAlignment(Qt.AlignCenter)
        self.label_system_time.setStyleSheet("""
            QLabel {
                font-size: 15pt;
                font-weight: 600;
                color: #007AFF;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 rgba(255, 255, 255, 0.9),
                            stop:1 rgba(248, 248, 250, 0.9));
                border: 1px solid rgba(0, 122, 255, 0.15);
                border-radius: 10px;
                padding: 10px;
                min-height: 35px;
            }
        """)
        system_clock_layout.addWidget(self.label_system_time)
        
        self.label_system_date = QLabel('----/--/--')
        self.label_system_date.setAlignment(Qt.AlignCenter)
        self.label_system_date.setStyleSheet("""
            QLabel {
                font-size: 10pt;
                color: #86868B;
                background-color: transparent;
                padding: 5px;
                font-weight: 500;
            }
        """)
        system_clock_layout.addWidget(self.label_system_date)
        
        clock_splitter.addWidget(system_clock_group)
        
        # 接收时钟
        received_clock_group = QGroupBox('接收时间')
        # 移除固定尺寸限制，使用布局自适应
        # received_clock_group.setMinimumWidth(340)
        # received_clock_group.setMaximumWidth(340)
        # received_clock_group.setMaximumHeight(200)
        received_clock_layout = QVBoxLayout(received_clock_group)
        received_clock_layout.setContentsMargins(8, 12, 8, 8)
        received_clock_layout.setSpacing(6)
        self.label_received_time = QLabel('--:--:--')
        self.label_received_time.setAlignment(Qt.AlignCenter)
        self.label_received_time.setStyleSheet("""
            QLabel {
                font-size: 15pt;
                font-weight: 600;
                color: #34C759;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 rgba(255, 255, 255, 0.9),
                            stop:1 rgba(248, 248, 250, 0.9));
                border: 1px solid rgba(52, 199, 89, 0.15);
                border-radius: 10px;
                padding: 10px;
                min-height: 35px;
            }
        """)
        self.label_received_time.setContextMenuPolicy(Qt.CustomContextMenu)
        self.label_received_time.customContextMenuRequested.connect(self.configure_received_clock)
        self.label_received_time.setCursor(Qt.PointingHandCursor)
        received_clock_layout.addWidget(self.label_received_time)
        
        self.label_received_date = QLabel('----/--/--')
        self.label_received_date.setAlignment(Qt.AlignCenter)
        self.label_received_date.setStyleSheet("""
            QLabel {
                font-size: 10pt;
                color: #86868B;
                background-color: transparent;
                padding: 5px;
                font-weight: 500;
            }
        """)
        received_clock_layout.addWidget(self.label_received_date)
        
        # 添加配置和校准按钮
        btn_layout = QHBoxLayout()
        btn_config_clock = QPushButton('配置')
        btn_config_clock.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #f5f7fa);
                border: 1px solid #dcdfe6;
                border-bottom: 2px solid #c0c4cc;
                border-radius: 10px;
                color: #606266;
                font-size: 7pt;
                padding: 2px;
            }
            QPushButton:hover {
                background-color: #ecf5ff;
                color: #409EFF;
                border-color: #c6e2ff;
            }
            QPushButton:pressed {
                background-color: #f5f7fa;
                border-top: 2px solid #c0c4cc;
                border-bottom: none;
            }
        """)
        btn_config_clock.setMaximumHeight(20)
        btn_config_clock.clicked.connect(self.configure_received_clock)
        btn_layout.addWidget(btn_config_clock)
        
        btn_calibrate_time = QPushButton('时间校准')
        btn_calibrate_time.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FF9800, stop:1 #F57C00);
                border: 1px solid #F57C00;
                border-bottom: 2px solid #E65100;
                border-radius: 10px;
                color: white;
                font-size: 7pt;
                font-weight: bold;
                padding: 2px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FFB74D, stop:1 #FF9800);
            }
            QPushButton:pressed {
                background-color: #F57C00;
                border-top: 2px solid #E65100;
                border-bottom: none;
            }
        """)
        btn_calibrate_time.setMaximumHeight(20)
        btn_calibrate_time.setToolTip('发送电脑当前时间到设备')
        btn_calibrate_time.clicked.connect(self.calibrate_time)
        btn_layout.addWidget(btn_calibrate_time)
        
        received_clock_layout.addLayout(btn_layout)
        
        clock_splitter.addWidget(received_clock_group)
        
        # 设置初始宽度比例，允许自由缩放
        clock_splitter.setSizes([200, 200])
        # clock_splitter.setStretchFactor(0, 0)
        # clock_splitter.setStretchFactor(1, 0)
        
        # 保存时钟分割器引用
        self.clock_splitter = clock_splitter
        
        # 系统时钟更新定时器
        self.system_clock_timer = QTimer()
        self.system_clock_timer.timeout.connect(self.update_system_clock)
        self.system_clock_timer.start(1000)  # 每秒1秒更新
        
        # 曲线控制栏
        curve_control_layout = QHBoxLayout()
        curve_control_layout.setSpacing(8)
        
        # 曲线配置按钮
        btn_configure_curves = QPushButton('配置曲线')
        btn_configure_curves.setMinimumHeight(28)
        btn_configure_curves.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2196F3, stop:1 #1976D2);
                border: 1px solid #1976D2;
                border-bottom: 2px solid #0D47A1;
                border-radius: 6px;
                color: white;
                font-size: 8pt;
                font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #42A5F5, stop:1 #2196F3);
            }
            QPushButton:pressed {
                background-color: #1976D2;
                border-top: 2px solid #0D47A1;
                border-bottom: none;
            }
        """)
        btn_configure_curves.setToolTip('配置曲线名称、数据源、颜色等')
        btn_configure_curves.clicked.connect(self.configure_curves)
        curve_control_layout.addWidget(btn_configure_curves)
        
        # 显示曲线复选框（最多50条）
        curve_control_layout.addWidget(QLabel('显示:'))
        self.curve_checkboxes = []
        for i in range(50):
            checkbox = QCheckBox(f'曲线{i+1}')
            checkbox.setChecked(i < 3)  # 默认前3条启用
            checkbox.toggled.connect(lambda checked, idx=i: self.toggle_curve(idx, checked))
            curve_control_layout.addWidget(checkbox)
            self.curve_checkboxes.append(checkbox)
            if i >= 3:  # 默认隐藏后两个
                checkbox.hide()
        
        # 添加分隔线
       
        separator1 = QLabel('|')
        separator1.setStyleSheet('color: #ccc; font-weight: bold;')
        curve_control_layout.addWidget(separator1)
        
        # X轴缩放控制
        curve_control_layout.addWidget(QLabel('X轴:'))
        btn_x_zoom_in = QPushButton('放大')
        btn_x_zoom_in.setMinimumWidth(50)
        btn_x_zoom_in.setMaximumWidth(65)
        btn_x_zoom_in.setMinimumHeight(28)
        btn_x_zoom_in.setStyleSheet('''
            QPushButton {
                font-size: 8pt;
                font-weight: 600;
                color: #007AFF;
                background-color: rgba(255, 255, 255, 0.8);
                border: 1px solid rgba(0, 122, 255, 0.3);
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #007AFF;
                color: white;
            }
            QPushButton:pressed {
                background-color: rgba(0, 100, 220, 0.9);
            }
        ''')
        btn_x_zoom_in.setToolTip('放大X轴(显示更少时间范围)')
        btn_x_zoom_in.clicked.connect(lambda: self.plot_canvas.zoom_x(0.8))
        curve_control_layout.addWidget(btn_x_zoom_in)
        
        btn_x_zoom_out = QPushButton('缩小')
        btn_x_zoom_out.setMinimumWidth(50)
        btn_x_zoom_out.setMaximumWidth(65)
        btn_x_zoom_out.setMinimumHeight(28)
        btn_x_zoom_out.setStyleSheet('''
            QPushButton {
                font-size: 8pt;
                font-weight: 600;
                color: #007AFF;
                background-color: rgba(255, 255, 255, 0.8);
                border: 1px solid rgba(0, 122, 255, 0.3);
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #007AFF;
                color: white;
            }
            QPushButton:pressed {
                background-color: rgba(0, 100, 220, 0.9);
            }
        ''')
        btn_x_zoom_out.setToolTip('缩小X轴(显示更多时间范围)')
        btn_x_zoom_out.clicked.connect(lambda: self.plot_canvas.zoom_x(1.25))
        curve_control_layout.addWidget(btn_x_zoom_out)
        
        # 添加分隔线
        separator2 = QLabel('|')
        separator2.setStyleSheet('color: #ccc; font-weight: bold;')
        curve_control_layout.addWidget(separator2)
        
        # Y轴缩放控制
        curve_control_layout.addWidget(QLabel('Y轴:'))
        btn_y_zoom_in = QPushButton('放大')
        btn_y_zoom_in.setMinimumWidth(50)
        btn_y_zoom_in.setMaximumWidth(65)
        btn_y_zoom_in.setMinimumHeight(28)
        btn_y_zoom_in.setStyleSheet('''
            QPushButton {
                font-size: 8pt;
                font-weight: 600;
                color: #34C759;
                background-color: rgba(255, 255, 255, 0.8);
                border: 1px solid rgba(52, 199, 89, 0.3);
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #34C759;
                color: white;
            }
            QPushButton:pressed {
                background-color: rgba(40, 180, 70, 0.9);
            }
        ''')
        btn_y_zoom_in.setToolTip('放大Y轴(显示更小数值范围)')
        btn_y_zoom_in.clicked.connect(lambda: self.plot_canvas.zoom_y(0.8))
        curve_control_layout.addWidget(btn_y_zoom_in)
        
        btn_y_zoom_out = QPushButton('缩小')
        btn_y_zoom_out.setMinimumWidth(50)
        btn_y_zoom_out.setMaximumWidth(65)
        btn_y_zoom_out.setMinimumHeight(28)
        btn_y_zoom_out.setStyleSheet('''
            QPushButton {
                font-size: 8pt;
                font-weight: 600;
                color: #34C759;
                background-color: rgba(255, 255, 255, 0.8);
                border: 1px solid rgba(52, 199, 89, 0.3);
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #34C759;
                color: white;
            }
            QPushButton:pressed {
                background-color: rgba(40, 180, 70, 0.9);
            }
        ''')
        btn_y_zoom_out.setToolTip('缩小Y轴(显示更大数值范围)')
        btn_y_zoom_out.clicked.connect(lambda: self.plot_canvas.zoom_y(1.25))
        curve_control_layout.addWidget(btn_y_zoom_out)
        
        # 添加分隔线
        separator3 = QLabel('|')
        separator3.setStyleSheet('color: #ccc; font-weight: bold;')
        curve_control_layout.addWidget(separator3)
        
        # 重置视图按钮
        btn_reset_view = QPushButton('重置视图')
        btn_reset_view.setMinimumWidth(60)
        btn_reset_view.setMinimumHeight(28)
        btn_reset_view.setStyleSheet('font-size: 8pt;')
        btn_reset_view.setToolTip('重置为自动缩放模式，显示全部数据')
        btn_reset_view.clicked.connect(self.reset_plot_view)
        curve_control_layout.addWidget(btn_reset_view)
        
        curve_control_layout.addStretch()
        
        # 在控制栏右侧添加时钟显示
        curve_control_layout.addWidget(clock_splitter)
        
        # 绘图画布
        self.plot_canvas = PlotCanvas(self, width=8, height=6)
        # 设置大小策略为Expanding，确保随窗口缩放
        self.plot_canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.plot_canvas.updateGeometry()
        
        toolbar = NavigationToolbar(self.plot_canvas, self)
        
        main_layout.addLayout(curve_control_layout)
        main_layout.addWidget(toolbar)
        main_layout.addWidget(self.plot_canvas, stretch=1)
        
        dock.setWidget(plot_widget)
        self.addDockWidget(Qt.TopDockWidgetArea, dock)
        self.plot_dock = dock
    
    def create_dock_windows(self):
        """创建所有dock窗口"""
        if self.splash:
            self.splash.update_progress(30, '正在创建界面组件...')
        
        # 0. 曲线绘图dock
        self.create_plot_dock()
        
        if self.splash:
            self.splash.update_progress(35, '正在创建串口配置...')
        
        # 1. 串口配置dock
        self.create_serial_config_dock()
        
        if self.splash:
            self.splash.update_progress(40, '正在创建数据显示...')
        
        # 2. 实时数据显示dock
        self.create_custom_display_dock()
        
        # 3. 位状态显示dock（新增独立窗口）
        self.create_bit_display_dock()
        
        if self.splash:
            self.splash.update_progress(45, '正在创建发送模块...')
        
        # 4. 数据发送dock
        self.create_send_dock()
        
        if self.splash:
            self.splash.update_progress(55, '正在创建监视模块...')
        
        # 5. 数据监视dock
        self.create_data_monitor_dock()
        
        if self.splash:
            self.splash.update_progress(65, '正在初始化时钟...')
        
        # 6. 时钟dock
        self.create_clock_dock()
        
        if self.splash:
            self.splash.update_progress(70, '正在创建日志窗口...')
        
        # 7. 日志窗口dock
        self.create_log_dock()
    
    def create_serial_config_dock(self):
        """创建串口配置dock窗口"""
        dock = QDockWidget('串口配置', self)
        dock.setObjectName('SerialConfigDock')
        dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        # dock.setMinimumWidth(280)  # 移除最小宽度限制
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)
        layout.setContentsMargins(5, 5, 5, 5)
        
        self.create_serial_config_group(layout)
        layout.addStretch()
        
        scroll.setWidget(panel)
        dock.setWidget(scroll)
        self.addDockWidget(Qt.LeftDockWidgetArea, dock)
        self.serial_config_dock = dock
    
    def create_custom_display_dock(self):
        """创建实时数据显示dock窗口"""
        dock = QDockWidget('实时数据显示', self)
        dock.setObjectName('CustomDisplayDock')
        dock.setAllowedAreas(Qt.AllDockWidgetAreas)
        # dock.setMinimumWidth(280)  # 移除最小宽度限制
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)
        layout.setContentsMargins(5, 5, 5, 5)
        
        self.create_custom_display_group(layout)
        layout.addStretch()
        
        scroll.setWidget(panel)
        dock.setWidget(scroll)
        self.addDockWidget(Qt.LeftDockWidgetArea, dock)
        self.custom_display_dock = dock
    
    def create_bit_display_dock(self):
        """创建位(Bit)状态显示dock窗口"""
        dock = QDockWidget('位(Bit)状态显示', self)
        dock.setObjectName('BitDisplayDock')
        dock.setAllowedAreas(Qt.AllDockWidgetAreas)
        # dock.setMinimumWidth(280)  # 移除最小宽度限制
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)
        layout.setContentsMargins(5, 5, 5, 5)
        
        self.create_bit_display_group(layout)
        layout.addStretch()
        
        scroll.setWidget(panel)
        dock.setWidget(scroll)
        self.addDockWidget(Qt.LeftDockWidgetArea, dock)
        self.bit_display_dock = dock
    
    def create_send_dock(self):
        """创建数据发送dock窗口"""
        dock = QDockWidget('数据发送', self)
        dock.setObjectName('SendDock')
        dock.setAllowedAreas(Qt.AllDockWidgetAreas)
        # dock.setMinimumWidth(280)  # 移除最小宽度限制
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)
        layout.setContentsMargins(5, 5, 5, 5)
        
        self.create_send_group(layout)
        layout.addStretch()
        
        scroll.setWidget(panel)
        dock.setWidget(scroll)
        self.addDockWidget(Qt.LeftDockWidgetArea, dock)
        self.send_dock = dock
    
    def create_data_monitor_dock(self):
        """创建数据监视dock窗口"""
        dock = QDockWidget('收发数据监视', self)
        dock.setObjectName('DataMonitorDock')
        dock.setAllowedAreas(Qt.AllDockWidgetAreas)
        dock.setMinimumWidth(280)  # 设置最小宽度，防止被挤压
        
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # 显示控制栏
        control_layout = QHBoxLayout()
        self.radio_hex = QCheckBox("HEX显示")
        self.radio_hex.setChecked(True)
        self.radio_hex.toggled.connect(self.refresh_data_display)
        self.radio_ascii = QCheckBox("ASCII显示")
        self.radio_ascii.toggled.connect(self.refresh_data_display)
        self.check_display_timestamp = QCheckBox("显示时间戳")
        self.check_display_timestamp.setChecked(True)
        self.check_display_timestamp.toggled.connect(self.refresh_data_display)
        
        btn_clear = QPushButton("清空")
        btn_clear.setMaximumWidth(80)
        btn_clear.clicked.connect(self.clear_data_display)
        
        control_layout.addWidget(self.radio_hex)
        control_layout.addWidget(self.radio_ascii)
        control_layout.addWidget(self.check_display_timestamp)
        
        # 帧间隔设置
        control_layout.addWidget(QLabel('|'))
        control_layout.addWidget(QLabel('帧间隔(ms):'))
        self.spin_frame_interval = QSpinBox()
        self.spin_frame_interval.setRange(10, 5000)
        self.spin_frame_interval.setValue(100)
        self.spin_frame_interval.setSingleStep(10)
        self.spin_frame_interval.setMaximumWidth(80)
        self.spin_frame_interval.setToolTip('数据接收间隔超过此时间则作为新帧显示')
        self.spin_frame_interval.valueChanged.connect(self.on_frame_interval_changed)
        control_layout.addWidget(self.spin_frame_interval)
        
        control_layout.addStretch()
        control_layout.addWidget(btn_clear)
        layout.addLayout(control_layout)
        
        # 收发数据显示区
        self.text_data_display = QTextEdit()
        self.text_data_display.setReadOnly(True)
        self.text_data_display.setFont(QFont("Courier New", 9))
        layout.addWidget(self.text_data_display)
        
        # 数据缓存
        self.data_log = []
        
        dock.setWidget(panel)
        self.addDockWidget(Qt.BottomDockWidgetArea, dock)
        self.data_monitor_dock = dock
    
    def create_clock_dock(self):
        """创建时钟dock窗口"""
        dock = QDockWidget('时间显示', self)
        dock.setObjectName('ClockDock')
        dock.setAllowedAreas(Qt.TopDockWidgetArea | Qt.BottomDockWidgetArea | Qt.RightDockWidgetArea)
        dock.setMinimumWidth(280)  # 设置最小宽度，防止被挤压
        
        panel = QWidget()
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # 直接将时钟分割器添加到dock
        layout.addWidget(self.clock_splitter)
        layout.addStretch()
        
        dock.setWidget(panel)
        self.addDockWidget(Qt.TopDockWidgetArea, dock)
        self.clock_dock = dock
    
    def create_log_dock(self):
        """创建日志窗口dock"""
        dock = QDockWidget('调试日志', self)
        dock.setObjectName('LogDock')
        dock.setAllowedAreas(Qt.AllDockWidgetAreas)
        dock.setMinimumWidth(350)  # 设置最小宽度，防止被挤压
        
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # 控制栏
        control_layout = QHBoxLayout()
        
        self.check_auto_scroll = QCheckBox("自动滚动")
        self.check_auto_scroll.setChecked(True)
        control_layout.addWidget(self.check_auto_scroll)
        
        btn_clear_log = QPushButton("清空日志")
        btn_clear_log.setMaximumWidth(80)
        btn_clear_log.clicked.connect(self.clear_log)
        control_layout.addWidget(btn_clear_log)
        
        btn_export_log = QPushButton("导出日志")
        btn_export_log.setMaximumWidth(80)
        btn_export_log.clicked.connect(self.export_log)
        control_layout.addWidget(btn_export_log)
        
        control_layout.addStretch()
        layout.addLayout(control_layout)
        
        # 日志显示区
        self.text_log = QTextEdit()
        self.text_log.setReadOnly(True)
        self.text_log.setFont(QFont("Consolas", 9))
        self.text_log.setStyleSheet("""
            QTextEdit {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 rgba(30, 30, 30, 0.95),
                            stop:1 rgba(20, 20, 20, 0.95));
                color: #e8e8e8;
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 8px;
                padding: 10px;
                font-family: 'Consolas', 'Courier New', monospace;
            }
        """)
        layout.addWidget(self.text_log)
        
        dock.setWidget(panel)
        self.addDockWidget(Qt.BottomDockWidgetArea, dock)
        self.log_dock = dock
    
    def create_window_menu(self):
        """创建窗口菜单用于显示/隐藏各个dock"""
        menubar = self.menuBar()
        
        # 创建数据菜单
        data_menu = menubar.addMenu('数据')
        
        # 数据导出子菜单
        export_action = QAction('导出Excel...', self)
        export_action.setShortcut('Ctrl+E')
        export_action.setToolTip('导出曲线数据到Excel文件')
        export_action.triggered.connect(self.export_to_excel)
        data_menu.addAction(export_action)
        
        export_image_action = QAction('导出图片...', self)
        export_image_action.setShortcut('Ctrl+I')
        export_image_action.setToolTip('保存曲线截图为图片')
        export_image_action.triggered.connect(self.export_to_image)
        data_menu.addAction(export_image_action)
        
        export_raw_action = QAction('导出原始数据...', self)
        export_raw_action.setToolTip('导出收到的原始串口数据')
        export_raw_action.triggered.connect(self.export_raw_data)
        data_menu.addAction(export_raw_action)
        
        data_menu.addSeparator()
        
        # 自动保存路径设置
        auto_save_path_action = QAction('设置自动保存路径...', self)
        auto_save_path_action.setToolTip('设置超过2小时自动保存的路径')
        auto_save_path_action.triggered.connect(self.browse_auto_save_path)
        data_menu.addAction(auto_save_path_action)
        
        data_menu.addSeparator()
        
        # 清空数据
        clear_action = QAction('清空曲线数据', self)
        clear_action.setShortcut('Ctrl+Shift+C')
        clear_action.setToolTip('清除所有曲线数据')
        clear_action.triggered.connect(self.clear_plot_data)
        data_menu.addAction(clear_action)
        
        # 创建配置菜单
        config_menu = menubar.addMenu('配置')
        
        # 保存配置
        save_config_action = QAction('保存配置', self)
        save_config_action.setShortcut('Ctrl+S')
        save_config_action.setToolTip('保存当前配置到serial_config.json')
        save_config_action.triggered.connect(self.save_config)
        config_menu.addAction(save_config_action)
        
        config_menu.addSeparator()
        
        # 协议管理
        protocol_action = QAction('接收协议管理...', self)
        protocol_action.setToolTip('配置和管理接收数据协议')
        protocol_action.triggered.connect(self.open_protocol_manager)
        config_menu.addAction(protocol_action)
        
        config_menu.addSeparator()
        
        export_config_action = QAction('导出配置...', self)
        export_config_action.setShortcut('Ctrl+Shift+E')
        export_config_action.setToolTip('保存当前所有配置到文件')
        export_config_action.triggered.connect(self.export_config_file)
        config_menu.addAction(export_config_action)
        
        import_config_action = QAction('导入配置...', self)
        import_config_action.setShortcut('Ctrl+Shift+I')
        import_config_action.setToolTip('从文件加载配置')
        import_config_action.triggered.connect(self.import_config_file)
        config_menu.addAction(import_config_action)
        
        config_menu.addSeparator()
        
        # 窗口布局导入/导出
        export_layout_action = QAction('导出窗口布局...', self)
        export_layout_action.setToolTip('保存当前窗口布局到文件')
        export_layout_action.triggered.connect(self.export_window_layout_file)
        config_menu.addAction(export_layout_action)
        
        import_layout_action = QAction('导入窗口布局...', self)
        import_layout_action.setToolTip('从文件加载窗口布局')
        import_layout_action.triggered.connect(self.import_window_layout_file)
        config_menu.addAction(import_layout_action)
        
        # 创建窗口菜单
        window_menu = menubar.addMenu('窗口')
        
        # 添加各个dock的显示/隐藏动作
        window_menu.addAction(self.plot_dock.toggleViewAction())
        window_menu.addSeparator()
        window_menu.addAction(self.serial_config_dock.toggleViewAction())
        window_menu.addAction(self.custom_display_dock.toggleViewAction())
        window_menu.addAction(self.bit_display_dock.toggleViewAction())
        window_menu.addAction(self.send_dock.toggleViewAction())
        window_menu.addAction(self.data_monitor_dock.toggleViewAction())
        window_menu.addAction(self.clock_dock.toggleViewAction())
        window_menu.addAction(self.log_dock.toggleViewAction())
        
        window_menu.addSeparator()
        
        # 保存当前布局为默认
        save_layout_action = QAction('保存当前布局为默认', self)
        save_layout_action.setShortcut('Ctrl+Shift+S')
        save_layout_action.setToolTip('将当前窗口布局保存为默认布局')
        save_layout_action.triggered.connect(self.save_current_as_default_layout)
        window_menu.addAction(save_layout_action)
        
        # 恢复默认布局
        restore_action = QAction('恢复默认布局', self)
        restore_action.setShortcut('Ctrl+0')
        restore_action.setToolTip('恢复到保存的默认窗口布局')
        restore_action.triggered.connect(self.restore_default_layout)
        window_menu.addAction(restore_action)
        
        # 创建外观菜单
        appearance_menu = menubar.addMenu('外观')
        
        # 背景设置子菜单
        background_menu = appearance_menu.addMenu('背景设置')
        
        # 选择背景图片
        choose_bg_action = QAction('选择背景图片...', self)
        choose_bg_action.triggered.connect(self.choose_background_image)
        background_menu.addAction(choose_bg_action)
        
        # 清除背景图片
        clear_bg_action = QAction('清除背景图片', self)
        clear_bg_action.triggered.connect(self.clear_background_image)
        background_menu.addAction(clear_bg_action)
        
        background_menu.addSeparator()
        
        # 透明度设置
        opacity_action = QAction('调整透明度...', self)
        opacity_action.triggered.connect(self.adjust_background_opacity)
        background_menu.addAction(opacity_action)
        
        # 创建关于菜单
        help_menu = menubar.addMenu('关于')
        
        # 关于软件
        about_action = QAction('关于本软件', self)
        about_action.setShortcut('F1')
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)
    
    def setup_dock_signals(self):
        """设置dock窗口信号，在布局改变时保存"""
        # 连接所有dock窗口的信号
        docks = [
            self.serial_config_dock,
            self.custom_display_dock,
            self.bit_display_dock,
            self.send_dock,
            self.data_monitor_dock,
            self.clock_dock,
            self.log_dock
        ]
        
        for dock in docks:
            # 当dock窗口停靠位置改变时保存布局
            dock.dockLocationChanged.connect(self.on_dock_layout_changed)
            # 当dock窗口可见性改变时保存布局
            dock.visibilityChanged.connect(self.on_dock_layout_changed)
    
    def on_dock_layout_changed(self):
        """Dock布局改变时的处理"""
        # 如果正在恢复布局，不触发保存
        if self.is_restoring_layout:
            return
        
        # 不自动保存布局，仅记录日志
        # print("布局已改变（不自动保存，请手动保存）")
        pass
    
    def save_window_layout(self):
        """保存窗口布局到独立文件"""
        try:
            # 如果正在恢复布局，不保存
            if self.is_restoring_layout:
                return
            
            layout_data = {
                'window_state': self.saveState().toHex().data().decode(),
                'window_geometry': self.saveGeometry().toHex().data().decode(),
                'clock_splitter_sizes': self.clock_splitter.sizes()
            }
            
            with open(self.LAYOUT_FILE, 'w', encoding='utf-8') as f:
                json.dump(layout_data, f, indent=4, ensure_ascii=False)
            
            print(f"窗口布局已保存到: {self.LAYOUT_FILE}")
        except Exception as e:
            print(f"保存窗口布局失败: {e}")
    
    def load_window_layout(self):
        """从独立文件加载窗口布局"""
        if not os.path.exists(self.LAYOUT_FILE):
            print("窗口布局文件不存在，使用默认布局")
            return False
        
        try:
            self.is_restoring_layout = True
            
            with open(self.LAYOUT_FILE, 'r', encoding='utf-8') as f:
                layout_data = json.load(f)
            
            # 恢复窗口几何信息（大小和位置）
            if 'window_geometry' in layout_data:
                window_geometry = bytes.fromhex(layout_data['window_geometry'])
                self.restoreGeometry(window_geometry)
            
            # 恢复dock窗口状态
            if 'window_state' in layout_data:
                window_state = bytes.fromhex(layout_data['window_state'])
                result = self.restoreState(window_state)
                
                if not result:
                    print("警告: 窗口状态恢复失败")
                    return False
            
            # 恢复时钟分割器宽度
            if 'clock_splitter_sizes' in layout_data:
                self.clock_splitter.setSizes(layout_data['clock_splitter_sizes'])
            
            # 强制处理事件，确保布局完全应用
            QApplication.processEvents()
            
            print("窗口布局恢复成功")
            if hasattr(self, 'text_log'):
                self.log("窗口布局已恢复", "SUCCESS")
            return True
            
        except Exception as e:
            print(f"恢复窗口布局错误: {e}")
            if hasattr(self, 'text_log'):
                self.log(f"恢复窗口布局错误: {e}", "ERROR")
            import traceback
            traceback.print_exc()
            return False
        finally:
            self.is_restoring_layout = False
    
    def save_current_as_default_layout(self):
        """将当前布局保存为默认布局（覆盖当前布局文件）"""
        try:
            layout_data = {
                'window_state': self.saveState().toHex().data().decode(),
                'window_geometry': self.saveGeometry().toHex().data().decode(),
                'clock_splitter_sizes': self.clock_splitter.sizes()
            }
            
            with open(self.LAYOUT_FILE, 'w', encoding='utf-8') as f:
                json.dump(layout_data, f, indent=4, ensure_ascii=False)
            
            QMessageBox.information(self, "成功", "当前布局已保存为默认布局")
            print("默认布局已保存")
            if hasattr(self, 'text_log'):
                self.log("当前布局已保存为默认布局", "SUCCESS")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存默认布局失败: {e}")
            print(f"保存默认布局失败: {e}")
    
    def restore_default_layout(self):
        """恢复默认布局（从布局文件）"""
        try:
            # 检查是否存在布局文件
            if not os.path.exists(self.LAYOUT_FILE):
                reply = QMessageBox.question(
                    self, 
                    "提示", 
                    "未找到布局文件。\n是否将当前布局保存为默认布局？",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    self.save_current_as_default_layout()
                return
            
            # 加载布局
            result = self.load_window_layout()
            
            if result:
                QMessageBox.information(self, "成功", "默认布局已恢复")
                print("默认布局已恢复")
                if hasattr(self, 'text_log'):
                    self.log("默认布局已恢复", "SUCCESS")
            else:
                QMessageBox.warning(self, "警告", "恢复默认布局失败，请检查布局文件")
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"恢复默认布局失败: {e}")
            print(f"恢复默认布局失败: {e}")
    

    def create_serial_config_group(self, parent_layout):
        """创建串口配置组"""
        group = QGroupBox("串口配置")
        layout = QGridLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(10, 15, 10, 10)
        
        # 标题栏布局（COM口标签 + 配置按钮）
        title_layout = QHBoxLayout()
        title_layout.addWidget(QLabel("COM口:"))
        title_layout.addStretch()
        
        # 配置按钮（右上角）
        btn_config = QPushButton("⚙ 配置")
        btn_config.setMinimumHeight(24)
        btn_config.setMaximumWidth(60)
        btn_config.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #f5f7fa);
                border: 1px solid #dcdfe6;
                border-bottom: 2px solid #c0c4cc;
                border-radius: 4px;
                color: #606266;
                font-size: 8pt;
                padding: 2px;
            }
            QPushButton:hover {
                background-color: #ecf5ff;
                color: #409EFF;
                border-color: #c6e2ff;
            }
            QPushButton:pressed {
                background-color: #f5f7fa;
                border-top: 2px solid #c0c4cc;
                border-bottom: none;
            }
        """)
        btn_config.setToolTip('配置波特率、数据位、停止位、校验位')
        btn_config.clicked.connect(self.open_serial_params_dialog)
        title_layout.addWidget(btn_config)
        
        layout.addLayout(title_layout, 0, 0, 1, 3)
        
        # COM口选择
        self.combo_port = QComboBox()
        self.combo_port.setMinimumHeight(28)
        self.refresh_com_ports()
        layout.addWidget(self.combo_port, 1, 0, 1, 2)
        
        btn_refresh = QPushButton("刷新")
        btn_refresh.setMinimumHeight(32)
        btn_refresh.setMaximumWidth(70)
        btn_refresh.setStyleSheet('font-size: 8pt;')
        btn_refresh.clicked.connect(self.refresh_com_ports)
        layout.addWidget(btn_refresh, 1, 2)
        
        # 参数显示标签（显示当前配置）
        self.label_serial_params = QLabel("参数: 115200-8-1-None")
        self.label_serial_params.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 5px;
                color: #495057;
                font-size: 8pt;
            }
        """)
        self.label_serial_params.setAlignment(Qt.AlignCenter)
        self.label_serial_params.setToolTip('当前串口参数配置')
        layout.addWidget(self.label_serial_params, 2, 0, 1, 3)
        
        # 隐藏的配置控件（用于保存配置）
        self.combo_baudrate = QComboBox()
        self.combo_baudrate.addItems(['2400', '4800', '9600', '14400', '19200', '38400', 
                                      '57600', '115200', '128000', '256000', '460800', '921600'])
        self.combo_baudrate.setCurrentText('115200')
        self.combo_baudrate.hide()
        
        self.combo_databits = QComboBox()
        self.combo_databits.addItems(['5', '6', '7', '8'])
        self.combo_databits.setCurrentText('8')
        self.combo_databits.hide()
        
        self.combo_stopbits = QComboBox()
        self.combo_stopbits.addItems(['1', '1.5', '2'])
        self.combo_stopbits.setCurrentText('1')
        self.combo_stopbits.hide()
        
        self.combo_parity = QComboBox()
        self.combo_parity.addItems(['None', 'Even', 'Odd', 'Mark', 'Space'])
        self.combo_parity.setCurrentText('None')
        self.combo_parity.hide()
        
        # 自动重连
        self.check_auto_reconnect = QCheckBox("自动重连")
        layout.addWidget(self.check_auto_reconnect, 3, 0, 1, 3)
        
        # 恢复默认配置按钮
        btn_restore_default = QPushButton("恢复默认配置")
        btn_restore_default.setMinimumHeight(28)
        btn_restore_default.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #f5f7fa);
                border: 1px solid #dcdfe6;
                border-bottom: 2px solid #c0c4cc;
                border-radius: 6px;
                color: #606266;
                font-size: 8pt;
            }
            QPushButton:hover {
                background-color: #ecf5ff;
                color: #409EFF;
                border-color: #c6e2ff;
            }
            QPushButton:pressed {
                background-color: #f5f7fa;
                border-top: 2px solid #c0c4cc;
                border-bottom: none;
            }
        """)
        btn_restore_default.clicked.connect(self.restore_default_serial_config)
        layout.addWidget(btn_restore_default, 4, 0, 1, 3)
        
        # 连接按钮
        self.btn_connect = QPushButton("打开串口")
        self.btn_connect.setMinimumHeight(38)
        self.btn_connect.setStyleSheet("""
            QPushButton { 
                background-color: #4CAF50; 
                color: white; 
                font-weight: bold; 
                font-size: 8pt;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.btn_connect.clicked.connect(self.toggle_serial)
        layout.addWidget(self.btn_connect, 5, 0, 1, 3)
        
        # 曲线记录控制
        record_group = QGroupBox("曲线记录")
        record_layout = QVBoxLayout()
        record_layout.setSpacing(5)
        
        # 状态标签
        self.label_recording_status = QLabel("等待开始记录")
        self.label_recording_status.setAlignment(Qt.AlignCenter)
        self.label_recording_status.setStyleSheet("color: #666; font-size: 8pt;")
        record_layout.addWidget(self.label_recording_status)
        
        # 按钮布局
        btn_record_layout = QHBoxLayout()
        
        self.btn_start_record = QPushButton("开始记录")
        self.btn_start_record.setMinimumHeight(28)
        self.btn_start_record.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #67c23a, stop:1 #529b2e);
                border: 1px solid #529b2e;
                border-bottom: 2px solid #3e7523;
                border-radius: 6px;
                color: white;
                font-weight: bold;
                font-size: 8pt;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #85ce61, stop:1 #67c23a);
            }
            QPushButton:pressed {
                background-color: #529b2e;
                border-top: 2px solid #3e7523;
                border-bottom: none;
            }
        """)
        self.btn_start_record.clicked.connect(self.start_curve_recording)
        btn_record_layout.addWidget(self.btn_start_record)
        
        self.btn_stop_record = QPushButton("完成记录")
        self.btn_stop_record.setMinimumHeight(28)
        self.btn_stop_record.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #E6A23C, stop:1 #d48806);
                border: 1px solid #d48806;
                border-bottom: 2px solid #b87100;
                border-radius: 6px;
                color: white;
                font-weight: bold;
                font-size: 8pt;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ebb563, stop:1 #E6A23C);
            }
            QPushButton:pressed {
                background-color: #d48806;
                border-top: 2px solid #b87100;
                border-bottom: none;
            }
        """)
        self.btn_stop_record.clicked.connect(self.finish_curve_recording)
        btn_record_layout.addWidget(self.btn_stop_record)
        
        record_layout.addLayout(btn_record_layout)
        record_group.setLayout(record_layout)
        layout.addWidget(record_group, 6, 0, 1, 3)
        
        group.setLayout(layout)
        parent_layout.addWidget(group)
    
    def create_custom_display_group(self, parent_layout):
        """创建自定义显示窗口组"""
        self.custom_display_group = QGroupBox("实时数据显示")
        self.custom_display_group.setCheckable(True)
        self.custom_display_group.setChecked(True)
        layout = QVBoxLayout()
        layout.setSpacing(4)
        layout.setContentsMargins(5, 10, 5, 5)
        
        # 显示窗口数量控制
        control_layout = QHBoxLayout()
        display_label = QLabel('显示数量:')
        display_label.setMinimumWidth(70)  # 设置最小宽度以完整显示
        control_layout.addWidget(display_label)
        self.spin_display_count = QSpinBox()
        self.spin_display_count.setRange(1, 50)
        self.spin_display_count.setValue(5)
        self.spin_display_count.setMinimumHeight(22)
        self.spin_display_count.setMinimumWidth(65)  # 增加最小宽度
        self.spin_display_count.setMaximumWidth(65)
        self.spin_display_count.setToolTip('设置显示的窗口数量(1-50)')
        self.spin_display_count.valueChanged.connect(self.update_display_count)
        control_layout.addWidget(self.spin_display_count)
        control_layout.addStretch()
        layout.addLayout(control_layout)
        
        # 创建50个显示窗口
        self.display_widgets = []  # 保存每个显示窗口的容器
        for i in range(50):
            display_widget = QWidget()
            display_layout = QHBoxLayout(display_widget)
            display_layout.setContentsMargins(0, 1, 0, 1)
            
            # 显示标签(可点击配置)
            label = QLabel(f'显示{i+1}: --')
            label.setStyleSheet("""
                QLabel {
                    background-color: #f8f9fa;
                    border: 2px solid #dee2e6;
                    border-radius: 4px;
                    padding: 3px;
                    font-size: 7pt;
                    font-weight: bold;
                }
                QLabel:hover {
                    background-color: #e9ecef;
                    border-color: #adb5bd;
                }
            """)
            label.setMinimumHeight(28)
            label.setAlignment(Qt.AlignCenter)
            label.setContextMenuPolicy(Qt.CustomContextMenu)
            label.customContextMenuRequested.connect(lambda pos, idx=i: self.configure_custom_display(idx))
            label.setCursor(Qt.PointingHandCursor)
            label.setToolTip('右键配置此显示窗口')
            
            # 配置按钮
            btn_config = QPushButton('⚙')
            btn_config.setMinimumWidth(26)
            btn_config.setMaximumWidth(26)
            btn_config.setMinimumHeight(28)
            btn_config.setToolTip('配置此显示窗口')
            btn_config.setStyleSheet("""
                QPushButton {
                    background-color: #6c757d;
                    color: white;
                    border-radius: 5px;
                    font-size: 8pt;
                }
                QPushButton:hover {
                    background-color: #5a6268;
                }
            """)
            btn_config.clicked.connect(lambda checked, idx=i: self.configure_custom_display(idx))
            
            display_layout.addWidget(label, stretch=1)
            display_layout.addWidget(btn_config)
            
            layout.addWidget(display_widget)
            self.custom_display_labels.append(label)
            self.display_widgets.append(display_widget)
            
            # 默认只显示前5个
            if i >= 5:
                display_widget.hide()
        
        # 提示
        tip_label = QLabel('提示: 右键或点击⚙配置显示内容')
        tip_label.setStyleSheet('color: #666; font-size: 8pt; font-style: italic;')
        layout.addWidget(tip_label)
        
        self.custom_display_group.setLayout(layout)
        parent_layout.addWidget(self.custom_display_group)
    
    def create_bit_display_group(self, parent_layout):
        """创建位显示窗口组"""
        group = QGroupBox("位(Bit)状态显示")
        group.setCheckable(True)
        group.setChecked(True)
        layout = QVBoxLayout()
        layout.setSpacing(4)
        layout.setContentsMargins(5, 10, 5, 5)
        
        # 显示窗口数量控制
        control_layout = QHBoxLayout()
        bit_display_label = QLabel('显示数量:')
        bit_display_label.setMinimumWidth(70)  # 设置最小宽度以完整显示
        control_layout.addWidget(bit_display_label)
        self.spin_bit_display_count = QSpinBox()
        self.spin_bit_display_count.setRange(1, 50)
        self.spin_bit_display_count.setValue(2)
        self.spin_bit_display_count.setMinimumHeight(22)
        self.spin_bit_display_count.setMinimumWidth(65)  # 增加最小宽度
        self.spin_bit_display_count.setMaximumWidth(65)
        self.spin_bit_display_count.setToolTip('设置显示的位窗口数量(1-50)')
        self.spin_bit_display_count.valueChanged.connect(self.update_bit_display_count)
        control_layout.addWidget(self.spin_bit_display_count)
        control_layout.addStretch()
        layout.addLayout(control_layout)
        
        # 创建50个位显示窗口
        for window_idx in range(50):
            window_group = QGroupBox(f'位显示窗口 {window_idx + 1}')
            window_layout = QVBoxLayout()
            window_layout.setSpacing(1)
            
            # 配置按钮行
            config_layout = QHBoxLayout()
            name_label = QLabel(f'窗口名称: 未配置')
            self.bit_display_name_labels.append(name_label)
            config_layout.addWidget(name_label)
            btn_config = QPushButton('⚙ 配置')
            btn_config.setMaximumWidth(60)
            btn_config.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #f5f7fa);
                    border: 1px solid #dcdfe6;
                    border-bottom: 2px solid #c0c4cc;
                    border-radius: 4px;
                    color: #606266;
                    font-size: 7pt;
                    padding: 1px;
                }
                QPushButton:hover {
                    background-color: #ecf5ff;
                    color: #409EFF;
                    border-color: #c6e2ff;
                }
                QPushButton:pressed {
                    background-color: #f5f7fa;
                    border-top: 2px solid #c0c4cc;
                    border-bottom: none;
                }
            """)
            btn_config.clicked.connect(lambda checked, idx=window_idx: self.configure_bit_display(idx))
            config_layout.addStretch()
            config_layout.addWidget(btn_config)
            window_layout.addLayout(config_layout)
            
            # 位显示区域（4行x2列，显示8个位）
            bits_grid = QGridLayout()
            bits_grid.setSpacing(1)
            
            window_bit_labels = []
            for bit_idx in range(8):
                row = bit_idx // 2
                col = bit_idx % 2
                
                bit_widget = QWidget()
                bit_layout = QHBoxLayout(bit_widget)
                bit_layout.setContentsMargins(1, 1, 1, 1)
                bit_layout.setSpacing(2)
                
                # 位名称
                name_label = QLabel(f'Bit{bit_idx}')
                name_label.setMinimumWidth(60)
                name_label.setStyleSheet('font-size: 7pt;')
                bit_layout.addWidget(name_label)
                
                # 状态指示器（圆点）
                status_label = QLabel('●')
                status_label.setStyleSheet('color: #ccc; font-size: 12pt;')
                status_label.setAlignment(Qt.AlignCenter)
                status_label.setMinimumWidth(25)
                bit_layout.addWidget(status_label)
                
                bit_widget.setStyleSheet("""
                    QWidget {
                        background-color: #f8f9fa;
                        border: 1px solid #dee2e6;
                        border-radius: 4px;
                    }
                """)
                
                bits_grid.addWidget(bit_widget, row, col)
                window_bit_labels.append((name_label, status_label))
            
            window_layout.addLayout(bits_grid)
            window_group.setLayout(window_layout)
            layout.addWidget(window_group)
            
            self.bit_display_labels.append(window_bit_labels)
            self.bit_display_widgets.append(window_group)
            
            # 默认只显示前2个
            if window_idx >= 2:
                window_group.hide()
        
        # 提示
        tip_label = QLabel('提示: 点击配置按钮设置要监控的字节和位名称')
        tip_label.setStyleSheet('color: #666; font-size: 8pt; font-style: italic;')
        layout.addWidget(tip_label)
        
        group.setLayout(layout)
        parent_layout.addWidget(group)
    
    def create_send_group(self, parent_layout):
        """创建数据发送组"""
        group = QGroupBox("数据发送")
        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.setContentsMargins(5, 10, 5, 5)
        
        # 发送输入框
        self.text_send = QTextEdit()
        self.text_send.setMaximumHeight(50)
        self.text_send.setPlaceholderText("输入要发送的数据...")
        layout.addWidget(self.text_send)
        
        # 发送模式第一行
        mode_layout1 = QHBoxLayout()
        self.check_send_hex = QCheckBox("HEX发送")
        self.check_send_hex.setStyleSheet("font-size: 8pt;")
        self.check_send_hex.setChecked(True)
        self.check_send_hex.toggled.connect(self.refresh_data_display)
        self.check_send_timestamp = QCheckBox("添加时间戳")
        self.check_send_timestamp.setStyleSheet("font-size: 8pt;")
        self.check_send_timestamp.toggled.connect(self.refresh_data_display)
        mode_layout1.addWidget(self.check_send_hex)
        mode_layout1.addWidget(self.check_send_timestamp)
        mode_layout1.addStretch()
        layout.addLayout(mode_layout1)
        
        # 发送模式第二行 - CRC校验
        mode_layout2 = QHBoxLayout()
        crc_label = QLabel("CRC校验:")
        crc_label.setStyleSheet("font-size: 8pt;")
        mode_layout2.addWidget(crc_label)
        self.combo_send_crc = QComboBox()
        self.combo_send_crc.setStyleSheet("font-size: 8pt;")
        self.combo_send_crc.addItems(['无', 'CCITT-CRC16', 'Modbus-CRC16', 'CRC16-XMODEM', '累加和', '异或'])
        mode_layout2.addWidget(self.combo_send_crc)
        mode_layout2.addStretch()
        layout.addLayout(mode_layout2)
        
        # 发送按钮
        btn_layout = QHBoxLayout()
        btn_send = QPushButton("发送")
        btn_send.setMinimumHeight(26)
        btn_send.setStyleSheet("""
            QPushButton { 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #409EFF, stop:1 #337ecc);
                border: 1px solid #337ecc;
                border-bottom: 2px solid #28619e;
                border-radius: 6px;
                color: white; 
                font-weight: bold;
                font-size: 7pt;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #66b1ff, stop:1 #409EFF);
            }
            QPushButton:pressed {
                background-color: #337ecc;
                border-top: 2px solid #28619e;
                border-bottom: none;
            }
        """)
        btn_send.clicked.connect(self.send_data)
        btn_send_file = QPushButton("发送文件")
        btn_send_file.setMinimumHeight(26)
        btn_send_file.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #f5f7fa);
                border: 1px solid #dcdfe6;
                border-bottom: 2px solid #c0c4cc;
                border-radius: 6px;
                color: #606266;
                font-size: 7pt;
            }
            QPushButton:hover {
                background-color: #ecf5ff;
                color: #409EFF;
                border-color: #c6e2ff;
            }
            QPushButton:pressed {
                background-color: #f5f7fa;
                border-top: 2px solid #c0c4cc;
                border-bottom: none;
            }
        """)
        btn_send_file.clicked.connect(self.send_file)
        btn_layout.addWidget(btn_send)
        btn_layout.addWidget(btn_send_file)
        layout.addLayout(btn_layout)
        
        # 预设命令区域
        preset_header = QHBoxLayout()
        preset_label = QLabel('预设命令')
        preset_label.setStyleSheet('font-weight: bold; font-size: 8pt; margin-top: 8px;')
        preset_label.setMinimumWidth(80)  # 设置最小宽度
        preset_header.addWidget(preset_label)
        
        # 预设命令显示数量控制
        preset_display_label = QLabel('提示:')
        preset_display_label.setMinimumWidth(45)
        preset_display_label.setStyleSheet('font-size: 7pt;')
        preset_header.addWidget(preset_display_label)
        self.spin_preset_count = QSpinBox()
        self.spin_preset_count.setRange(1, 50)
        self.spin_preset_count.setValue(20)
        self.spin_preset_count.setMinimumHeight(22)
        self.spin_preset_count.setMinimumWidth(70)  # 增加宽度以显示完整
        self.spin_preset_count.setMaximumWidth(70)
        self.spin_preset_count.setToolTip('设置显示的预设命令按钮数量(1-50)')
        self.spin_preset_count.valueChanged.connect(self.update_preset_count)
        preset_header.addWidget(self.spin_preset_count)
        preset_header.addStretch()
        layout.addLayout(preset_header)
        
        # 创建50个预设命令按钮(单列布局)
        preset_vbox = QVBoxLayout()
        preset_vbox.setSpacing(2)
        
        self.preset_button_widgets = []  # 保存按钮的容器
        
        for i in range(50):
            # 创建容器widget
            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            btn_layout.setSpacing(4)
            
            btn = QPushButton(f'预设{i+1}')
            btn.setMinimumHeight(24)
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda pos, idx=i: self.configure_preset(idx))
            btn.clicked.connect(lambda checked, idx=i: self.send_preset_command(idx))
            btn.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffffff, stop:1 #f5f7fa);
                    border: 1px solid #dcdfe6;
                    border-bottom: 2px solid #c0c4cc;
                    border-radius: 6px;
                    color: #909399;
                    font-size: 7pt;
                    text-align: left;
                    padding-left: 5px;
                }
                QPushButton:hover {
                    background-color: #ecf5ff;
                    color: #409EFF;
                    border-color: #c6e2ff;
                }
                QPushButton:pressed {
                    background-color: #f5f7fa;
                    border-top: 2px solid #c0c4cc;
                    border-bottom: none;
                }
            """)
            
            btn_layout.addWidget(btn, stretch=1)
            
            # 添加数据填充输入框
            data_spinbox = QSpinBox()
            data_spinbox.setRange(0, 65535)
            data_spinbox.setValue(0)
            data_spinbox.setMaximumWidth(70)
            data_spinbox.setMinimumHeight(24)
            data_spinbox.setToolTip('数据填充框(0-65535)')
            data_spinbox.setStyleSheet("""
                QSpinBox {
                    border: 1px solid #dee2e6;
                    border-radius: 4px;
                    padding: 2px;
                }
            """)
            btn_layout.addWidget(data_spinbox)
            self.preset_data_spinboxes.append(data_spinbox)
            
            preset_vbox.addWidget(btn_widget)
            
            self.preset_buttons.append(btn)
            self.preset_button_widgets.append(btn_widget)
            
            # 创建定时器
            timer = QTimer()
            timer.timeout.connect(lambda idx=i: self.send_preset_command(idx))
            self.preset_timers.append(timer)
        
        layout.addLayout(preset_vbox)
        
        # 提示文本
        tip_label = QLabel('右键点击配置命令')
        tip_label.setStyleSheet('color: #868e96; font-size: 7.5pt; font-style: italic;')
        layout.addWidget(tip_label)
        
        group.setLayout(layout)
        parent_layout.addWidget(group)
    
    def open_serial_params_dialog(self):
        """打开串口参数配置对话框"""
        # 获取当前参数
        current_params = {
            'baudrate': self.combo_baudrate.currentText(),
            'databits': self.combo_databits.currentText(),
            'stopbits': self.combo_stopbits.currentText(),
            'parity': self.combo_parity.currentText()
        }
        
        # 打开对话框
        dialog = SerialParamsDialog(self, current_params)
        if dialog.exec_() == QDialog.Accepted:
            params = dialog.get_params()
            # 更新参数
            self.combo_baudrate.setCurrentText(params['baudrate'])
            self.combo_databits.setCurrentText(params['databits'])
            self.combo_stopbits.setCurrentText(params['stopbits'])
            self.combo_parity.setCurrentText(params['parity'])
            # 更新显示标签
            self.update_serial_params_label()
    
    def update_serial_params_label(self):
        """更新串口参数显示标签"""
        baudrate = self.combo_baudrate.currentText()
        databits = self.combo_databits.currentText()
        stopbits = self.combo_stopbits.currentText()
        parity = self.combo_parity.currentText()
        self.label_serial_params.setText(f"参数: {baudrate}-{databits}-{stopbits}-{parity}")
    
    def restore_default_serial_config(self):
        """恢复串口默认配置"""
        self.combo_baudrate.setCurrentText('2400')
        self.combo_databits.setCurrentText('8')
        self.combo_stopbits.setCurrentText('1')
        self.combo_parity.setCurrentText('None')
        self.update_serial_params_label()
        QMessageBox.information(self, "提示", "已恢复默认串口配置")

    def start_curve_recording(self):
        """开始曲线记录"""
        if self.is_curve_recording:
            QMessageBox.information(self, "提示", "正在记录曲线中...")
            return
            
        # 自动打开串口
        if self.serial_port is None or not self.serial_port.is_open:
            self.open_serial()
            if self.serial_port is None or not self.serial_port.is_open:
                return  # 打开失败
        
        # 清空曲线数据
        self.plot_canvas.clear_data()
        
        # 设置状态
        self.is_curve_recording = True
        self.label_recording_status.setText("曲线记录中: 0 点")
        self.label_recording_status.setStyleSheet("color: #67c23a; font-weight: bold; font-size: 9pt;")
        
        # 禁用开始按钮，启用停止按钮
        self.btn_start_record.setEnabled(False)
        self.btn_stop_record.setEnabled(True)
        
    def finish_curve_recording(self):
        """完成曲线记录"""
        if not self.is_curve_recording:
            return
            
        # 导出数据
        self.export_to_excel()
        
        # 重置状态
        self.is_curve_recording = False
        self.label_recording_status.setText("等待重新开始记录")
        self.label_recording_status.setStyleSheet("color: #666; font-size: 8pt;")
        
        # 启用开始按钮，禁用停止按钮
        self.btn_start_record.setEnabled(True)
        self.btn_stop_record.setEnabled(False)


    def refresh_com_ports(self):
        """刷新COM口列表"""
        current_port = self.combo_port.currentText()
        self.combo_port.clear()
        
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        
        self.combo_port.addItems(port_list)
        
        # 尝试恢复之前选择的端口
        index = self.combo_port.findText(current_port)
        if index >= 0:
            self.combo_port.setCurrentIndex(index)
    
    def toggle_serial(self):
        """打开/关闭串口"""
        if self.serial_port is None or not self.serial_port.is_open:
            self.open_serial()
        else:
            self.close_serial()
    
    def open_serial(self):
        """打开串口"""
        try:
            port = self.combo_port.currentText()
            if not port:
                QMessageBox.warning(self, "警告", "请选择COM口")
                return
            
            baudrate = int(self.combo_baudrate.currentText())
            
            # 数据位
            databits_map = {'5': 5, '6': 6, '7': 7, '8': 8}
            databits = databits_map[self.combo_databits.currentText()]
            
            # 停止位
            stopbits_map = {'1': serial.STOPBITS_ONE, '1.5': serial.STOPBITS_ONE_POINT_FIVE, '2': serial.STOPBITS_TWO}
            stopbits = stopbits_map[self.combo_stopbits.currentText()]
            
            # 校验位
            parity_map = {'None': serial.PARITY_NONE, 'Even': serial.PARITY_EVEN, 
                         'Odd': serial.PARITY_ODD, 'Mark': serial.PARITY_MARK, 'Space': serial.PARITY_SPACE}
            parity = parity_map[self.combo_parity.currentText()]
            
            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=databits,
                stopbits=stopbits,
                parity=parity,
                timeout=0.1
            )
            
            # 启动接收线程
            self.serial_thread = SerialThread()
            self.serial_thread.set_serial(self.serial_port)
            self.serial_thread.data_received.connect(self.on_data_received)
            self.serial_thread.connection_lost.connect(self.on_connection_lost)
            self.serial_thread.start()
            
            # 启动绘图定时器
            self.plot_timer.start(1000)  # 1Hz
            
            self.btn_connect.setText("关闭串口")
            self.btn_connect.setStyleSheet("QPushButton { background-color: rgba(255, 59, 48, 0.9); color: white; font-weight: 600; padding: 8px; border-radius: 8px; }")
            
            # 启动已配置的周期发送
            for i, preset in enumerate(self.preset_commands):
                if preset and preset.get('periodic', False) and preset.get('period', 0) > 0:
                    # 使用round()确保精确到毫秒，避免截断误差
                    interval_ms = round(preset['period'] * 1000)
                    self.preset_timers[i].start(interval_ms)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开串口失败: {str(e)}")
    
    def close_serial(self):
        """关闭串口"""
        try:
            # 显示缓冲区剩余数据
            if len(self.rx_buffer) > 0:
                self.add_data_to_display('RX', bytes(self.rx_buffer))
                self.rx_buffer.clear()
            self.last_rx_time = None
            if self.serial_thread:
                self.serial_thread.stop()
                self.serial_thread = None
            
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
            
            self.serial_port = None
            
            # 停止绘图定时器
            self.plot_timer.stop()
            
            # 停止所有周期发送定时器
            for timer in self.preset_timers:
                if timer.isActive():
                    timer.stop()
        
            self.btn_connect.setText("打开串口")
            self.btn_connect.setStyleSheet("QPushButton { background-color: rgba(52, 199, 89, 0.9); color: white; font-weight: 600; padding: 8px; border-radius: 8px; }")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"关闭串口失败: {str(e)}")
    
    def on_data_received(self, data):
        """接收到数据的处理"""
        # 保存原始数据
        self.raw_data_buffer.append(data)
        
        # 帧间隔处理：判断是否为新帧
        current_time = datetime.now()
        is_new_frame = False
        
        if self.last_rx_time is None:
            is_new_frame = True
        else:
            time_diff = (current_time - self.last_rx_time).total_seconds() * 1000  # 转为毫秒
            if time_diff >= self.frame_interval_ms:
                is_new_frame = True
        
        if is_new_frame:
            # 如果缓冲区有数据，先显示缓冲的完整帧
            if len(self.rx_buffer) > 0:
                self.add_data_to_display('RX', bytes(self.rx_buffer))
                self.rx_buffer.clear()
            # 开始新帧
            self.rx_buffer.extend(data)
        else:
            # 继续累积当前帧
            self.rx_buffer.extend(data)
        
        self.last_rx_time = current_time
        
        # 检查是否有自定义协议与默认协议帧头冲突（默认协议: 0xA8 0xA8）
        use_default_parser = True
        default_header = bytes([0xA8, 0xA8])
        
        # 如果任何自定义协议使用了与默认协议相同的帧头，则禁用默认解析器避免重复
        for protocol_config in self.protocol_configs:
            if protocol_config.get('enabled', True):
                custom_header = bytes(protocol_config.get('header', []))
                if custom_header == default_header:
                    use_default_parser = False
                    break
        
        # 解析数据 - 使用默认解析器（仅当没有帧头冲突时）
        if use_default_parser:
            self.parser.add_data(data)
            results = self.parser.parse()
            
            for result in results:
                # 只在曲线记录模式下添加到曲线
                if self.is_curve_recording:
                    self.plot_canvas.add_data(result)
                
                # 更新自定义显示窗口 - 默认协议
                self.parse_custom_display_data(result['raw_frame'], protocol_name=None)
        
        # 解析数据 - 使用所有自定义协议解析器（始终执行，支持多协议并行）
        for parser in self.protocol_parsers:
            parser.add_data(data)
            protocol_results = parser.parse()
            
            for result in protocol_results:
                protocol_name = result['protocol_name']
                raw_frame = result['raw_frame']
                frame_hex = result['frame_hex']
                
                # 在调试日志中打印
                self.log(f"[{protocol_name}] 接收: {frame_hex}", "INFO")
                
                # 只在曲线记录模式下添加到曲线
                if self.is_curve_recording:
                    self.plot_canvas.add_data(result)
                
                # 更新自定义显示窗口 - 使用指定协议
                self.parse_custom_display_data(raw_frame, protocol_name=protocol_name)
    
    def on_frame_interval_changed(self, value):
        """帧间隔时间改变处理"""
        self.frame_interval_ms = value
    
    def flush_rx_buffer(self):
        """刷新接收缓冲区，显示超时的数据帧"""
        if len(self.rx_buffer) > 0 and self.last_rx_time is not None:
            current_time = datetime.now()
            time_diff = (current_time - self.last_rx_time).total_seconds() * 1000  # 转为毫秒
            
            # 如果超过帧间隔时间，显示缓冲区数据
            if time_diff >= self.frame_interval_ms:
                self.add_data_to_display('RX', bytes(self.rx_buffer))
                self.rx_buffer.clear()
                self.last_rx_time = None
    
    def on_connection_lost(self):
        """连接丢失处理"""
        # 显示缓冲区剩余数据
        if len(self.rx_buffer) > 0:
            self.add_data_to_display('RX', bytes(self.rx_buffer))
            self.rx_buffer.clear()
        self.last_rx_time = None
        
        # 先记录当前串口信息和周期发送状态，再停止定时器
        if self.serial_port:
            try:
                self.lost_port_info = {
                    'port': self.combo_port.currentText(),
                    'baudrate': int(self.combo_baudrate.currentText()),
                    'databits': int(self.combo_databits.currentText()),
                    'stopbits': self.combo_stopbits.currentText(),
                    'parity': self.combo_parity.currentText()
                }
                # 记录周期发送配置（在停止之前记录正在运行的定时器）
                active_presets = []
                for i, timer in enumerate(self.preset_timers):
                    if timer.isActive() and i < len(self.preset_commands):
                        preset = self.preset_commands[i]
                        if preset:
                            active_presets.append((i, preset))
                            self.log(f"记录预设命令{i+1}的周期发送状态，间隔={preset.get('period', 0)}秒", "DEBUG")
                
                self.lost_port_config = {
                    'preset_states': active_presets
                }
                self.log(f"保存了 {len(active_presets)} 个周期发送配置", "INFO")
            except Exception as e:
                self.log(f"保存配置失败 - {str(e)}", "ERROR")
        
        # 停止所有周期发送定时器，防止串口断开后继续发送导致异常
        for timer in self.preset_timers:
            if timer.isActive():
                timer.stop()
        
        # 停止绘图定时器
        if self.plot_timer.isActive():
            self.plot_timer.stop()
        
        # 关闭串口
        if self.serial_thread:
            self.serial_thread.stop()
            self.serial_thread = None
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.serial_port = None
        
        self.btn_connect.setText("正在检测...")
        self.btn_connect.setStyleSheet("QPushButton { background-color: rgba(255, 149, 0, 0.9); color: white; font-weight: 600; padding: 8px; border-radius: 8px; }")
        
        # 启动智能恢复检测
        self.start_smart_recovery()
    
    def start_smart_recovery(self):
        """启动智能恢夏检测"""
        if not self.lost_port_info:
            self.log("没有保存的端口信息，无法启动智能恢复", "WARNING")
            self.btn_connect.setText("打开串口")
            self.btn_connect.setStyleSheet("QPushButton { background-color: rgba(52, 199, 89, 0.9); color: white; font-weight: 600; padding: 8px; border-radius: 8px; }")
            return
        
        self.log(f"启动智能恢复，目标端口={self.lost_port_info.get('port', '未知')}", "INFO")
        
        self.recovery_start_time = datetime.now()
        self.port_stable_time = None
        self.last_port_state = False
        
        # 创建检测定时器，每100ms检测一次
        if self.recovery_check_timer:
            self.recovery_check_timer.stop()
        self.recovery_check_timer = QTimer()
        self.recovery_check_timer.timeout.connect(self.check_port_recovery)
        self.recovery_check_timer.start(100)
        self.log("检测定时器已启动（100ms间隔）", "DEBUG")
    
    def check_port_recovery(self):
        """检测串口是否恢复"""
        try:
            if not self.lost_port_info or not self.recovery_start_time:
                self.stop_recovery_check()
                return
            
            current_time = datetime.now()
            elapsed = (current_time - self.recovery_start_time).total_seconds()
            
            # 超过10秒，停止检测
            if elapsed > 10.0:
                self.log("10秒内未检测到串口，停止智能恢复", "WARNING")
                self.stop_recovery_check()
                return
            
            # 检测目标端口是否存在
            target_port = self.lost_port_info['port']
            ports = serial.tools.list_ports.comports()
            port_exists = any(port.device == target_port for port in ports)
            
            # 添加调试信息
            if elapsed < 0.5:  # 只在开始时打印一次
                self.log(f"开始检测端口 {target_port}，当前可用端口: {[p.device for p in ports]}", "DEBUG")
            
            # 每秒打印一次状态
            if int(elapsed * 10) % 10 == 0:
                self.log(f"检测中...已过{elapsed:.1f}秒，端口存在={port_exists}", "DEBUG")
            
            if port_exists:
                if not self.last_port_state:
                    # 端口刚刚插入
                    self.port_stable_time = current_time
                    self.last_port_state = True
                    self.log(f"检测到端口 {target_port} 重新插入", "INFO")
                else:
                    # 端口已经存在，检查是否稳定1秒
                    if self.port_stable_time:
                        stable_duration = (current_time - self.port_stable_time).total_seconds()
                        if stable_duration >= 1.0:
                            self.log(f"端口稳定 {stable_duration:.1f}秒，开始恢复连接", "INFO")
                            self.recover_connection()
                            return
            else:
                if self.last_port_state:
                    # 端口又断开了
                    self.log(f"端口 {target_port} 又断开", "WARNING")
                    self.port_stable_time = None
                self.last_port_state = False
        except Exception as e:
            self.log(f"智能恢复检测异常: {str(e)}", "ERROR")
            import traceback
            traceback.print_exc()
    
    def recover_connection(self):
        """恢复串口连接"""
        if not self.lost_port_info:
            self.log("没有记录的端口信息，退出恢复", "ERROR")
            self.stop_recovery_check()
            return
        
        # 先保存配置信息，再停止检测
        saved_port_info = self.lost_port_info.copy()
        saved_port_config = self.lost_port_config.copy() if self.lost_port_config else None
        
        self.stop_recovery_check()
        
        try:
            self.log("尝试恢复串口连接...", "INFO")
            self.log(f"端口信息 = {saved_port_info}", "DEBUG")
            
            # 恢复串口连接
            port = saved_port_info['port']
            baudrate = saved_port_info['baudrate']
            
            # 数据位
            databits = saved_port_info['databits']
            
            # 停止位
            stopbits_map = {'1': serial.STOPBITS_ONE, '1.5': serial.STOPBITS_ONE_POINT_FIVE, '2': serial.STOPBITS_TWO}
            stopbits = stopbits_map[saved_port_info['stopbits']]
            
            # 校验位
            parity_map = {'None': serial.PARITY_NONE, 'Even': serial.PARITY_EVEN, 
                         'Odd': serial.PARITY_ODD, 'Mark': serial.PARITY_MARK, 'Space': serial.PARITY_SPACE}
            parity = parity_map[saved_port_info['parity']]
            
            self.log(f"打开串口 {port}, {baudrate}, {databits}, {stopbits}, {parity}", "DEBUG")
            
            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=databits,
                stopbits=stopbits,
                parity=parity,
                timeout=0.1
            )
            
            self.log("串口已打开，启动接收线程", "DEBUG")
            
            # 启动接收线程
            self.serial_thread = SerialThread()
            self.serial_thread.set_serial(self.serial_port)
            self.serial_thread.data_received.connect(self.on_data_received)
            self.serial_thread.connection_lost.connect(self.on_connection_lost)
            self.serial_thread.start()
            
            # 启动绘图定时器
            self.plot_timer.start(1000)
            
            self.log(f"配置信息 = {saved_port_config}", "DEBUG")
            
            # 恢复周期发送
            if saved_port_config and 'preset_states' in saved_port_config:
                self.log(f"准备恢复 {len(saved_port_config['preset_states'])} 个周期发送", "INFO")
                for i, preset in saved_port_config['preset_states']:
                    if i < len(self.preset_timers):
                        interval_ms = round(preset['period'] * 1000)
                        self.preset_timers[i].start(interval_ms)
                        self.log(f"恢复预设命令{i+1}的周期发送，间隔{interval_ms}ms", "DEBUG")
            else:
                self.log("没有需要恢复的周期发送", "INFO")
            
            self.btn_connect.setText("关闭串口")
            self.btn_connect.setStyleSheet("QPushButton { background-color: rgba(255, 59, 48, 0.9); color: white; font-weight: 600; padding: 8px; border-radius: 8px; }")
            
            self.log(f"串口 {port} 恢复成功！", "SUCCESS")
            
        except Exception as e:
            self.log(f"智能恢复失败: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            self.btn_connect.setText("打开串口")
            self.btn_connect.setStyleSheet("QPushButton { background-color: rgba(52, 199, 89, 0.9); color: white; font-weight: 600; padding: 8px; border-radius: 8px; }")
    
    def stop_recovery_check(self):
        """停止恢复检测"""
        if self.recovery_check_timer:
            self.recovery_check_timer.stop()
            self.recovery_check_timer = None
        
        self.recovery_start_time = None
        self.port_stable_time = None
        self.last_port_state = False
        
        if self.btn_connect.text() == "正在检测...":
            self.btn_connect.setText("打开串口")
            self.btn_connect.setStyleSheet("QPushButton { background-color: rgba(52, 199, 89, 0.9); color: white; font-weight: 600; padding: 8px; border-radius: 8px; }")
        
        # 清空记录
        self.lost_port_info = None
        self.lost_port_config = None
    
    def try_reconnect(self):
        """尝试重新连接"""
        if self.serial_port is None or not self.serial_port.is_open:
            self.open_serial()
    
    def update_plot(self):
        """更新曲线显示（仅刷新界面，不记录新数据点）"""
        self.plot_canvas.update_plot()
        
        # 更新曲线记录状态显示
        if self.is_curve_recording:
            point_count = len(self.plot_canvas.time_data)
            self.label_recording_status.setText(f"曲线记录中: {point_count} 点")
            self.label_recording_status.setStyleSheet("color: #67c23a; font-weight: bold; font-size: 9pt;")
    
    def reset_plot_view(self):
        """重置曲线视图"""
        self.plot_canvas.reset_view()
    
    def update_system_clock(self):
        """更新系统时钟显示"""
        now = datetime.now()
        self.label_system_time.setText(now.strftime('%H:%M:%S'))
        self.label_system_date.setText(now.strftime('%Y/%m/%d'))
    
    def decimal_to_bcd(self, decimal_value):
        """将十进制数转换为BCD码
        例如: 23 -> 0x23
        """
        tens = (decimal_value // 10) & 0x0F
        ones = (decimal_value % 10) & 0x0F
        return (tens << 4) | ones
    
    def calibrate_time(self):
        """时间校准 - 发送电脑时间到设备"""
        if not self.serial_port or not self.serial_port.is_open:
            QMessageBox.warning(self, "警告", "串口未打开")
            return
        
        try:
            # 获取当前时间
            now = datetime.now()
            
            # 构建数据包
            data = bytearray()
            
            # 固定头部: 06 00 10 24 08
            data.extend([0x06, 0x00, 0x10, 0x24, 0x08])
            
            # 时间数据转BCD码
            year = now.year % 100  # 只取年份后两位
            month = now.month
            day = now.day
            hour = now.hour
            minute = now.minute
            second = now.second
            
            # 添加BCD码时间
            data.append(self.decimal_to_bcd(year))
            data.append(self.decimal_to_bcd(month))
            data.append(self.decimal_to_bcd(day))
            data.append(self.decimal_to_bcd(hour))
            data.append(self.decimal_to_bcd(minute))
            data.append(self.decimal_to_bcd(second))
            
            # 计算CRC16-XMODEM校验
            crc = CRCCalculator.calculate_crc16_xmodem(bytes(data))
            
            # 添加CRC (低字节在前，高字节在后)
            data.append(crc & 0xFF)
            data.append((crc >> 8) & 0xFF)
            
            # 发送数据
            self.serial_port.write(data)
            
            # 添加到显示区
            self.add_data_to_display('TX', bytes(data))
            
            # 显示提示信息
            time_str = f"{now.year:04d}-{month:02d}-{day:02d} {hour:02d}:{minute:02d}:{second:02d}"
            QMessageBox.information(self, "成功", f"时间校准命令已发送\n时间: {time_str}")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"时间校准失败: {str(e)}")
    
    def configure_received_clock(self):
        """配置接收时钟"""
        dialog = ClockConfigDialog(self, self.clock_config, self.protocol_configs)
        if dialog.exec_() == QDialog.Accepted:
            self.clock_config = dialog.get_data()
    
    def update_received_clock(self, data):
        """更新接收时钟显示"""
        if not self.clock_config or not self.clock_config.get('enabled', False):
            return
        
        try:
            # 解析年
            year_start = self.clock_config.get('year_start', 10)
            year_count = self.clock_config.get('year_count', 2)
            year_type = self.clock_config.get('year_type', 'uint16 (LE)')
            
            if len(data) < year_start + year_count:
                return
            
            year_bytes = data[year_start:year_start + year_count]
            year = self.parse_data_bytes(year_bytes, year_type)
            
            # 解析月
            month_start = self.clock_config.get('month_start', 12)
            if len(data) <= month_start:
                return
            month = data[month_start]
            
            # 解析日
            day_start = self.clock_config.get('day_start', 13)
            if len(data) <= day_start:
                return
            day = data[day_start]
            
            # 解析时
            hour_start = self.clock_config.get('hour_start', 14)
            if len(data) <= hour_start:
                return
            hour = data[hour_start]
            
            # 解析分
            minute_start = self.clock_config.get('minute_start', 15)
            if len(data) <= minute_start:
                return
            minute = data[minute_start]
            
            # 解析秒
            second_start = self.clock_config.get('second_start', 16)
            if len(data) <= second_start:
                return
            second = data[second_start]
            
            # 更新显示
            self.label_received_time.setText(f'{hour:02d}:{minute:02d}:{second:02d}')
            self.label_received_date.setText(f'{year:04d}/{month:02d}/{day:02d}')
            
            # 保存接收到的时间
            self.received_time = {
                'year': year,
                'month': month,
                'day': day,
                'hour': hour,
                'minute': minute,
                'second': second
            }
            
        except Exception as e:
            print(f"解析接收时钟错误: {e}")
    
    def configure_curves(self):
        """配置曲线对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle('曲线配置')
        dialog.setMinimumWidth(800)
        dialog.setMinimumHeight(600)
        
        layout = QVBoxLayout(dialog)
        
        # 提示文本
        tip_label = QLabel('最多可配置50条曲线，每条曲线可自定义名称、数据来源、颜色等')
        tip_label.setStyleSheet('color: #666; font-size: 9pt; font-style: italic; padding: 10px;')
        layout.addWidget(tip_label)
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # 创建滚动区域内的容器
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setSpacing(8)
        scroll_layout.setContentsMargins(5, 5, 5, 5)
        
        # 曲线配置按钮组
        for i in range(50):
            btn_frame = QWidget()
            btn_layout = QHBoxLayout(btn_frame)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            
            # 配置按钮
            btn = QPushButton(f'曲线{i+1}: 点击配置')
            btn.setMinimumHeight(40)
            btn.setStyleSheet("""
                QPushButton {
                    text-align: left;
                    padding-left: 15px;
                    font-size: 9pt;
                    background-color: #f5f5f5;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
            """)
            
            # 如果有配置，显示配置信息
            if i < len(self.plot_canvas.curve_configs):
                config = self.plot_canvas.curve_configs[i]
                name = config.get('name', f'曲线{i+1}')
                unit = config.get('unit', '')
                enabled = config.get('enabled', True)
                status = '✓' if enabled else '✗'
                btn.setText(f'{status} {name} ({unit})' if unit else f'{status} {name}')
                if enabled:
                    btn.setStyleSheet("""
                        QPushButton {
                            text-align: left;
                            padding-left: 15px;
                            font-size: 9pt;
                            background-color: #e8f5e9;
                            color: #2e7d32;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #c8e6c9;
                        }
                    """)
            
            btn.clicked.connect(lambda checked, idx=i: self.configure_single_curve(idx, dialog))
            btn_layout.addWidget(btn)
            
            scroll_layout.addWidget(btn_frame)
        
        scroll_widget.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)
        
        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)
        
        dialog.exec_()
        
        # 更新复选框标签
        self.update_curve_checkbox_labels()
    
    def configure_single_curve(self, index, parent_dialog):
        """配置单条曲线"""
        curve_data = None
        if index < len(self.plot_canvas.curve_configs):
            curve_data = self.plot_canvas.curve_configs[index]
        
        dialog = CurveConfigDialog(self, curve_data, self.protocol_configs)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            
            # 确保配置列表足够长
            while len(self.plot_canvas.curve_configs) <= index:
                self.plot_canvas.curve_configs.append({
                    'name': f'曲线{len(self.plot_canvas.curve_configs)+1}',
                    'start_byte': 2,
                    'byte_count': 2,
                    'data_type': 'uint16 (LE)',
                    'coefficient': 1.0,
                    'divisor': 1.0,
                    'offset': 0.0,
                    'unit': '',
                    'color': '蓝色',
                    'enabled': False
                })
            
            # 更新配置
            self.plot_canvas.curve_configs[index] = data
            
            # 重新应用配置
            self.plot_canvas.set_curve_configs(self.plot_canvas.curve_configs)
            
            # 更新复选框标签
            self.update_curve_checkbox_labels()
            
            # 刷新父对话框的按钮显示
            self.refresh_curve_config_dialog(parent_dialog, index)
    
    def refresh_curve_config_dialog(self, dialog, changed_index):
        """刷新曲线配置对话框的按钮显示"""
        try:
            # 找到对话框中的按钮容器布局
            layout = dialog.layout()
            if layout and layout.count() > 1:
                # 获取按钮容器布局（第二个项目）
                button_layout_item = layout.itemAt(1)
                if button_layout_item:
                    button_layout = button_layout_item.layout()
                    if button_layout:
                        # 更新对应的按钮
                        if changed_index < button_layout.count():
                            btn_frame = button_layout.itemAt(changed_index).widget()
                            if btn_frame:
                                btn = btn_frame.findChild(QPushButton)
                                if btn and changed_index < len(self.plot_canvas.curve_configs):
                                    config = self.plot_canvas.curve_configs[changed_index]
                                    name = config.get('name', f'曲线{changed_index+1}')
                                    unit = config.get('unit', '')
                                    enabled = config.get('enabled', True)
                                    status = '✓' if enabled else '✗'
                                    btn.setText(f'{status} {name} ({unit})' if unit else f'{status} {name}')
                                    if enabled:
                                        btn.setStyleSheet("""
                                            QPushButton {
                                                text-align: left;
                                                padding-left: 15px;
                                                font-size: 9pt;
                                                background-color: #e8f5e9;
                                                color: #2e7d32;
                                                font-weight: bold;
                                            }
                                            QPushButton:hover {
                                                background-color: #c8e6c9;
                                            }
                                        """)
                                    else:
                                        btn.setStyleSheet("""
                                            QPushButton {
                                                text-align: left;
                                                padding-left: 15px;
                                                font-size: 9pt;
                                                background-color: #f5f5f5;
                                            }
                                            QPushButton:hover {
                                                background-color: #e0e0e0;
                                            }
                                        """)
        except Exception as e:
            print(f"刷新对话框错误: {e}")
    
    def update_curve_checkbox_labels(self):
        """更新曲线复选框的标签（只显示启用的曲线）"""
        if not hasattr(self, 'curve_checkboxes'):
            return
        
        for i, checkbox in enumerate(self.curve_checkboxes):
            if i < len(self.plot_canvas.curve_configs):
                config = self.plot_canvas.curve_configs[i]
                enabled = config.get('enabled', False)
                # 只显示启用的曲线
                if enabled:
                    name = config.get('name', f'曲线{i+1}')
                    checkbox.setText(name)
                    checkbox.setChecked(True)
                    checkbox.show()
                else:
                    # 未启用的曲线隐藏
                    checkbox.hide()
            else:
                checkbox.setText(f'曲线{i+1}')
                checkbox.setChecked(False)
                checkbox.hide()
    
    def toggle_curve(self, index, visible):
        """切换曲线可见性"""
        self.plot_canvas.set_curve_visibility(index, visible)
        if index < len(self.plot_canvas.curve_configs):
            self.plot_canvas.curve_configs[index]['enabled'] = visible
    
    def clear_data_display(self):
        """清空收发数据显示"""
        self.data_log.clear()
        self.text_data_display.clear()
    
    def refresh_data_display(self):
        """刷新数据显示(根据时间戳和显示模式重新格式化)"""
        self.text_data_display.clear()
        
        show_timestamp = self.check_display_timestamp.isChecked()
        show_hex = self.radio_hex.isChecked()
        show_ascii = self.radio_ascii.isChecked()
        
        for record in self.data_log:
            direction = record['direction']  # 'TX' or 'RX'
            timestamp = record['timestamp']
            data = record['data']
            
            # 构建显示文本
            display_text = ""
            
            if show_timestamp:
                display_text += f"[{timestamp}] "
            
            display_text += f"{direction}: "
            
            if show_hex:
                hex_str = ' '.join([f'{b:02X}' for b in data])
                display_text += hex_str
            
            if show_ascii:
                try:
                    ascii_str = data.decode('utf-8', errors='ignore')
                    if show_hex:
                        display_text += f"  [{ascii_str}]"
                    else:
                        display_text += ascii_str
                except:
                    pass
            
            self.text_data_display.append(display_text)
        
        # 滚动到底部
        self.text_data_display.moveCursor(QTextCursor.End)
    
    def add_data_to_display(self, direction, data):
        """添加数据到显示区
        direction: 'TX' 或 'RX'
        data: bytes数据
        """
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        
        # 保存到日志
        self.data_log.append({
            'direction': direction,
            'timestamp': timestamp,
            'data': data
        })
        
        # 限制日志大小
        if len(self.data_log) > 1000:
            self.data_log.pop(0)
        
        # 构建显示文本
        display_text = ""
        
        if self.check_display_timestamp.isChecked():
            display_text += f"[{timestamp}] "
        
        display_text += f"{direction}: "
        
        if self.radio_hex.isChecked():
            hex_str = ' '.join([f'{b:02X}' for b in data])
            display_text += hex_str
        
        if self.radio_ascii.isChecked():
            try:
                ascii_str = data.decode('utf-8', errors='ignore')
                if self.radio_hex.isChecked():
                    display_text += f"  [{ascii_str}]"
                else:
                    display_text += ascii_str
            except:
                pass
        
        self.text_data_display.append(display_text)
        
        # 限制显示区域行数，防止内存溢出
        document = self.text_data_display.document()
        if document.blockCount() > 2000:
            cursor = QTextCursor(document)
            cursor.movePosition(QTextCursor.Start)
            # 选中前100行（批量删除以提高性能）
            for _ in range(100):
                cursor.movePosition(QTextCursor.NextBlock, QTextCursor.KeepAnchor)
            cursor.removeSelectedText()
        
        # 滚动到底部
        self.text_data_display.moveCursor(QTextCursor.End)
    
    def update_display_count(self, count):
        """更新显示的窗口数量"""
        for i in range(50):
            if i < count:
                self.display_widgets[i].show()
            else:
                self.display_widgets[i].hide()
    
    def update_preset_count(self, count):
        """更新显示的预设命令按钮数量"""
        for i in range(50):
            if i < count:
                self.preset_button_widgets[i].show()
            else:
                self.preset_button_widgets[i].hide()
    
    def update_bit_display_count(self, count):
        """更新显示的位窗口数量"""
        for i in range(50):
            if i < count:
                self.bit_display_widgets[i].show()
            else:
                self.bit_display_widgets[i].hide()
    
    def configure_custom_display(self, index):
        """配置自定义显示窗口"""
        display_data = None
        if index < len(self.custom_displays):
            display_data = self.custom_displays[index]
        
        dialog = CustomDisplayDialog(self, display_data, self.protocol_configs)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            
            # 更新或添加显示配置
            if index < len(self.custom_displays):
                self.custom_displays[index] = data
            else:
                while len(self.custom_displays) <= index:
                    self.custom_displays.append(None)
                self.custom_displays[index] = data
            
            # 更新显示
            self.update_custom_display(index)
    
    def update_custom_display(self, index, value=None):
        """更新自定义显示窗口
        index: 窗口索引
        value: 解析的数值,如果为None则显示配置信息
        """
        if index >= len(self.custom_displays) or not self.custom_displays[index]:
            self.custom_display_labels[index].setText(f'显示{index+1}: --')
            self.custom_display_labels[index].setStyleSheet("""
                QLabel {
                    background-color: #f0f0f0;
                    border: 2px solid #c0c0c0;
                    border-radius: 4px;
                    padding: 8px;
                    font-size: 11pt;
                    font-weight: bold;
                }
            """)
            return
        
        config = self.custom_displays[index]
        
        if not config.get('enabled', True):
            self.custom_display_labels[index].setText(f"{config.get('name', f'显示{index+1}')}: 已禁用")
            self.custom_display_labels[index].setStyleSheet("""
                QLabel {
                    background-color: #e0e0e0;
                    border: 2px solid #a0a0a0;
                    border-radius: 4px;
                    padding: 8px;
                    font-size: 11pt;
                    color: #888;
                }
            """)
            return
        
        if value is None:
            # 显示配置信息
            name = config.get('name', f'显示{index+1}')
            unit = config.get('unit', '')
            self.custom_display_labels[index].setText(f"{name}: -- {unit}")
            self.custom_display_labels[index].setStyleSheet("""
                QLabel {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f0f9eb, stop:1 #e1f3d8);
                    border: 1px solid #c2e7b0;
                    border-radius: 6px;
                    padding: 3px;
                    font-size: 7.5pt;
                    font-weight: bold;
                    color: #67c23a;
                }
            """)
        else:
            # 显示实时数值
            name = config.get('name', f'显示{index+1}')
            unit = config.get('unit', '')
            decimals = config.get('decimals', 2)
            
            if isinstance(value, str):
                value_str = value
            else:
                value_str = f"{value:.{decimals}f}"
                
            self.custom_display_labels[index].setText(f"{name}: {value_str} {unit}")
            self.custom_display_labels[index].setStyleSheet("""
                QLabel {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #67c23a, stop:1 #529b2e);
                    border: 1px solid #529b2e;
                    border-bottom: 2px solid #3e7523;
                    border-radius: 6px;
                    padding: 3px;
                    font-size: 8pt;
                    font-weight: bold;
                    color: white;
                }
            """)
    
    def configure_bit_display(self, window_idx):
        """配置位显示窗口"""
        bit_display_data = None
        if window_idx < len(self.bit_displays):
            bit_display_data = self.bit_displays[window_idx]
        
        dialog = BitDisplayDialog(self, bit_display_data, self.protocol_configs)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            
            # 更新或添加位显示配置
            self.bit_displays[window_idx] = data
            
            # 更新窗口显示（更新标题和位名称）
            self.update_bit_display_config(window_idx)
    
    def update_bit_display_config(self, window_idx):
        """更新位显示窗口的配置信息（标题和位名称）"""
        if window_idx >= len(self.bit_displays) or not self.bit_displays[window_idx]:
            return
        
        config = self.bit_displays[window_idx]
        bit_names = config.get('bit_names', [''] * 8)
        window_name = config.get('name', '')
        
        # 更新窗口名称标签
        if window_idx < len(self.bit_display_name_labels):
            if window_name:
                self.bit_display_name_labels[window_idx].setText(f'窗口名称: {window_name}')
            else:
                self.bit_display_name_labels[window_idx].setText(f'窗口名称: 未配置')
        
        # 更新每个位的名称标签
        for bit_idx in range(8):
            name_label, status_label = self.bit_display_labels[window_idx][bit_idx]
            bit_name = bit_names[bit_idx] if bit_names[bit_idx] else f'Bit{bit_idx}'
            name_label.setText(bit_name)
    
    def update_bit_display_value(self, window_idx, byte_value):
        """更新位显示窗口的值
        window_idx: 窗口索引
        byte_value: 字节值(0-255)
        """
        if window_idx >= len(self.bit_displays) or not self.bit_displays[window_idx]:
            return
        
        config = self.bit_displays[window_idx]
        if not config.get('enabled', True):
            return
        
        # 解析每个位并更新显示
        for bit_idx in range(8):
            bit_value = (byte_value >> bit_idx) & 1
            name_label, status_label = self.bit_display_labels[window_idx][bit_idx]
            
            # 根据位值设置颜色
            if bit_value == 1:
                status_label.setStyleSheet('color: #4CAF50; font-size: 12pt;')  # 绿色-开
            else:
                status_label.setStyleSheet('color: #ccc; font-size: 12pt;')  # 灰色-关
    
    def parse_custom_display_data(self, data, protocol_name=None):
        """解析数据帧并更新所有自定义显示窗口
        
        Args:
            data: 数据帧字节
            protocol_name: 协议名称，None表示默认协议
        """
        # 更新接收时钟 - 检查协议匹配
        if self.clock_config:
            clock_protocol = self.clock_config.get('protocol', None)
            if clock_protocol == protocol_name:
                self.update_received_clock(data)
        
        # 更新位显示窗口 - 检查协议匹配
        for window_idx, config in enumerate(self.bit_displays):
            if config and config.get('enabled', True):
                bit_protocol = config.get('protocol', None)
                if bit_protocol == protocol_name:
                    target_byte = config.get('target_byte', 17)
                    if len(data) > target_byte:
                        byte_value = data[target_byte]
                        self.update_bit_display_value(window_idx, byte_value)
        
        for i, config in enumerate(self.custom_displays):
            if not config or not config.get('enabled', True):
                continue
            
            # 检查协议是否匹配
            config_protocol = config.get('protocol', None)
            if config_protocol != protocol_name:
                continue
            
            try:
                start_byte = config.get('start_byte', 0)
                byte_count = config.get('byte_count', 2)
                data_type = config.get('data_type', 'uint16 (LE)')
                coefficient = config.get('coefficient', 1.0)
                divisor = config.get('divisor', 1.0)
                offset = config.get('offset', 0.0)
                
                # 检查数据长度
                if len(data) < start_byte + byte_count:
                    continue
                
                # 提取字节
                data_bytes = data[start_byte:start_byte + byte_count]
                
                # 根据数据类型解析
                value = self.parse_data_bytes(data_bytes, data_type)
                
                # 应用系数、除法和偏移: (value * coefficient / divisor) + offset
                if data_type != 'string (ASCII)':
                    value = (value * coefficient / divisor) + offset
                
                # 更新显示
                self.update_custom_display(i, value)
                
            except Exception as e:
                print(f"解析显示{i+1}数据错误: {e}")
    
    def parse_data_bytes(self, data_bytes, data_type):
        """解析字节数据为数值"""
        if data_type == 'uint8':
            return data_bytes[0]
        elif data_type == 'int8':
            return struct.unpack('b', data_bytes[:1])[0]
        elif data_type == 'uint16 (LE)':
            return struct.unpack('<H', data_bytes[:2])[0]
        elif data_type == 'uint16 (BE)':
            return struct.unpack('>H', data_bytes[:2])[0]
        elif data_type == 'int16 (LE)':
            return struct.unpack('<h', data_bytes[:2])[0]
        elif data_type == 'int16 (BE)':
            return struct.unpack('>h', data_bytes[:2])[0]
        elif data_type == 'uint32 (LE)':
            return struct.unpack('<I', data_bytes[:4])[0]
        elif data_type == 'uint32 (BE)':
            return struct.unpack('>I', data_bytes[:4])[0]
        elif data_type == 'int32 (LE)':
            return struct.unpack('<i', data_bytes[:4])[0]
        elif data_type == 'int32 (BE)':
            return struct.unpack('>i', data_bytes[:4])[0]
        elif data_type == 'float (LE)':
            return struct.unpack('<f', data_bytes[:4])[0]
        elif data_type == 'float (BE)':
            return struct.unpack('>f', data_bytes[:4])[0]
        elif data_type == 'string (ASCII)':
            try:
                # 移除空字节并解码
                return data_bytes.rstrip(b'\x00').decode('ascii')
            except UnicodeDecodeError:
                return data_bytes.decode('ascii', errors='replace')
        else:
            return 0
    
    def send_data(self):
        """发送数据"""
        if not self.serial_port or not self.serial_port.is_open:
            QMessageBox.warning(self, "警告", "串口未打开")
            return
        
        text = self.text_send.toPlainText()
        if not text:
            return
        
        try:
            # 准备数据
            if self.check_send_hex.isChecked():
                # HEX模式
                hex_str = text.replace(' ', '').replace('\n', '').replace('\r', '')
                data = bytearray.fromhex(hex_str)
            else:
                # ASCII模式
                data = bytearray(text.encode('utf-8'))
            
            # 添加时间戳
            if self.check_send_timestamp.isChecked():
                timestamp = datetime.now().strftime('%Y%m%d%H%M%S').encode('ascii')
                data = bytearray(timestamp) + data
            
            # 添加CRC校验
            crc_type = self.combo_send_crc.currentText()
            data = self.add_crc_to_data(data, crc_type)
            
            # 发送数据
            self.serial_port.write(data)
            
            # 添加到显示区
            self.add_data_to_display('TX', bytes(data))
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发送失败: {str(e)}")
    
    def add_crc_to_data(self, data, crc_type):
        """给数据添加CRC校验"""
        if crc_type == '无':
            return data
        
        data_bytes = bytes(data)
        
        if crc_type == 'CCITT-CRC16':
            crc = CRCCalculator.calculate_ccitt_crc16(data_bytes)
            # 高字节在前
            data.append((crc >> 8) & 0xFF)
            data.append(crc & 0xFF)
        elif crc_type == 'Modbus-CRC16':
            crc = CRCCalculator.calculate_modbus_crc16(data_bytes)
            # Modbus标准:低字节在前,高字节在后
            data.append(crc & 0xFF)
            data.append((crc >> 8) & 0xFF)
        elif crc_type == 'CRC16-XMODEM':
            crc = CRCCalculator.calculate_crc16_xmodem(data_bytes)
            # 低字节在前,高字节在后
            data.append(crc & 0xFF)
            data.append((crc >> 8) & 0xFF)
        elif crc_type == '累加和':
            checksum = CRCCalculator.calculate_sum_check(data_bytes)
            data.append(checksum)
        elif crc_type == '异或':
            checksum = CRCCalculator.calculate_xor_check(data_bytes)
            data.append(checksum)
        
        return data
    
    def send_file(self):
        """发送文件"""
        if not self.serial_port or not self.serial_port.is_open:
            QMessageBox.warning(self, "警告", "串口未打开")
            return
        
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件")
        if not file_path:
            return
        
        try:
            with open(file_path, 'rb') as f:
                data = f.read()
            
            self.serial_port.write(data)
            self.add_data_to_display('TX', data)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发送文件失败: {str(e)}")
    
    def configure_preset(self, index):
        """配置预设命令"""
        preset_data = None
        if index < len(self.preset_commands):
            preset_data = self.preset_commands[index]
        
        dialog = PresetCommandDialog(self, preset_data)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            
            # 更新或添加预设命令
            if index < len(self.preset_commands):
                # 停止旧的定时器
                if self.preset_timers[index].isActive():
                    self.preset_timers[index].stop()
                self.preset_commands[index] = data
            else:
                while len(self.preset_commands) <= index:
                    self.preset_commands.append(None)
                self.preset_commands[index] = data
            
            # 更新按钮文本
            if data['name']:
                self.preset_buttons[index].setText(data['name'])
                self.preset_buttons[index].setStyleSheet("""
                    QPushButton {
                        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #67c23a, stop:1 #529b2e);
                        border: 1px solid #529b2e;
                        border-bottom: 2px solid #3e7523;
                        border-radius: 6px;
                        color: white;
                        font-weight: bold;
                        font-size: 7pt;
                        padding-left: 2px;
                        text-align: center;
                    }
                    QPushButton:hover {
                        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #85ce61, stop:1 #67c23a);
                    }
                    QPushButton:pressed {
                        background-color: #529b2e;
                        border-top: 2px solid #3e7523;
                        border-bottom: none;
                    }
                """)
            else:
                self.preset_buttons[index].setText(f'预设{index+1}')
                self.preset_buttons[index].setStyleSheet("""
                    QPushButton {
                        background-color: #f8f9fa;
                        border: 1px solid #dee2e6;
                        border-radius: 4px;
                        font-size: 7pt;
                        text-align: left;
                        padding-left: 10px;
                    }
                    QPushButton:hover {
                        background-color: #e0e0e0;
                    }
                """)
            
            # 启动周期发送
            if data['periodic'] and data['period'] > 0:
                if self.serial_port and self.serial_port.is_open:
                    # 使用round()确保精确到毫秒，避免截断误差
                    interval_ms = round(data['period'] * 1000)
                    self.preset_timers[index].start(interval_ms)
    
    def send_preset_command(self, index):
        """发送预设命令"""
        if not self.serial_port or not self.serial_port.is_open:
            # 串口未打开，停止该预设命令的周期发送
            if index < len(self.preset_timers) and self.preset_timers[index].isActive():
                self.preset_timers[index].stop()
            return
        
        if index >= len(self.preset_commands) or not self.preset_commands[index]:
            self.configure_preset(index)
            return
        
        preset = self.preset_commands[index]
        command = preset.get('command', '')
        
        if not command:
            return
        
        try:
            # 准备数据
            if preset.get('is_hex', False):
                # HEX模式
                hex_str = command.replace(' ', '').replace('\n', '').replace('\r', '')
                data = bytearray.fromhex(hex_str)
            else:
                # ASCII模式
                data = bytearray(command.encode('utf-8'))
            
            # 添加时间戳
            if preset.get('add_timestamp', False):
                timestamp = datetime.now().strftime('%Y%m%d%H%M%S').encode('ascii')
                data = bytearray(timestamp) + data
            
            # 数据填充功能 - 在CRC校验之前替换数据
            if preset.get('data_fill_enabled', False):
                # 获取填充框的值
                fill_value = self.preset_data_spinboxes[index].value()
                data_min = preset.get('data_min', 0)
                data_max = preset.get('data_max', 65535)
                
                # 检查数据范围
                if data_min <= fill_value <= data_max:
                    # 检查数据长度是否足够
                    if len(data) >= 3:
                        # 替换第2字节和第3字节 (高字节在前，低字节在后)
                        data[1] = (fill_value >> 8) & 0xFF  # 高字节
                        data[2] = fill_value & 0xFF  # 低字节
                    else:
                        QMessageBox.warning(self, "警告", "命令数据长度不足，无法填充数据")
                        return
                else:
                    QMessageBox.warning(self, "警告", f"填充数据{fill_value}超出范围[{data_min}, {data_max}]")
                    return
            
            # 添加CRC校验
            crc_type = preset.get('crc_type', '无')
            data = self.add_crc_to_data(data, crc_type)
            
            # 发送数据
            self.serial_port.write(data)
            
            # 添加到显示区
            self.add_data_to_display('TX', bytes(data))
            
        except Exception as e:
            # 发送失败，停止该预设命令的周期发送
            if index < len(self.preset_timers) and self.preset_timers[index].isActive():
                self.preset_timers[index].stop()
            print(f"发送预设命令失败: {str(e)}")
            # 不显示弹窗，避免周期发送时频繁弹窗
    
    def export_to_excel(self):
        """导出曲线数据到Excel"""
        df = self.plot_canvas.get_data_frame()
        if df is None or len(df) == 0:
            QMessageBox.warning(self, "警告", "没有数据可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(self, "保存Excel文件", 
                                                   f"data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                                   "Excel Files (*.xlsx)")
        if not file_path:
            return
        
        try:
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "成功", f"数据已导出到:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
    
    def export_to_image(self):
        """导出曲线截图"""
        file_path, _ = QFileDialog.getSaveFileName(self, "保存图片", 
                                                   f"plot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                                                   "PNG Files (*.png)")
        if not file_path:
            return
        
        try:
            self.plot_canvas.fig.savefig(file_path, dpi=300, bbox_inches='tight')
            QMessageBox.information(self, "成功", f"图片已保存到:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存图片失败: {str(e)}")
    
    def export_raw_data(self):
        """导出原始数据"""
        if not self.raw_data_buffer:
            QMessageBox.warning(self, "警告", "没有原始数据可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(self, "保存原始数据", 
                                                   f"raw_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                                                   "Text Files (*.txt)")
        if not file_path:
            return
        
        try:
            with open(file_path, 'w') as f:
                for data in self.raw_data_buffer:
                    hex_str = ' '.join([f'{b:02X}' for b in data])
                    f.write(hex_str + '\n')
            
            QMessageBox.information(self, "成功", f"原始数据已导出到:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
    
    def clear_plot_data(self):
        """清空曲线数据"""
        reply = QMessageBox.question(self, "确认", "确定要清空曲线数据吗?",
                                    QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.plot_canvas.clear_data()
            self.raw_data_buffer.clear()
    
    def browse_auto_save_path(self):
        """浏览选择自动保存路径"""
        path = QFileDialog.getExistingDirectory(self, "选择自动保存路径", 
                                                 self.plot_canvas.auto_save_path or os.getcwd())
        if path:
            self.plot_canvas.set_auto_save_path(path)
            self.log(f"自动保存路径已设置为: {path}", "SUCCESS")
            QMessageBox.information(self, "成功", f"自动保存路径已设置为:\n{path}")
    
    def save_config(self):
        """保存业务配置（不包括窗口布局）"""
        config = {
            'port': self.combo_port.currentText(),
            'baudrate': self.combo_baudrate.currentText(),
            'databits': self.combo_databits.currentText(),
            'stopbits': self.combo_stopbits.currentText(),
            'parity': self.combo_parity.currentText(),
            'auto_reconnect': self.check_auto_reconnect.isChecked(),
            'send_hex': self.check_send_hex.isChecked(),
            'send_timestamp': self.check_send_timestamp.isChecked(),
            'send_crc': self.combo_send_crc.currentText(),
            'display_hex': self.radio_hex.isChecked(),
            'display_ascii': self.radio_ascii.isChecked(),
            'display_timestamp': self.check_display_timestamp.isChecked(),
            'preset_commands': self.preset_commands,
            'custom_displays': self.custom_displays,
            'bit_displays': self.bit_displays,  # 保存位显示窗口配置
            'display_count': self.spin_display_count.value(),
            'preset_count': self.spin_preset_count.value(),
            'clock_config': self.clock_config,  # 保存时钟配置
            'clock_splitter_sizes': self.clock_splitter.sizes(),  # 保存时钟分割器宽度
            'curve_configs': self.plot_canvas.curve_configs,  # 保存曲线配置
            'auto_save_path': self.plot_canvas.auto_save_path,  # 保存自动保存路径
            'frame_interval_ms': self.frame_interval_ms,  # 保存帧间隔设置
            'bit_display_count': self.spin_bit_display_count.value(),  # 保存位显示数量
            'background_image': self.background_image_path,  # 保存背景图片路径
            'background_opacity': self.background_opacity,  # 保存背景透明度
            'protocol_configs': self.protocol_configs  # 保存协议配置
        }
        
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            self.log("配置已保存", "SUCCESS")
            QMessageBox.information(self, "成功", "配置已保存到 serial_config.json")
        except Exception as e:
            print(f"保存配置失败: {e}")
            self.log(f"保存配置失败: {e}", "ERROR")
            QMessageBox.warning(self, "错误", f"保存配置失败: {e}")
    
    def load_config(self, restore_layout=False):
        """加载业务配置
        
        Args:
            restore_layout: 是否恢复窗口布局，默认False
        """
        if not os.path.exists(self.CONFIG_FILE):
            return
        
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 恢复配置
            index = self.combo_port.findText(config.get('port', ''))
            if index >= 0:
                self.combo_port.setCurrentIndex(index)
            
            self.combo_baudrate.setCurrentText(config.get('baudrate', '115200'))
            self.combo_databits.setCurrentText(config.get('databits', '8'))
            self.combo_stopbits.setCurrentText(config.get('stopbits', '1'))
            self.combo_parity.setCurrentText(config.get('parity', 'None'))
            
            # 更新串口参数显示标签
            if hasattr(self, 'label_serial_params'):
                self.update_serial_params_label()
            
            self.check_auto_reconnect.setChecked(config.get('auto_reconnect', False))
            self.check_send_hex.setChecked(config.get('send_hex', False))
            self.check_send_timestamp.setChecked(config.get('send_timestamp', False))
            
            # 恢复发送CRC配置
            send_crc = config.get('send_crc', '无')
            index = self.combo_send_crc.findText(send_crc)
            if index >= 0:
                self.combo_send_crc.setCurrentIndex(index)
            
            self.radio_hex.setChecked(config.get('display_hex', True))
            self.radio_ascii.setChecked(config.get('display_ascii', False))
            self.check_display_timestamp.setChecked(config.get('display_timestamp', True))
            
            # 加载曲线配置
            if 'curve_configs' in config:
                curve_configs = config['curve_configs']
                if curve_configs and hasattr(self, 'plot_canvas'):
                    try:
                        self.plot_canvas.set_curve_configs(curve_configs)
                        self.update_curve_checkbox_labels()
                    except Exception as e:
                        print(f"加载曲线配置错误: {e}")
            
            # 加载预设命令
            self.preset_commands = config.get('preset_commands', [])
            for i, preset in enumerate(self.preset_commands):
                if preset and i < len(self.preset_buttons):
                    if preset.get('name'):
                        self.preset_buttons[i].setText(preset['name'])
                        self.preset_buttons[i].setStyleSheet("""
                            QPushButton {
                                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #67c23a, stop:1 #529b2e);
                                border: 1px solid #529b2e;
                                border-bottom: 2px solid #3e7523;
                                border-radius: 6px;
                                color: white;
                                font-weight: bold;
                                font-size: 7pt;
                                padding-left: 2px;
                                text-align: center;
                            }
                            QPushButton:hover {
                                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #85ce61, stop:1 #67c23a);
                            }
                            QPushButton:pressed {
                                background-color: #529b2e;
                                border-top: 2px solid #3e7523;
                                border-bottom: none;
                            }
                        """)
            
            # 加载自定义显示窗口配置
            loaded_displays = config.get('custom_displays', [])
            # 确保列表长度为50
            self.custom_displays = [None] * 50
            for i, display_config in enumerate(loaded_displays):
                if i < 15:
                    self.custom_displays[i] = display_config
            
            for i, display_config in enumerate(self.custom_displays):
                if display_config and i < len(self.custom_display_labels):
                    if display_config.get('enabled', True):
                        name = display_config.get('name', f'显示窗口{i+1}')
                        self.custom_display_labels[i].setText(f"{name}: --")
            
            # 加载位显示窗口配置
            loaded_bit_displays = config.get('bit_displays', [])
            for i, bit_display_config in enumerate(loaded_bit_displays):
                if i < 2 and bit_display_config:
                    self.bit_displays[i] = bit_display_config
                    # 更新配置显示
                    self.update_bit_display_config(i)
            
            # 恢复显示窗口数量设置
            if 'display_count' in config:
                self.spin_display_count.setValue(config['display_count'])
            
            # 恢复预设命令显示数量设置
            if 'preset_count' in config:
                self.spin_preset_count.setValue(config['preset_count'])
            
            # 恢复位显示窗口数量设置
            if 'bit_display_count' in config:
                self.spin_bit_display_count.setValue(config['bit_display_count'])
            
            # 加载时钟配置
            if 'clock_config' in config:
                self.clock_config = config['clock_config']
            
            # 恢复自动保存路径
            if 'auto_save_path' in config:
                auto_save_path = config['auto_save_path']
                self.plot_canvas.set_auto_save_path(auto_save_path)
            
            # 恢复帧间隔设置
            if 'frame_interval_ms' in config:
                self.frame_interval_ms = config['frame_interval_ms']
                self.spin_frame_interval.setValue(self.frame_interval_ms)
            
            # 恢复背景设置
            if 'background_image' in config:
                self.background_image_path = config['background_image']
            if 'background_opacity' in config:
                self.background_opacity = config['background_opacity']
            self.update_background()
            
            # 恢复协议配置
            if 'protocol_configs' in config:
                self.protocol_configs = config['protocol_configs']
                self.init_protocol_parsers()
            
            # 根据参数决定是否恢复窗口布局
            if restore_layout:
                self.load_window_layout()
            
        except Exception as e:
            print(f"加载配置失败: {e}")
    
    def export_config_file(self):
        """导出配置文件"""
        try:
            # 选择保存位置
            filename, _ = QFileDialog.getSaveFileName(
                self,
                "导出配置文件",
                f"serial_config_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                "JSON文件 (*.json)"
            )
            
            if filename:
                # 读取当前配置文件
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config_data = f.read()
                
                # 写入到选定位置
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(config_data)
                
                QMessageBox.information(self, '成功', f'配置已导出到:\n{filename}')
        
        except Exception as e:
            QMessageBox.critical(self, '错误', f'导出配置失败:\n{str(e)}')
    
    def export_window_layout_file(self):
        """导出窗口布局到文件"""
        try:
            # 先保存当前窗口布局
            self.save_window_layout()
            
            # 检查布局文件是否存在
            if not os.path.exists(self.LAYOUT_FILE):
                QMessageBox.warning(self, '警告', '窗口布局文件不存在，无法导出')
                return
            
            # 选择保存位置
            filename, _ = QFileDialog.getSaveFileName(
                self,
                "导出窗口布局",
                f"window_layout_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                "JSON文件 (*.json)"
            )
            
            if filename:
                # 复制布局文件到目标位置
                import shutil
                shutil.copy2(self.LAYOUT_FILE, filename)
                QMessageBox.information(self, '成功', 
                    f'窗口布局已导出到:\n{filename}\n\n'
                    f'您可以在其他设备上使用"导入窗口布局"功能恢复此布局')
        
        except Exception as e:
            QMessageBox.critical(self, '错误', f'导出窗口布局失败:\n{str(e)}')
    
    def import_window_layout_file(self):
        """从文件导入窗口布局"""
        try:
            # 选择布局文件
            filename, _ = QFileDialog.getOpenFileName(
                self,
                "导入窗口布局",
                "",
                "JSON文件 (*.json)"
            )
            
            if filename:
                # 读取布局文件
                with open(filename, 'r', encoding='utf-8') as f:
                    layout_data = json.load(f)
                
                # 验证布局文件格式
                if not isinstance(layout_data, dict):
                    raise ValueError("布局文件格式不正确")
                
                # 备份当前布局
                if os.path.exists(self.LAYOUT_FILE):
                    backup_layout_file = self.LAYOUT_FILE + '.backup'
                    import shutil
                    shutil.copy2(self.LAYOUT_FILE, backup_layout_file)
                
                # 写入新布局
                with open(self.LAYOUT_FILE, 'w', encoding='utf-8') as f:
                    json.dump(layout_data, f, indent=4, ensure_ascii=False)
                
                # 询问是否立即应用
                reply = QMessageBox.question(
                    self,
                    '应用布局',
                    '窗口布局已导入！\n\n'
                    '是否立即应用新布局？\n\n'
                    '选择"是"：立即应用新布局\n'
                    '选择"否"：下次启动时生效',
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                
                if reply == QMessageBox.Yes:
                    # 立即应用布局
                    self.load_window_layout()
                    QMessageBox.information(self, '成功', '窗口布局已导入并应用！')
                else:
                    QMessageBox.information(self, '成功', 
                        '窗口布局已导入！\n'
                        '将在下次启动程序时自动应用')
        
        except Exception as e:
            QMessageBox.critical(self, '错误', f'导入窗口布局失败:\n{str(e)}')
    
    def import_config_file(self):
        """导入配置文件（不包含窗口布局）"""
        backup_file = None  # 初始化backup_file变量
        try:
            # 选择配置文件
            filename, _ = QFileDialog.getOpenFileName(
                self,
                "导入配置文件",
                "",
                "JSON文件 (*.json)"
            )
            
            if filename:
                # 读取配置文件
                with open(filename, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 验证配置文件格式
                if not isinstance(config, dict):
                    raise ValueError("配置文件格式不正确")
                
                # 备份当前配置
                if os.path.exists(self.CONFIG_FILE):
                    backup_file = self.CONFIG_FILE + '.backup'
                    with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                        backup_data = f.read()
                    with open(backup_file, 'w', encoding='utf-8') as f:
                        f.write(backup_data)
                
                # 写入新配置
                with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4, ensure_ascii=False)
                
                # 导入配置，不恢复窗口布局
                

                self.load_config(restore_layout=False)
                
                QMessageBox.information(self, '成功', 
                    f'配置已导入并生效!\n'
                    f'窗口布局保持不变。\n\n'
                    f'原配置已备份至:\n{backup_file}\n\n'
                    f'如需恢复窗口布局，请使用"导入窗口布局"功能。')
        
        except Exception as e:
            QMessageBox.critical(self, '错误', f'导入配置失败:\n{str(e)}')
    
    def log(self, message, level='INFO'):
        """添加日志信息
        level: INFO, WARNING, ERROR, SUCCESS
        """
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        
        # 根据级别设置颜色
        color_map = {
            'INFO': '#4fc3f7',      # 蓝色
            'WARNING': '#ffb74d',   # 橙色
            'ERROR': '#e57373',     # 红色
            'SUCCESS': '#81c784',   # 绿色
            'DEBUG': '#9575cd'      # 紫色
        }
        color = color_map.get(level, '#d4d4d4')
        
        # 格式化日志
        log_html = f'<span style="color: #757575;">[{timestamp}]</span> '
        log_html += f'<span style="color: {color}; font-weight: bold;">[{level}]</span> '
        log_html += f'<span style="color: #d4d4d4;">{message}</span>'
        
        # 添加到日志窗口
        self.text_log.append(log_html)
        
        # 限制日志行数
        document = self.text_log.document()
        if document.blockCount() > 1000:
            cursor = QTextCursor(document)
            cursor.movePosition(QTextCursor.Start)
            for _ in range(50):
                cursor.movePosition(QTextCursor.NextBlock, QTextCursor.KeepAnchor)
            cursor.removeSelectedText()
        
        # 自动滚动到底部
        if self.check_auto_scroll.isChecked():
            cursor = self.text_log.textCursor()
            cursor.movePosition(QTextCursor.End)
            self.text_log.setTextCursor(cursor)
    
    def clear_log(self):
        """清空日志"""
        self.text_log.clear()
        self.log("日志已清空", "INFO")
    
    def export_log(self):
        """导出日志到文件"""
        filename, _ = QFileDialog.getSaveFileName(
            self, 
            "导出日志", 
            f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "文本文件 (*.txt);;所有文件 (*.*)"
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    # 获取纯文本内容（去除HTML标签）
                    f.write(self.text_log.toPlainText())
                self.log(f"日志已导出到: {filename}", "SUCCESS")
            except Exception as e:
                self.log(f"导出日志失败: {str(e)}", "ERROR")
    
    def update_background(self):
        """更新窗口背景"""
        from PyQt5.QtGui import QPixmap, QPalette, QBrush, QPainter, QColor
        from PyQt5.QtCore import Qt
        
        if self.background_image_path and os.path.exists(self.background_image_path):
            try:
                # 创建新的调色板（不使用旧的）
                palette = QPalette()
                
                # 重新加载背景图片
                pixmap = QPixmap(self.background_image_path)
                
                if pixmap.isNull():
                    if hasattr(self, 'text_log'):
                        self.log(f"背景图片加载失败: {self.background_image_path}", "ERROR")
                    return
                
                # 缩放到窗口大小
                pixmap = pixmap.scaled(self.size(), Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
                
                # 创建半透明效果 - 关键：先在白色背景上应用透明度
                result_pixmap = QPixmap(pixmap.size())
                result_pixmap.fill(QColor(240, 245, 250))  # 浅色背景
                
                painter = QPainter(result_pixmap)
                painter.setOpacity(self.background_opacity)
                painter.drawPixmap(0, 0, pixmap)
                painter.end()
                
                # 使用 Window 角色设置背景
                brush = QBrush(result_pixmap)
                palette.setBrush(QPalette.Window, brush)
                self.setPalette(palette)
                
                # 必须设置为 True 才能显示背景
                self.setAutoFillBackground(True)
                
                # 强制重绘
                self.update()
                
                # 只在日志窗口存在时输出日志
                if hasattr(self, 'text_log'):
                    self.log(f"背景已更新: {os.path.basename(self.background_image_path)}, 透明度: {self.background_opacity:.1%}", "SUCCESS")
                    
            except Exception as e:
                if hasattr(self, 'text_log'):
                    self.log(f"背景更新失败: {str(e)}", "ERROR")
                import traceback
                traceback.print_exc()
        else:
            # 无背景图片，使用默认浅色背景
            palette = QPalette()
            from PyQt5.QtGui import QColor, QBrush
            from PyQt5.QtCore import Qt
            # 设置默认浅蓝色背景
            palette.setBrush(QPalette.Window, QBrush(QColor(240, 248, 255)))
            self.setPalette(palette)
            self.setAutoFillBackground(True)
    
    def choose_background_image(self):
        """选择背景图片"""
        from PyQt5.QtWidgets import QMessageBox, QPushButton
        
        # 创建选择对话框
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("背景图片设置")
        msgBox.setText("请选择背景图片来源：")
        
        # 添加按钮
        btn_file = msgBox.addButton("选择本地文件", QMessageBox.ActionRole)
        btn_generate = msgBox.addButton("生成科幻背景", QMessageBox.ActionRole)
        msgBox.addButton("取消", QMessageBox.RejectRole)
        
        msgBox.exec_()
        clicked_button = msgBox.clickedButton()
        
        if clicked_button == btn_file:
            # 选择本地文件
            filename, _ = QFileDialog.getOpenFileName(
                self,
                "选择背景图片",
                "",
                "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif);;所有文件 (*.*)"
            )
            if filename:
                self.background_image_path = filename
                self.update_background()
        
        elif clicked_button == btn_generate:
            # 生成科幻背景
            self.generate_scifi_background()
    
    def clear_background_image(self):
        """清除背景图片"""
        self.background_image_path = ''
        self.update_background()
        self.log("背景图片已清除", "INFO")
    
    def adjust_background_opacity(self):
        """调整背景透明度"""
        from PyQt5.QtWidgets import QInputDialog
        
        value, ok = QInputDialog.getInt(
            self,
            "调整透明度",
            "设置背景透明度 (0-100):",
            int(self.background_opacity * 100),
            0, 100, 5
        )
        if ok:
            self.background_opacity = value / 100.0
            self.update_background()
    
    def show_about_dialog(self):
        """显示关于对话框"""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QDialogButtonBox
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QPixmap
        
        dialog = QDialog(self)
        dialog.setWindowTitle('关于')
        dialog.setMinimumWidth(500)
        
        layout = QVBoxLayout(dialog)
        layout.setSpacing(15)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Logo图标
        if os.path.exists('lihua_logo.png'):
            logo_label = QLabel()
            pixmap = QPixmap('lihua_logo.png')
            # 缩放到合适大小 (64x64)
            pixmap = pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            logo_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(logo_label)
        
        # 软件名称
        title_label = QLabel('力华亘金-上位机监控系统')
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #007AFF;
                padding: 10px;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 版本信息
        version_info = """
<div style="line-height: 1.8; font-size: 10pt;">
<p><b>版本号（Version）：</b>V1.12</p>
<p><b>发布日期：</b>2025-11-22</p>
<p><b>作者（Author）：</b>Liwei</p>
<p><b>版权所有（Copyright）：</b>力华亘金科技有限公司</p>
</div>
        """
        
        version_label = QLabel(version_info)
        version_label.setTextFormat(Qt.RichText)
        version_label.setWordWrap(True)
        version_label.setStyleSheet("""
            QLabel {
                background-color: rgba(255, 255, 255, 0.5);
                border: 1px solid rgba(0, 0, 0, 0.1);
                border-radius: 8px;
                padding: 15px;
            }
        """)
        layout.addWidget(version_label)
        
        # 版权声明
        copyright_notice = """
<div style="line-height: 1.6; font-size: 9pt;">
<p><b>版权声明（Copyright Notice）：</b></p>
<p>本软件著作权归力华亘金科技有限公司所有，未经授权不得复制、传播或用于商业用途。</p>
</div>
        """
        
        copyright_label = QLabel(copyright_notice)
        copyright_label.setTextFormat(Qt.RichText)
        copyright_label.setWordWrap(True)
        copyright_label.setStyleSheet("""
            QLabel {
                background-color: rgba(255, 248, 220, 0.8);
                border: 1px solid rgba(255, 193, 7, 0.3);
                border-radius: 8px;
                padding: 12px;
                color: #856404;
            }
        """)
        layout.addWidget(copyright_label)
        
        # 联系方式
        contact_info = """
<div style="line-height: 1.6; font-size: 9pt;">
<p><b>联系方式（Contact）：</b></p>
<p>📧 Email：<a href="mailto:1013344248@qq.com" style="color: #007AFF; text-decoration: none;">1013344248@qq.com</a></p>
</div>
        """
        
        contact_label = QLabel(contact_info)
        contact_label.setTextFormat(Qt.RichText)
        contact_label.setWordWrap(True)
        contact_label.setOpenExternalLinks(True)
        contact_label.setStyleSheet("""
            QLabel {
                background-color: rgba(255, 255, 255, 0.5);
                border: 1px solid rgba(0, 0, 0, 0.1);
                border-radius: 8px;
                padding: 12px;
            }
        """)
        layout.addWidget(contact_label)
        
        # 免责声明
        disclaimer = """
<div style="line-height: 1.6; font-size: 8pt; color: #666;">
<p><b>免责声明（Disclaimer）：</b></p>
<p>本软件旨在提供设备监控、数据展示与系统管理等功能。软件已在典型测试环境下验证，但不保证在所有运行环境中的绝对稳定性和可靠性。</p>
<p>用户在使用本软件过程中产生的任何直接或间接损失，力华亘金科技有限公司不承担任何法律责任。用户应确保所监控设备与软件版本兼容，并自行做好数据备份与风险控制。</p>
</div>
        """
        
        disclaimer_label = QLabel(disclaimer)
        disclaimer_label.setTextFormat(Qt.RichText)
        disclaimer_label.setWordWrap(True)
        disclaimer_label.setStyleSheet("""
            QLabel {
                background-color: rgba(248, 249, 250, 0.9);
                border: 1px solid rgba(0, 0, 0, 0.08);
                border-radius: 8px;
                padding: 12px;
            }
        """)
        layout.addWidget(disclaimer_label)
        
        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        button_box.setStyleSheet("""
            QPushButton {
                min-width: 80px;
                min-height: 32px;
                font-size: 9pt;
            }
        """)
        layout.addWidget(button_box)
        
        dialog.exec_()
    
    def generate_scifi_background(self):
        """生成科幻风格背景"""
        try:
            from PyQt5.QtGui import QImage, QPainter, QLinearGradient, QRadialGradient, QPen, QColor
            from PyQt5.QtCore import QPointF, QRectF
            import random
            
            # 创建图像
            width, height = 1920, 1080
            image = QImage(width, height, QImage.Format_RGB32)
            
            painter = QPainter(image)
            painter.setRenderHint(QPainter.Antialiasing)
            
            # 背景渐变 - 深蓝到黑色
            gradient = QLinearGradient(0, 0, 0, height)
            gradient.setColorAt(0, QColor(10, 20, 40))
            gradient.setColorAt(0.5, QColor(5, 10, 25))
            gradient.setColorAt(1, QColor(0, 5, 15))
            painter.fillRect(0, 0, width, height, gradient)
            
            # 添加发光圆圈（光晕效果）
            for _ in range(15):
                x = random.randint(0, width)
                y = random.randint(0, height)
                radius = random.randint(100, 300)
                
                radial_gradient = QRadialGradient(QPointF(x, y), radius)
                r, g, b = random.choice([(0, 122, 255), (52, 199, 89), (255, 59, 48), (175, 82, 222)])
                radial_gradient.setColorAt(0, QColor(r, g, b, 100))
                radial_gradient.setColorAt(0.5, QColor(r, g, b, 30))
                radial_gradient.setColorAt(1, QColor(r, g, b, 0))
                
                painter.setBrush(radial_gradient)
                painter.setPen(Qt.NoPen)
                painter.drawEllipse(QPointF(x, y), radius, radius)
            
            # 添加网格线
            pen = QPen(QColor(0, 122, 255, 40))
            pen.setWidth(1)
            painter.setPen(pen)
            
            grid_spacing = 100
            for x in range(0, width, grid_spacing):
                painter.drawLine(x, 0, x, height)
            for y in range(0, height, grid_spacing):
                painter.drawLine(0, y, width, y)
            
            # 添加随机光点（星星效果）
            painter.setPen(Qt.NoPen)
            for _ in range(200):
                x = random.randint(0, width)
                y = random.randint(0, height)
                size = random.randint(1, 3)
                alpha = random.randint(100, 255)
                painter.setBrush(QColor(255, 255, 255, alpha))
                painter.drawEllipse(QPointF(x, y), size, size)
            
            # 添加发光线条
            pen = QPen(QColor(0, 122, 255, 60))
            pen.setWidth(2)
            painter.setPen(pen)
            for _ in range(20):
                x1 = random.randint(0, width)
                y1 = random.randint(0, height)
                x2 = x1 + random.randint(-200, 200)
                y2 = y1 + random.randint(-200, 200)
                painter.drawLine(x1, y1, x2, y2)
            
            painter.end()
            
            # 保存图像
            bg_path = os.path.join(os.path.dirname(self.CONFIG_FILE), 'scifi_background.png')
            image.save(bg_path)
            
            self.background_image_path = bg_path
            self.update_background()
            
            self.log("科幻背景已生成", "SUCCESS")
            
        except Exception as e:
            self.log(f"生成背景失败: {str(e)}", "ERROR")
            import traceback
            traceback.print_exc()
    
    def resizeEvent(self, event):
        """窗口大小改变时重新应用背景"""
        super().resizeEvent(event)
        if self.background_image_path:
            self.update_background()
        
        # 窗口调整大小后延迟保存布局
        if not self.is_restoring_layout and hasattr(self, 'layout_save_timer'):
            self.layout_save_timer.stop()
            self.layout_save_timer.start(2000)  # 2秒后保存
    
    def changeEvent(self, event):
        """窗口状态改变事件（最小化、最大化等）"""
        super().changeEvent(event)
        # 当窗口从最小化恢复或最大化/正常化切换时，延迟保存布局
        if event.type() == event.WindowStateChange:
            # 不在恢复布局时保存，避免循环
            if not self.is_restoring_layout:
                # 使用延迟保存，避免在状态变化过程中保存不完整的布局
                if hasattr(self, 'layout_save_timer'):
                    self.layout_save_timer.stop()
                    self.layout_save_timer.start(1000)  # 1秒后保存
    
    def init_protocol_parsers(self):
        """初始化协议解析器"""
        self.protocol_parsers = []
        for config in self.protocol_configs:
            if config and config.get('enabled', True):
                parser = GenericProtocolParser(config)
                self.protocol_parsers.append(parser)
    
    def open_protocol_manager(self):
        """打开协议管理器窗口"""
        dialog = QDialog(self)
        dialog.setWindowTitle('接收协议管理')
        dialog.setMinimumSize(800, 500)
        
        layout = QVBoxLayout(dialog)
        
        # 说明标签
        info_label = QLabel('配置接收数据协议，支持多种帧格式和CRC校验')
        info_label.setStyleSheet('color: #666; padding: 5px;')
        layout.addWidget(info_label)
        
        # 协议列表
        list_widget = QListWidget()
        list_widget.setStyleSheet("""
            QListWidget {
                background: white;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:selected {
                background: #e3f2fd;
                color: #1976d2;
            }
        """)
        
        def update_list():
            list_widget.clear()
            for i, proto in enumerate(self.protocol_configs):
                if proto:
                    status = '✓' if proto.get('enabled', True) else '✗'
                    name = proto.get('name', '未命名')
                    length = proto.get('length', 0)
                    crc = proto.get('crc_type', '无')
                    list_widget.addItem(f"{status} {name} - {length}字节 - CRC:{crc}")
        
        update_list()
        layout.addWidget(list_widget)
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        
        btn_add = QPushButton('添加协议')
        btn_add.setStyleSheet("""
            QPushButton {
                background: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #45a049;
            }
        """)
        
        btn_edit = QPushButton('编辑')
        btn_edit.setStyleSheet("""
            QPushButton {
                background: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #0b7dda;
            }
        """)
        
        btn_delete = QPushButton('删除')
        btn_delete.setStyleSheet("""
            QPushButton {
                background: #f44336;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #da190b;
            }
        """)
        
        btn_close = QPushButton('关闭')
        btn_close.setStyleSheet("""
            QPushButton {
                background: #757575;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #616161;
            }
        """)
        
        def add_protocol():
            dialog_proto = ProtocolConfigDialog(self)
            if dialog_proto.exec_() == QDialog.Accepted:
                data = dialog_proto.get_data()
                if data:
                    self.protocol_configs.append(data)
                    update_list()
                    self.init_protocol_parsers()
                    self.log(f"已添加协议: {data['name']}", "SUCCESS")
        
        def edit_protocol():
            current_row = list_widget.currentRow()
            if current_row >= 0 and current_row < len(self.protocol_configs):
                dialog_proto = ProtocolConfigDialog(self, self.protocol_configs[current_row])
                if dialog_proto.exec_() == QDialog.Accepted:
                    data = dialog_proto.get_data()
                    if data:
                        self.protocol_configs[current_row] = data
                        update_list()
                        self.init_protocol_parsers()
                        self.log(f"已更新协议: {data['name']}", "SUCCESS")
        
        def delete_protocol():
            current_row = list_widget.currentRow()
            if current_row >= 0 and current_row < len(self.protocol_configs):
                proto_name = self.protocol_configs[current_row].get('name', '未命名')
                reply = QMessageBox.question(dialog, '确认删除',
                                            f'确定要删除协议 "{proto_name}" 吗？',
                                            QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    del self.protocol_configs[current_row]
                    update_list()
                    self.init_protocol_parsers()
                    self.log(f"已删除协议: {proto_name}", "WARNING")
        
        btn_add.clicked.connect(add_protocol)
        btn_edit.clicked.connect(edit_protocol)
        btn_delete.clicked.connect(delete_protocol)
        btn_close.clicked.connect(dialog.accept)
        
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_close)
        
        layout.addLayout(btn_layout)
        
        dialog.exec_()
    
    def closeEvent(self, event):
        """程序关闭"""
        # 不自动保存配置和布局，需要用户手动保存
        
        # 关闭串口和定时器
        if self.serial_port and self.serial_port.is_open:
            self.close_serial()
        
        # 停止恢复检测定时器
        self.stop_recovery_check()
        
        event.accept()


if __name__ == '__main__':
    # 启用高DPI缩放
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格,界面更美观
    
    # 创建并显示启动画面
    splash = SplashScreenWidget()
    splash.show()
    splash.update_progress(5, '正在启动应用程序...')
    app.processEvents()
    time.sleep(0.2)  # 短暂延迟让启动画面显示
    
    # 创建主窗口（这会触发各个阶段的进度更新）
    window = SerialMonitorApp(splash=splash)
    
    # 更新进度到100%
    splash.update_progress(100, '启动完成！')
    app.processEvents()
    time.sleep(0.3)  # 让用户看到完成状态
    
    # 显示主窗口并关闭启动画面
    window.show()
    splash.finish(window)
    
    sys.exit(app.exec_())
