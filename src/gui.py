"""
PDF转Word软件的GUI界面
"""

import os
import sys
import traceback
from typing import List, Optional, Tuple, Dict
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog, 
                             QLabel, QProgressBar, QVBoxLayout, QHBoxLayout, 
                             QWidget, QComboBox, QSpinBox, QCheckBox, QMessageBox,
                             QListWidget, QGroupBox, QTabWidget, QSlider)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QFont, QDragEnterEvent, QDropEvent

from .converter import PDFConverter


class WorkerThread(QThread):
    """工作线程，用于后台处理PDF转换任务"""
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, converter: PDFConverter, files: List[str], output_dir: str, 
                 quality: str, use_ocr: bool, pages: Optional[List[int]] = None):
        super().__init__()
        self.converter = converter
        self.files = files
        self.output_dir = output_dir
        self.quality = quality
        self.use_ocr = use_ocr
        self.pages = pages
        
    def run(self):
        try:
            successful_files = []
            failed_files = []
            
            for i, pdf_file in enumerate(self.files):
                file_name = os.path.basename(pdf_file)
                progress = int((i / len(self.files)) * 100)
                self.progress_signal.emit(progress, f"正在转换: {file_name}")
                
                output_file = os.path.join(
                    self.output_dir, 
                    os.path.splitext(file_name)[0] + ".docx"
                )
                
                try:
                    # 检查文件是否存在
                    if not os.path.exists(pdf_file):
                        raise FileNotFoundError(f"文件不存在: {pdf_file}")
                    
                    # 检查文件是否是PDF
                    if not pdf_file.lower().endswith('.pdf'):
                        raise ValueError(f"文件不是PDF格式: {pdf_file}")
                    
                    # 转换PDF到Word
                    self.converter.convert_pdf_to_word(
                        pdf_file, 
                        output_file, 
                        quality=self.quality,
                        use_ocr=self.use_ocr,
                        pages=self.pages
                    )
                    
                    successful_files.append(file_name)
                except FileNotFoundError as e:
                    failed_files.append(f"{file_name}: 文件不存在")
                except ImportError as e:
                    # 依赖问题
                    failed_files.append(f"{file_name}: {str(e)}")
                    self.finished_signal.emit(False, f"缺少必要的依赖项: {str(e)}")
                    return
                except ValueError as e:
                    # 格式或参数问题
                    failed_files.append(f"{file_name}: {str(e)}")
                except Exception as e:
                    # 其他错误
                    error_msg = f"{file_name}: 未知错误 - {str(e)}"
                    failed_files.append(error_msg)
            
            # 更新进度为100%
            self.progress_signal.emit(100, "转换完成！")
            
            # 构建完成消息
            if len(successful_files) == len(self.files):
                # 全部成功
                self.finished_signal.emit(True, f"所有 {len(successful_files)} 个PDF文件已成功转换！")
            elif len(successful_files) > 0:
                # 部分成功
                success_msg = f"成功转换 {len(successful_files)}/{len(self.files)} 个文件。\n\n"
                fail_msg = "转换失败的文件:\n" + "\n".join(failed_files)
                self.finished_signal.emit(False, success_msg + fail_msg)
            else:
                # 全部失败
                fail_msg = "所有文件转换失败:\n" + "\n".join(failed_files)
                self.finished_signal.emit(False, fail_msg)
                
        except Exception as e:
            # 捕获所有异常
            error_details = traceback.format_exc()
            self.finished_signal.emit(False, f"转换过程中出现意外错误: {str(e)}\n\n详细信息: {error_details}")


class MainWindow(QMainWindow):
    """主窗口类"""
    def __init__(self):
        super().__init__()
        
        self.pdf_files: List[str] = []
        self.output_directory: str = os.path.expanduser("~/Documents")
        self.converter = PDFConverter()
        
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("PDF转Word软件")
        self.setMinimumSize(800, 600)
        
        # 创建主部件和布局
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        
        # 创建标题标签
        title_label = QLabel("PDF转Word转换工具")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)
        
        # 创建选项卡
        tab_widget = QTabWidget()
        
        # 转换选项卡
        convert_tab = QWidget()
        convert_layout = QVBoxLayout(convert_tab)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout(file_group)
        
        select_btn_layout = QHBoxLayout()
        self.select_files_btn = QPushButton("选择PDF文件")
        self.select_files_btn.clicked.connect(self.select_pdf_files)
        self.select_output_btn = QPushButton("选择输出目录")
        self.select_output_btn.clicked.connect(self.select_output_directory)
        select_btn_layout.addWidget(self.select_files_btn)
        select_btn_layout.addWidget(self.select_output_btn)
        file_layout.addLayout(select_btn_layout)
        
        self.file_list = QListWidget()
        self.file_list.setAcceptDrops(True)
        self.file_list.setMinimumHeight(150)
        file_layout.addWidget(QLabel("已选择的PDF文件:"))
        file_layout.addWidget(self.file_list)
        
        # 添加清除文件按钮
        clear_btn_layout = QHBoxLayout()
        self.clear_files_btn = QPushButton("清除文件列表")
        self.clear_files_btn.clicked.connect(self.clear_file_list)
        clear_btn_layout.addStretch()
        clear_btn_layout.addWidget(self.clear_files_btn)
        file_layout.addLayout(clear_btn_layout)
        
        self.output_label = QLabel(f"输出目录: {self.output_directory}")
        file_layout.addWidget(self.output_label)
        
        convert_layout.addWidget(file_group)
        
        # 转换选项区域
        options_group = QGroupBox("转换选项")
        options_layout = QVBoxLayout(options_group)
        
        quality_layout = QHBoxLayout()
        quality_layout.addWidget(QLabel("转换质量:"))
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["低质量 (快速)", "中等质量", "高质量 (慢速)"])
        self.quality_combo.setCurrentIndex(1)  # 默认中等质量
        quality_layout.addWidget(self.quality_combo)
        options_layout.addLayout(quality_layout)
        
        self.ocr_checkbox = QCheckBox("使用OCR识别文字(适用于扫描版PDF)")
        options_layout.addWidget(self.ocr_checkbox)
        
        # 添加OCR说明
        ocr_info = QLabel("注意: OCR功能需要安装Tesseract，并且转换速度会较慢")
        ocr_info.setStyleSheet("color: gray; font-style: italic;")
        options_layout.addWidget(ocr_info)
        
        page_select_layout = QHBoxLayout()
        self.specific_pages_checkbox = QCheckBox("仅转换特定页面")
        self.specific_pages_checkbox.stateChanged.connect(self.toggle_page_selection)
        page_select_layout.addWidget(self.specific_pages_checkbox)
        
        self.page_input = QWidget()
        page_input_layout = QHBoxLayout(self.page_input)
        page_input_layout.addWidget(QLabel("页码范围 (例如: 1-5,8,10-12):"))
        self.page_range_input = QComboBox()
        self.page_range_input.setEditable(True)
        self.page_range_input.setEnabled(False)
        page_input_layout.addWidget(self.page_range_input)
        self.page_input.setLayout(page_input_layout)
        
        page_select_layout.addWidget(self.page_input)
        options_layout.addLayout(page_select_layout)
        
        convert_layout.addWidget(options_group)
        
        # 进度和控制区域
        control_group = QGroupBox("转换控制")
        control_layout = QVBoxLayout(control_group)
        
        self.progress_bar = QProgressBar()
        control_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("准备就绪")
        control_layout.addWidget(self.status_label)
        
        convert_btn_layout = QHBoxLayout()
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.clicked.connect(self.start_conversion)
        convert_btn_layout.addWidget(self.convert_btn)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_conversion)
        convert_btn_layout.addWidget(self.cancel_btn)
        
        control_layout.addLayout(convert_btn_layout)
        
        convert_layout.addWidget(control_group)
        
        # 设置此选项卡的布局
        convert_tab.setLayout(convert_layout)
        
        # 关于选项卡
        about_tab = QWidget()
        about_layout = QVBoxLayout(about_tab)
        
        about_text = """
        <h2>PDF转Word转换软件</h2>
        <p>版本: 1.0.0</p>
        <p>这是一款功能强大的PDF转Word转换工具，能够:</p>
        <ul>
            <li>高质量转换PDF文件到Word文档(.docx)</li>
            <li>支持批量处理多个PDF文件</li>
            <li>允许选择特定页面进行转换</li>
            <li>通过OCR技术识别扫描版PDF中的文字</li>
            <li>最大程度保留原PDF的文本格式、图片和排版</li>
        </ul>
        <p>使用方法:</p>
        <ol>
            <li>选择一个或多个PDF文件</li>
            <li>选择输出目录</li>
            <li>设置转换选项</li>
            <li>点击"开始转换"按钮</li>
        </ol>
        <p>常见问题:</p>
        <ul>
            <li>如果转换失败，请尝试使用中等或低质量选项</li>
            <li>对于扫描版PDF，请确保安装了Tesseract OCR引擎</li>
            <li>如遇到"__enter__"相关错误，可能是PDF格式不兼容，请尝试使用OCR模式</li>
        </ul>
        """
        
        about_label = QLabel(about_text)
        about_label.setWordWrap(True)
        about_label.setOpenExternalLinks(True)
        about_label.setTextFormat(Qt.TextFormat.RichText)
        about_layout.addWidget(about_label)
        about_layout.addStretch()
        
        about_tab.setLayout(about_layout)
        
        # 添加选项卡到选项卡部件
        tab_widget.addTab(convert_tab, "转换")
        tab_widget.addTab(about_tab, "关于")
        
        main_layout.addWidget(tab_widget)
        
        # 设置中央部件
        self.setCentralWidget(main_widget)
        
        # 设置工作线程
        self.worker_thread = None
        
    def toggle_page_selection(self, state):
        """启用或禁用页面选择输入框"""
        self.page_range_input.setEnabled(state == Qt.CheckState.Checked.value)
        
    def select_pdf_files(self):
        """打开文件对话框选择PDF文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择PDF文件", "", "PDF文件 (*.pdf)"
        )
        
        if files:
            self.pdf_files.extend(files)
            self.update_file_list()
    
    def clear_file_list(self):
        """清除文件列表"""
        self.pdf_files = []
        self.update_file_list()
            
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_list.clear()
        for file in self.pdf_files:
            self.file_list.addItem(os.path.basename(file))
            
    def select_output_directory(self):
        """打开目录对话框选择输出目录"""
        directory = QFileDialog.getExistingDirectory(
            self, "选择输出目录", self.output_directory
        )
        
        if directory:
            self.output_directory = directory
            self.output_label.setText(f"输出目录: {self.output_directory}")
            
    def parse_page_range(self, page_range_str: str) -> List[int]:
        """解析页码范围字符串，返回页码列表"""
        pages = []
        if not page_range_str:
            return pages
            
        parts = page_range_str.split(',')
        for part in parts:
            if '-' in part:
                start, end = part.split('-')
                try:
                    start_int = int(start.strip())
                    end_int = int(end.strip())
                    pages.extend(range(start_int, end_int + 1))
                except ValueError:
                    continue
            else:
                try:
                    pages.append(int(part.strip()))
                except ValueError:
                    continue
                    
        return sorted(list(set(pages)))
        
    def start_conversion(self):
        """开始转换流程"""
        if not self.pdf_files:
            QMessageBox.warning(self, "警告", "请先选择PDF文件！")
            return
        
        # 检查输出目录是否存在
        if not os.path.exists(self.output_directory):
            try:
                os.makedirs(self.output_directory)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法创建输出目录: {str(e)}")
                return
            
        # 获取转换选项
        quality_map = {0: "low", 1: "medium", 2: "high"}
        quality = quality_map[self.quality_combo.currentIndex()]
        use_ocr = self.ocr_checkbox.isChecked()
        
        pages = None
        if self.specific_pages_checkbox.isChecked():
            page_range_str = self.page_range_input.currentText()
            pages = self.parse_page_range(page_range_str)
            if not pages:
                QMessageBox.warning(self, "警告", "页码范围格式无效！请使用格式如: 1-5,8,10-12")
                return
        
        # 创建并启动工作线程
        self.worker_thread = WorkerThread(
            self.converter, self.pdf_files, self.output_directory, 
            quality, use_ocr, pages
        )
        
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.conversion_finished)
        
        self.worker_thread.start()
        
        # 更新UI状态
        self.convert_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.status_label.setText("正在转换...")
        
    def cancel_conversion(self):
        """取消转换过程"""
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.terminate()
            self.worker_thread.wait()
            
            self.status_label.setText("转换已取消")
            self.progress_bar.setValue(0)
            self.convert_btn.setEnabled(True)
            self.cancel_btn.setEnabled(False)
            
    def update_progress(self, value: int, status: str):
        """更新进度条和状态标签"""
        self.progress_bar.setValue(value)
        self.status_label.setText(status)
        
    def conversion_finished(self, success: bool, message: str):
        """转换完成的处理"""
        self.convert_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        
        if success:
            QMessageBox.information(self, "完成", message)
            
            # 询问是否打开输出目录
            reply = QMessageBox.question(
                self, "打开文件夹", "是否打开输出目录查看转换结果？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.open_output_directory()
        else:
            # 创建详细的错误对话框
            error_dialog = QMessageBox(self)
            error_dialog.setIcon(QMessageBox.Icon.Critical)
            error_dialog.setWindowTitle("转换错误")
            error_dialog.setText("转换过程中出现错误")
            error_dialog.setDetailedText(message)
            error_dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
            error_dialog.exec()
            
        # 重置进度
        self.status_label.setText("准备就绪")
        
    def open_output_directory(self):
        """打开输出目录"""
        try:
            # 使用系统默认方式打开文件夹
            if sys.platform == 'darwin':  # macOS
                os.system(f'open "{self.output_directory}"')
            elif sys.platform == 'win32':  # Windows
                os.startfile(self.output_directory)
            else:  # Linux
                os.system(f'xdg-open "{self.output_directory}"')
        except Exception as e:
            QMessageBox.warning(self, "警告", f"无法打开输出目录: {str(e)}")


def run_gui():
    """运行GUI应用程序"""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 