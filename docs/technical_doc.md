# PDF转Word软件技术文档

## 项目概述

本文档提供了PDF转Word软件的详细技术设计和实现说明。该软件使用Python开发，提供图形用户界面，能够将PDF文件转换为Word文档(.docx)，支持批量处理、页面选择和OCR文字识别等功能。

## 系统架构

软件采用模块化设计，主要由以下组件构成：

### 核心组件

1. **GUI模块** (`src/gui.py`)
   - 实现用户界面
   - 处理用户交互
   - 提供转换进度反馈

2. **转换引擎** (`src/converter.py`)
   - 提供PDF到Word的核心转换功能
   - 支持不同质量级别的转换
   - 支持OCR文字识别

3. **主程序** (`src/main.py`)
   - 程序入口点
   - 协调各模块工作

### 依赖关系

```
run.py (启动脚本)
    |
    +--> src/main.py (主程序)
            |
            +--> src/gui.py (GUI界面)
            |       |
            |       +--> src/converter.py (转换器)
            |
            +--> src/converter.py (转换器)
```

## 技术实现

### 转换原理

#### 常规PDF转换

使用`pdf2docx`库实现基础的PDF到Word转换。转换流程：

1. 解析PDF文档结构
2. 提取文本、图像和表格
3. 保留原始格式生成docx文件

转换质量设置通过调整不同的参数实现：
- 低质量：关闭多处理和高级格式保留功能，追求速度
- 中等质量：平衡格式保留和处理速度
- 高质量：启用所有格式保留功能，优化表格和段落识别

#### OCR文字识别

对于扫描版PDF，使用以下流程：

1. 使用`pdf2image`将PDF页面转换为图像
2. 使用`pytesseract`对图像进行OCR文字识别
3. 使用`python-docx`创建包含识别文本的Word文档

### 多线程实现

为避免在转换大文件时界面冻结，转换任务运行在单独的工作线程中：

1. 创建`QThread`子类(`WorkerThread`)
2. 实现进度信号(`progress_signal`)和完成信号(`finished_signal`)
3. 通过信号机制更新GUI显示进度

### 内存管理

针对大型PDF文件，采取以下策略：

1. 每次只处理一个页面范围
2. 使用临时文件存储中间结果
3. 完成后清理临时文件

## 开发环境

- Python 3.8+
- PyQt6
- 核心依赖：
  - pdf2docx
  - pytesseract
  - pdf2image
  - PyPDF2
  - python-docx

## 已知限制

1. OCR识别准确度受图片质量影响
2. 复杂表格和特殊排版可能无法完全保留
3. 依赖Tesseract OCR引擎，需要额外安装

## 未来扩展方向

1. 支持更多输出格式(Excel, PPT等)
2. 添加云存储集成
3. 优化转换速度和质量
4. 添加自定义样式模板功能
5. 实现预览功能

## API文档

### PDFConverter类

核心转换类，提供PDF到Word转换功能。

#### 主要方法

```python
def convert_pdf_to_word(self, pdf_path: str, output_path: str, 
                       quality: str = "medium", use_ocr: bool = False,
                       pages: Optional[List[int]] = None) -> None:
    """
    将PDF文件转换为Word文档
    
    参数:
        pdf_path: PDF文件路径
        output_path: 输出Word文档路径
        quality: 转换质量，可选值为 "low", "medium", "high"
        use_ocr: 是否使用OCR识别文字
        pages: 要转换的页码列表，如果为None则转换所有页面
    """
```

#### 内部方法

- `_convert_without_ocr()`: 使用pdf2docx直接转换
- `_convert_with_ocr()`: 使用OCR技术转换
- `_get_convert_options()`: 获取转换选项

### MainWindow类

GUI主窗口类，处理用户交互。

#### 主要方法

- `select_pdf_files()`: 选择PDF文件
- `select_output_directory()`: 选择输出目录
- `start_conversion()`: 开始转换
- `parse_page_range()`: 解析页码范围 