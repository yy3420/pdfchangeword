"""
PDF转Word转换器核心实现
"""

import os
import tempfile
from typing import List, Optional, Dict, Any
from PyPDF2 import PdfReader
from pdf2docx import Converter as PdfToDocxConverter
import pytesseract
from PIL import Image
import logging
from tqdm import tqdm


class PDFConverter:
    """PDF转Word转换器类，提供PDF文件到Word文档的转换功能"""
    
    def __init__(self):
        """初始化转换器"""
        self.logger = self._setup_logger()
        
    def _setup_logger(self) -> logging.Logger:
        """设置日志记录器"""
        logger = logging.getLogger("pdf_converter")
        logger.setLevel(logging.INFO)
        
        # 创建控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # 创建格式化器
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(formatter)
        
        # 添加处理器到日志记录器
        logger.addHandler(console_handler)
        
        return logger
        
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
        self.logger.info(f"开始转换: {pdf_path} -> {output_path}")
        self.logger.info(f"转换质量: {quality}, 使用OCR: {use_ocr}, 页码: {pages}")
        
        try:
            # 检查PDF文件是否存在
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
                
            # 检查PDF文件是否可读
            try:
                pdf_reader = PdfReader(pdf_path)
                total_pages = len(pdf_reader.pages)
                self.logger.info(f"PDF总页数: {total_pages}")
            except Exception as e:
                raise ValueError(f"无法读取PDF文件: {str(e)}")
            
            # 根据用户选择执行转换
            if use_ocr:
                self._convert_with_ocr(pdf_path, output_path, quality, pages)
            else:
                self._convert_without_ocr(pdf_path, output_path, quality, pages)
                
            self.logger.info(f"转换完成: {output_path}")
        except Exception as e:
            self.logger.error(f"转换失败: {str(e)}")
            raise
            
    def _convert_without_ocr(self, pdf_path: str, output_path: str, 
                            quality: str, pages: Optional[List[int]]) -> None:
        """使用pdf2docx直接转换PDF到Word"""
        # 设置转换选项
        convert_options = self._get_convert_options(quality)
        
        # 确定页码范围
        pdf_reader = PdfReader(pdf_path)
        total_pages = len(pdf_reader.pages)
        
        # 如果指定了页码，则只转换指定页码
        if pages:
            # 过滤掉超出范围的页码
            valid_pages = [p - 1 for p in pages if 0 < p <= total_pages]  # 转为0索引
            page_ranges = []
            
            for page in valid_pages:
                page_ranges.append((page, page + 1))
        else:
            # 转换所有页码
            page_ranges = [(0, total_pages)]
            
        # 开始转换
        try:
            # 创建转换器实例
            cv = PdfToDocxConverter(pdf_path)
            
            # 对每个页面范围进行转换
            for start, end in page_ranges:
                self.logger.info(f"转换页面范围: {start+1} - {end}")
                cv.convert(output_path, start=start, end=end, **convert_options)
                
            # 关闭转换器
            cv.close()
        except Exception as e:
            self.logger.error(f"PDF转换异常: {str(e)}")
            # 尝试使用另一种方法
            try:
                self.logger.info("尝试使用替代方法转换...")
                # 使用更简单的方法进行转换，直接指定页面范围
                PdfToDocxConverter(pdf_path).convert(output_path, **convert_options)
            except Exception as second_e:
                self.logger.error(f"替代转换方法也失败: {str(second_e)}")
                raise ValueError(f"PDF转换失败，可能是文件格式不支持或者已被加密: {str(second_e)}")
                
    def _convert_with_ocr(self, pdf_path: str, output_path: str,
                         quality: str, pages: Optional[List[int]]) -> None:
        """使用OCR技术将PDF转换为Word"""
        from docx import Document
        from docx.shared import Pt
        
        # 创建一个新的Word文档
        doc = Document()
        
        # 设置文档样式
        style = doc.styles['Normal']
        style.font.name = '宋体'
        style.font.size = Pt(12)
        
        # 打开PDF文件
        pdf_reader = PdfReader(pdf_path)
        total_pages = len(pdf_reader.pages)
        
        # 确定要处理的页码
        if pages:
            # 过滤掉超出范围的页码
            process_pages = [p - 1 for p in pages if 0 < p <= total_pages]  # 转为0索引
        else:
            # 处理所有页码
            process_pages = list(range(total_pages))
            
        # 处理每个页面
        for i, page_num in enumerate(process_pages):
            # 打印进度
            self.logger.info(f"正在处理页面 {page_num + 1}/{total_pages}")
            
            # 提取页面
            page = pdf_reader.pages[page_num]
            
            # 将页面转换为图像
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                temp_image_path = temp_file.name
                
            # 这里需要使用pdf2image库将PDF页面转换为图像
            try:
                from pdf2image import convert_from_path
                images = convert_from_path(pdf_path, first_page=page_num+1, last_page=page_num+1)
                if not images:
                    raise ValueError(f"无法将页面 {page_num+1} 转换为图像")
                images[0].save(temp_image_path, 'PNG')
            except ImportError:
                self.logger.error("未安装pdf2image库，无法使用OCR功能")
                raise ImportError("请安装pdf2image库以使用OCR功能: pip install pdf2image")
            except Exception as e:
                self.logger.error(f"PDF转图像失败: {str(e)}")
                raise ValueError(f"无法将PDF转换为图像，可能需要安装poppler依赖: {str(e)}")
                
            # 使用OCR识别文字
            try:
                img = Image.open(temp_image_path)
                
                # 检查Tesseract是否可用
                try:
                    pytesseract.get_tesseract_version()
                except Exception:
                    self.logger.error("Tesseract OCR引擎未安装或不可用")
                    raise ImportError("请安装Tesseract OCR引擎: https://github.com/UB-Mannheim/tesseract/wiki")
                
                # 识别文字
                text = pytesseract.image_to_string(img, lang='chi_sim+eng')
                
                # 将识别的文字添加到Word文档
                if i > 0:
                    doc.add_page_break()
                doc.add_paragraph(text)
            except Exception as e:
                self.logger.error(f"OCR识别失败: {str(e)}")
                raise
            finally:
                # 删除临时文件
                if os.path.exists(temp_image_path):
                    os.remove(temp_image_path)
                    
        # 保存Word文档
        doc.save(output_path)
        
    def _get_convert_options(self, quality: str) -> Dict[str, Any]:
        """根据质量设置，返回转换选项"""
        # 默认选项
        options = {
            "debug": False
        }
        
        # 根据质量级别设置选项
        if quality == "low":
            options.update({
                "multi_processing": False,
                "first_paragraph": False,
                "connected_border": False
            })
        elif quality == "medium":
            options.update({
                "multi_processing": True,
                "first_paragraph": True,
                "connected_border": True
            })
        elif quality == "high":
            options.update({
                "multi_processing": True,
                "first_paragraph": True,
                "connected_border": True,
                "line_overlap_threshold": 0.9,
                "line_break_width_threshold": 1.0,
                "line_break_free_space_ratio": 0.1,
                "line_break_mode": True
            })
            
        return options 