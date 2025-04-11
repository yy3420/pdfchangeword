"""
测试PDF转换器功能
"""

import os
import sys
import unittest
from pathlib import Path

# 添加项目根目录到Python路径
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.converter import PDFConverter


class TestPDFConverter(unittest.TestCase):
    """PDF转换器功能测试类"""
    
    def setUp(self):
        """测试准备工作"""
        self.converter = PDFConverter()
        
        # 测试文件路径
        self.test_dir = Path(__file__).parent
        self.output_dir = self.test_dir / "output"
        
        # 确保输出目录存在
        os.makedirs(self.output_dir, exist_ok=True)
        
    def test_init(self):
        """测试初始化"""
        self.assertIsNotNone(self.converter)
        self.assertIsNotNone(self.converter.logger)
        
    def test_get_convert_options(self):
        """测试获取转换选项"""
        low_options = self.converter._get_convert_options("low")
        self.assertFalse(low_options["multi_processing"])
        
        medium_options = self.converter._get_convert_options("medium")
        self.assertTrue(medium_options["multi_processing"])
        
        high_options = self.converter._get_convert_options("high")
        self.assertTrue(high_options["multi_processing"])
        self.assertTrue("line_break_mode" in high_options)


if __name__ == "__main__":
    unittest.main() 