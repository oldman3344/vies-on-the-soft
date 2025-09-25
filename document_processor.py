#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档处理模块
用于处理Excel数据提取和Word文档表格填充
"""

import pandas as pd
import openpyxl
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
import os
import logging
from typing import List, Dict, Any, Optional, Tuple

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class DocumentProcessor:
    """文档处理器类"""
    
    def __init__(self):
        self.excel_data = None
        self.word_doc = None
    
    def set_table_borders(self, table):
        """
        为表格设置实线边框
        """
        try:
            from docx.oxml import parse_xml
            
            # 获取表格的XML元素
            tbl = table._tbl
            
            # 创建表格边框属性
            tblPr = tbl.tblPr
            if tblPr is None:
                tblPr = parse_xml('<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                tbl.insert(0, tblPr)
            
            # 创建表格边框XML
            borders_xml = '''
            <w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            </w:tblBorders>
            '''
            
            # 解析边框XML
            tblBorders = parse_xml(borders_xml)
            
            # 移除现有的边框设置（如果有）
            existing_borders = tblPr.find(qn('w:tblBorders'))
            if existing_borders is not None:
                tblPr.remove(existing_borders)
            
            # 添加新的边框设置
            tblPr.append(tblBorders)
            
        except Exception as e:
            logger.warning(f"设置表格边框时出错: {e}")
            # 如果设置边框失败，继续执行，不影响主要功能
        
    def extract_excel_data(self, excel_path: str, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        从Excel文件中提取数据
        
        Args:
            excel_path: Excel文件路径
            sheet_name: 工作表名称，如果为None则使用第一个工作表
            
        Returns:
            提取的数据列表
        """
        try:
            # 读取Excel文件
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(excel_path)
            
            logger.info(f"成功读取Excel文件: {excel_path}")
            logger.info(f"数据形状: {df.shape}")
            logger.info(f"列名: {list(df.columns)}")
            
            # 清理数据，移除空行
            df = df.dropna(how='all')
            
            # 转换为字典列表
            data_list = []
            for idx, row in df.iterrows():
                # 过滤包含VAT信息的行
                row_dict = row.to_dict()
                row_str = str(row.to_list()).lower()
                
                if 'vat' in row_str:
                    # 清理数据，将NaN替换为空字符串
                    cleaned_row = {}
                    for key, value in row_dict.items():
                        if pd.isna(value):
                            cleaned_row[key] = ""
                        else:
                            cleaned_row[key] = str(value)
                    
                    data_list.append(cleaned_row)
            
            self.excel_data = data_list
            logger.info(f"提取到 {len(data_list)} 行VAT相关数据")
            
            return data_list
            
        except Exception as e:
            logger.error(f"提取Excel数据时出错: {e}")
            raise
    
    def get_excel_columns(self, excel_path: str, sheet_name: Optional[str] = None) -> List[str]:
        """
        获取Excel文件的列名
        
        Args:
            excel_path: Excel文件路径
            sheet_name: 工作表名称
            
        Returns:
            列名列表
        """
        try:
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(excel_path)
            
            return list(df.columns)
            
        except Exception as e:
            logger.error(f"获取Excel列名时出错: {e}")
            raise
    
    def analyze_word_document(self, word_path: str) -> Dict[str, Any]:
        """
        分析Word文档结构
        
        Args:
            word_path: Word文档路径
            
        Returns:
            文档分析结果
        """
        try:
            doc = Document(word_path)
            
            analysis = {
                'paragraphs_count': len(doc.paragraphs),
                'tables_count': len(doc.tables),
                'tables_info': []
            }
            
            # 分析表格
            for i, table in enumerate(doc.tables):
                table_info = {
                    'table_index': i,
                    'rows': len(table.rows),
                    'columns': len(table.columns),
                    'content': []
                }
                
                # 获取表格内容
                for row_idx, row in enumerate(table.rows):
                    row_content = []
                    for cell in row.cells:
                        row_content.append(cell.text.strip())
                    table_info['content'].append(row_content)
                
                analysis['tables_info'].append(table_info)
            
            logger.info(f"Word文档分析完成: {analysis}")
            return analysis
            
        except Exception as e:
            logger.error(f"分析Word文档时出错: {e}")
            raise
    
    def fill_word_table(self, word_path: str, excel_data: List[Dict[str, Any]], 
                       output_path: str, table_index: int = 0,
                       column_mapping: Optional[Dict[str, str]] = None) -> str:
        """
        将Excel数据填充到Word表格中
        
        Args:
            word_path: Word模板文件路径
            excel_data: Excel数据
            output_path: 输出文件路径
            table_index: 要填充的表格索引
            column_mapping: 列映射关系 {excel_column: word_table_column}
            
        Returns:
            输出文件路径
        """
        try:
            # 打开Word文档
            doc = Document(word_path)
            
            if table_index >= len(doc.tables):
                raise ValueError(f"表格索引 {table_index} 超出范围，文档只有 {len(doc.tables)} 个表格")
            
            table = doc.tables[table_index]
            
            # 如果没有提供列映射，使用默认映射
            if column_mapping is None:
                # 根据Excel数据的列名创建默认映射
                if excel_data:
                    excel_columns = list(excel_data[0].keys())
                    column_mapping = {}
                    # 获取Word表格表头
                    word_headers = [cell.text.strip() for cell in table.rows[0].cells]
                    for i, excel_col in enumerate(excel_columns):
                        if i < len(word_headers):
                            # 将Excel列名映射到Word表头
                            column_mapping[word_headers[i]] = excel_col
            
            logger.info(f"使用列映射: {column_mapping}")
            
            # 确保表格有足够的行
            current_rows = len(table.rows)
            needed_rows = len(excel_data) + 1  # +1 for header
            
            # 添加行如果需要
            while len(table.rows) < needed_rows:
                table.add_row()
            
            # 删除所有非表头行（避免合并单元格问题）
            rows_to_remove = []
            for row_idx in range(len(table.rows) - 1, 0, -1):  # 从后往前删除
                rows_to_remove.append(row_idx)
            
            for row_idx in rows_to_remove:
                table._tbl.remove(table.rows[row_idx]._tr)
            
            # 填充数据
            for row_idx, data_row in enumerate(excel_data):
                # 从第二行开始填充（第一行通常是表头）
                word_row_idx = row_idx + 1
                
                if word_row_idx >= len(table.rows):
                    table.add_row()
                
                word_row = table.rows[word_row_idx]
                
                # 填充每一列
                if column_mapping:
                    # 获取Word表格表头
                    word_headers = [cell.text.strip() for cell in table.rows[0].cells]
                    
                    # 遍历Word表格的每一列
                    for word_col_idx, word_header in enumerate(word_headers):
                        if word_header in column_mapping:
                            # 获取对应的Excel列名
                            excel_col_name = column_mapping[word_header]
                            if excel_col_name in data_row:
                                cell = word_row.cells[word_col_idx]
                                cell.text = str(data_row[excel_col_name])
                                
                                # 设置单元格对齐方式
                                for paragraph in cell.paragraphs:
                                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 设置表格边框为实线
            self.set_table_borders(table)
            
            # 保存文档
            doc.save(output_path)
            logger.info(f"Word文档已保存到: {output_path}")
            
            return output_path
            
        except Exception as e:
            logger.error(f"填充Word表格时出错: {e}")
            raise
    
    def process_documents(self, excel_path: str, word_template_path: str, 
                         output_path: str, table_index: int = 0,
                         column_mapping: Optional[Dict[str, str]] = None,
                         max_rows: Optional[int] = None) -> Dict[str, Any]:
        """
        完整的文档处理流程
        
        Args:
            excel_path: Excel文件路径
            word_template_path: Word模板路径
            output_path: 输出文件路径
            table_index: 表格索引
            column_mapping: 列映射
            max_rows: 最大处理行数
            
        Returns:
            包含处理结果的字典
        """
        try:
            # 1. 提取Excel数据
            excel_data = self.extract_excel_data(excel_path)
            
            if not excel_data:
                return {
                    'success': False,
                    'error': '未找到有效的Excel数据',
                    'output_path': None,
                    'rows_filled': 0
                }
            
            # 限制行数
            if max_rows and max_rows > 0:
                excel_data = excel_data[:max_rows]
            
            # 2. 分析Word文档
            word_analysis = self.analyze_word_document(word_template_path)
            
            # 3. 生成输出路径（如果未提供）
            if not output_path:
                base_name = os.path.splitext(os.path.basename(word_template_path))[0]
                output_dir = os.path.dirname(word_template_path)
                output_path = os.path.join(output_dir, f"{base_name}_已填充.docx")
            
            # 4. 填充Word表格
            result_path = self.fill_word_table(
                word_template_path, 
                excel_data, 
                output_path, 
                table_index, 
                column_mapping
            )
            
            logger.info(f"文档处理完成，输出文件: {result_path}")
            return {
                'success': True,
                'error': None,
                'output_path': result_path,
                'rows_filled': len(excel_data)
            }
            
        except Exception as e:
            logger.error(f"文档处理过程中出错: {e}")
            return {
                'success': False,
                'error': str(e),
                'output_path': None,
                'rows_filled': 0
            }


def create_default_column_mapping() -> Dict[str, str]:
    """
    创建默认的列映射关系
    基于Word表格字段与Excel数据列的对应关系
    Word表格字段 -> Excel列名
    """
    return {
        '客户': '公司名称',           # 对应Excel的"公司名称"列
        '国家': '国家',             # 对应Excel的"国家"列  
        '申报方式': '申报方式',       # 对应Excel的"申报方式"列
        '申报时段': '申报时段',       # 对应Excel的"申报时段"列
        '下载数据格式': '下载数据格式',   # 对应Excel的"下载数据格式"列
        '备注(如已续费，请忽略）': '备注'  # 对应Excel的"备注"列
    }


if __name__ == "__main__":
    # 测试代码
    processor = DocumentProcessor()
    
    # 测试路径
    excel_path = "/Volumes/oldman_space/work_space/vies-on-the-soft/25.08月申报明细汇总表.xlsx"
    word_path = "/Volumes/oldman_space/work_space/vies-on-the-soft/VAT申报明细表模板.docx"
    output_path = "/Volumes/oldman_space/work_space/vies-on-the-soft/VAT申报明细表_已填充.docx"
    
    try:
        # 测试Excel数据提取
        data = processor.extract_excel_data(excel_path)
        print(f"提取到 {len(data)} 行数据")
        
        if data:
            print("前3行数据:")
            for i, row in enumerate(data[:3]):
                print(f"第{i+1}行: {row}")
        
        # 测试Word文档分析
        if os.path.exists(word_path):
            analysis = processor.analyze_word_document(word_path)
            print(f"Word文档分析结果: {analysis}")
            
            # 测试完整处理流程
            column_mapping = create_default_column_mapping()
            result = processor.process_documents(
                excel_path, word_path, output_path, 
                table_index=0, column_mapping=column_mapping
            )
            print(f"处理完成，输出文件: {result}")
        else:
            print(f"Word模板文件不存在: {word_path}")
            
    except Exception as e:
        print(f"测试过程中出错: {e}")