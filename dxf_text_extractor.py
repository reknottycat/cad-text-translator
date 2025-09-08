#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DXF文本提取引擎 - 模块化架构
功能：从DXF文件中提取文本内容，采用模块化设计
作者：CAD翻译工具团队
版本：2.0
"""

import os
import sys
import logging
from pathlib import Path
from typing import List, Set, Dict, Optional, Tuple
from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum

try:
    import ezdxf
    from ezdxf.document import Drawing
    from ezdxf.entities import DXFEntity
except ImportError:
    print("错误：未安装ezdxf库，请运行: pip install ezdxf")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("错误：未安装pandas库，请运行: pip install pandas")
    sys.exit(1)

# 确保日志目录存在
os.makedirs('logs', exist_ok=True)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/text_extraction.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class ExtractionMethod(Enum):
    """文本提取方法枚举"""
    MODEL_SPACE = "model_space"
    PAPER_SPACE = "paper_space"
    BLOCK_DEFINITIONS = "block_definitions"
    DXF_TAGS = "dxf_tags"
    EXTENDED_DATA = "extended_data"


@dataclass
class ExtractionResult:
    """提取结果数据类"""
    texts: List[str]
    method: ExtractionMethod
    file_path: str
    success: bool
    error_message: Optional[str] = None


class TextExtractor(ABC):
    """文本提取器抽象基类"""
    
    @abstractmethod
    def extract(self, doc: Drawing, file_path: str) -> ExtractionResult:
        """提取文本的抽象方法"""
        pass


class ModelSpaceExtractor(TextExtractor):
    """模型空间文本提取器"""
    
    def extract(self, doc: Drawing, file_path: str) -> ExtractionResult:
        """从模型空间提取文本"""
        texts = []
        try:
            msp = doc.modelspace()
            # 提取TEXT实体
            for entity in msp.query('TEXT'):
                if hasattr(entity, 'dxf') and hasattr(entity.dxf, 'text'):
                    text = entity.dxf.text.strip()
                    if text:
                        texts.append(text)
            
            # 提取MTEXT实体
            for entity in msp.query('MTEXT'):
                if hasattr(entity, 'text'):
                    text = entity.text.strip()
                    if text:
                        texts.append(text)
            
            # 提取INSERT实体的属性
            for entity in msp.query('INSERT'):
                if hasattr(entity, 'attribs'):
                    for attrib in entity.attribs:
                        if hasattr(attrib, 'dxf') and hasattr(attrib.dxf, 'text'):
                            text = attrib.dxf.text.strip()
                            if text:
                                texts.append(text)
            
            return ExtractionResult(
                texts=texts,
                method=ExtractionMethod.MODEL_SPACE,
                file_path=file_path,
                success=True
            )
            
        except Exception as e:
            logger.error(f"模型空间文本提取失败: {e}")
            return ExtractionResult(
                texts=[],
                method=ExtractionMethod.MODEL_SPACE,
                file_path=file_path,
                success=False,
                error_message=str(e)
            )


class PaperSpaceExtractor(TextExtractor):
    """图纸空间文本提取器"""
    
    def extract(self, doc: Drawing, file_path: str) -> ExtractionResult:
        """从图纸空间提取文本"""
        texts = []
        try:
            # 遍历所有布局（图纸空间）
            for layout in doc.layouts:
                if layout.name != 'Model':  # 跳过模型空间
                    # 提取TEXT和MTEXT实体
                    for entity in layout.query('TEXT MTEXT'):
                        text = self._extract_text_from_entity(entity)
                        if text:
                            texts.append(text)
                    
                    # 提取INSERT实体的属性
                    for entity in layout.query('INSERT'):
                        if hasattr(entity, 'attribs'):
                            for attrib in entity.attribs:
                                text = self._extract_text_from_entity(attrib)
                                if text:
                                    texts.append(text)
            
            return ExtractionResult(
                texts=texts,
                method=ExtractionMethod.PAPER_SPACE,
                file_path=file_path,
                success=True
            )
            
        except Exception as e:
            logger.error(f"图纸空间文本提取失败: {e}")
            return ExtractionResult(
                texts=[],
                method=ExtractionMethod.PAPER_SPACE,
                file_path=file_path,
                success=False,
                error_message=str(e)
            )
    
    def _extract_text_from_entity(self, entity: DXFEntity) -> Optional[str]:
        """从实体中提取文本"""
        try:
            if hasattr(entity, 'dxf') and hasattr(entity.dxf, 'text'):
                return entity.dxf.text.strip()
            elif hasattr(entity, 'text'):
                return entity.text.strip()
        except:
            pass
        return None


class BlockDefinitionExtractor(TextExtractor):
    """块定义文本提取器"""
    
    def extract(self, doc: Drawing, file_path: str) -> ExtractionResult:
        """从块定义中提取文本"""
        texts = []
        try:
            # 遍历所有块定义
            for block_name, block in doc.blocks.items():
                if not block_name.startswith('*'):  # 跳过匿名块
                    # 提取TEXT和MTEXT实体
                    for entity in block.query('TEXT MTEXT'):
                        if hasattr(entity, 'dxf') and hasattr(entity.dxf, 'text'):
                            text = entity.dxf.text.strip()
                            if text:
                                texts.append(text)
                        elif hasattr(entity, 'text'):
                            text = entity.text.strip()
                            if text:
                                texts.append(text)
                    
                    # 提取属性定义
                    for entity in block.query('ATTDEF'):
                        if hasattr(entity, 'dxf'):
                            # 提取默认值
                            if hasattr(entity.dxf, 'text'):
                                text = entity.dxf.text.strip()
                                if text:
                                    texts.append(text)
                            # 提取标签
                            if hasattr(entity.dxf, 'tag'):
                                tag = entity.dxf.tag.strip()
                                if tag:
                                    texts.append(tag)
            
            return ExtractionResult(
                texts=texts,
                method=ExtractionMethod.BLOCK_DEFINITIONS,
                file_path=file_path,
                success=True
            )
            
        except Exception as e:
            logger.error(f"块定义文本提取失败: {e}")
            return ExtractionResult(
                texts=[],
                method=ExtractionMethod.BLOCK_DEFINITIONS,
                file_path=file_path,
                success=False,
                error_message=str(e)
            )


class DXFTagExtractor(TextExtractor):
    """DXF标签文本提取器"""
    
    def extract(self, doc: Drawing, file_path: str) -> ExtractionResult:
        """从DXF标签中提取文本"""
        texts = set()  # 使用集合避免重复
        try:
            # 读取原始DXF文件内容
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
            
            # 解析DXF标签
            for i in range(0, len(lines) - 1, 2):
                try:
                    group_code = int(lines[i].strip())
                    value = lines[i + 1].strip()
                    
                    # 只提取文本相关的组码
                    if group_code in [1, 3, 7, 8] and value:
                        # 过滤掉技术值和无意义的文本
                        if self._is_meaningful_text(value):
                            texts.add(value)
                            
                except (ValueError, IndexError):
                    continue
            
            return ExtractionResult(
                texts=list(texts),
                method=ExtractionMethod.DXF_TAGS,
                file_path=file_path,
                success=True
            )
            
        except Exception as e:
            logger.error(f"DXF标签文本提取失败: {e}")
            return ExtractionResult(
                texts=[],
                method=ExtractionMethod.DXF_TAGS,
                file_path=file_path,
                success=False,
                error_message=str(e)
            )
    
    def _is_technical_value(self, value: str) -> bool:
        """判断是否为技术值"""
        # 检查是否为数字
        try:
            float(value)
            return True
        except ValueError:
            pass
        
        # 检查是否为坐标格式
        if ',' in value and all(part.replace('.', '').replace('-', '').isdigit() for part in value.split(',')):
            return True
        
        return False
    
    def _is_handle_value(self, value: str) -> bool:
        """判断是否为句柄值"""
        # DXF句柄通常是十六进制字符串
        if len(value) <= 8 and all(c in '0123456789ABCDEFabcdef' for c in value):
            return True
        return False
    
    def _is_layer_name(self, value: str) -> bool:
        """判断是否为图层名"""
        # 常见的图层名模式
        layer_patterns = [
            '0',  # 默认图层
            'DEFPOINTS',
            'TEXT',
            'DIM',
            'HATCH'
        ]
        
        if value in layer_patterns:
            return True
        
        # 检查是否为图层名格式（通常包含特定前缀或后缀）
        if value.startswith(('LAYER_', 'L_', 'LAY_')):
            return True
        
        return False
    
    def _is_short_hex(self, value: str) -> bool:
        """判断是否为短十六进制值"""
        if len(value) <= 4 and all(c in '0123456789ABCDEFabcdef' for c in value):
            return True
        return False
    
    def _is_meaningful_text(self, value: str) -> bool:
        """判断是否为有意义的文本"""
        # 过滤条件
        if not value or len(value.strip()) == 0:
            return False
        
        # 过滤技术值
        if self._is_technical_value(value):
            return False
        
        # 过滤句柄值
        if self._is_handle_value(value):
            return False
        
        # 过滤图层名
        if self._is_layer_name(value):
            return False
        
        # 过滤短十六进制值
        if self._is_short_hex(value):
            return False
        
        # 过滤CAD实体类型
        if self._is_cad_entity_type(value):
            return False
        
        # 过滤过短的文本（可能是代码）
        if len(value) < 2:
            return False
        
        # 通过所有过滤条件，认为是有意义的文本
        return True
    
    def _is_cad_entity_type(self, value: str) -> bool:
        """判断是否为CAD实体类型"""
        entity_types = {
            'SECTION', 'ENDSEC', 'HEADER', 'CLASSES', 'TABLES', 'BLOCKS', 'ENTITIES',
            'OBJECTS', 'EOF', 'LINE', 'CIRCLE', 'ARC', 'TEXT', 'MTEXT', 'INSERT',
            'POLYLINE', 'LWPOLYLINE', 'POINT', 'ELLIPSE', 'SPLINE', 'HATCH',
            'DIMENSION', 'LEADER', 'VIEWPORT', 'ACDBTEXT', 'ACDBMTEXT'
        }
        return value.upper() in entity_types


class TextFilter:
    """文本过滤器"""
    
    def __init__(self, min_length: int = 1, max_length: int = 1000):
        self.min_length = min_length
        self.max_length = max_length
        self.exclude_patterns = {
            r'^\d+$',  # 纯数字
            r'^[A-Fa-f0-9]+$',  # 十六进制
            r'^[\s\-_\.]+$',  # 只包含空格、横线、下划线、点
        }
    
    def filter_texts(self, texts: List[str]) -> List[str]:
        """过滤文本列表"""
        filtered = []
        for text in texts:
            if self._is_valid_text(text):
                cleaned = self._clean_text(text)
                if cleaned and cleaned not in filtered:
                    filtered.append(cleaned)
        return filtered
    
    def _is_valid_text(self, text: str) -> bool:
        """判断文本是否有效"""
        if not text or not isinstance(text, str):
            return False
        
        text = text.strip()
        if len(text) < self.min_length or len(text) > self.max_length:
            return False
        
        # 检查排除模式
        import re
        for pattern in self.exclude_patterns:
            if re.match(pattern, text):
                return False
        
        return True
    
    def _clean_text(self, text: str) -> str:
        """清理文本"""
        # 移除首尾空白
        text = text.strip()
        
        # 移除多余的空白字符
        import re
        text = re.sub(r'\s+', ' ', text)
        
        return text


class ExcelExporter:
    """Excel导出器"""
    
    def export_to_excel(self, texts: List[str], output_path: str) -> bool:
        """导出文本到Excel文件"""
        try:
            # 创建DataFrame
            df = pd.DataFrame({
                '序号': range(1, len(texts) + 1),
                '原文': texts,
                '译文': [''] * len(texts)  # 空的译文列
            })
            
            # 导出到Excel
            df.to_excel(output_path, index=False, engine='openpyxl')
            logger.info(f"成功导出 {len(texts)} 条文本到 {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"导出Excel失败: {e}")
            return False


class TextExtractionEngine:
    """文本提取引擎"""
    
    def __init__(self):
        self.extractors = {
            ExtractionMethod.MODEL_SPACE: ModelSpaceExtractor(),
            ExtractionMethod.PAPER_SPACE: PaperSpaceExtractor(),
            ExtractionMethod.BLOCK_DEFINITIONS: BlockDefinitionExtractor(),
            ExtractionMethod.DXF_TAGS: DXFTagExtractor()
        }
        self.text_filter = TextFilter()
        self.excel_exporter = ExcelExporter()
    
    def extract_from_file(self, file_path: str) -> List[str]:
        """从单个文件提取文本"""
        all_texts = set()  # 使用集合避免重复
        
        try:
            # 尝试打开DXF文件
            doc = ezdxf.readfile(file_path)
            logger.info(f"成功打开文件: {file_path}")
            
            # 使用所有提取器
            for method, extractor in self.extractors.items():
                try:
                    result = extractor.extract(doc, file_path)
                    if result.success:
                        all_texts.update(result.texts)
                        logger.debug(f"{method.value} 提取到 {len(result.texts)} 条文本")
                    else:
                        logger.warning(f"{method.value} 提取失败: {result.error_message}")
                except Exception as e:
                    logger.error(f"{method.value} 提取器异常: {e}")
            
        except ezdxf.DXFStructureError:
            logger.warning(f"DXF结构错误，尝试修复: {file_path}")
            # 尝试修复并重新提取
            repaired_texts = self._try_repair_and_extract(file_path)
            all_texts.update(repaired_texts)
            
        except Exception as e:
            logger.error(f"文件处理失败: {file_path}, 错误: {e}")
        
        # 过滤和清理文本
        filtered_texts = self.text_filter.filter_texts(list(all_texts))
        logger.info(f"从 {file_path} 提取到 {len(filtered_texts)} 条有效文本")
        
        return filtered_texts
    
    def _try_repair_and_extract(self, file_path: str) -> List[str]:
        """尝试修复文件并提取文本"""
        try:
            # 使用DXF标签提取器作为备用方案
            extractor = self.extractors[ExtractionMethod.DXF_TAGS]
            result = extractor.extract(None, file_path)
            
            if result.success:
                logger.info(f"通过标签提取器成功提取文本: {len(result.texts)} 条")
                return result.texts
            else:
                logger.error(f"标签提取器也失败: {result.error_message}")
                return []
                
        except Exception as e:
            logger.error(f"修复提取失败: {e}")
            return []
    
    def extract_from_directory(self, directory_path: str) -> List[str]:
        """从目录中的所有DXF文件提取文本"""
        directory = Path(directory_path)
        if not directory.exists() or not directory.is_dir():
            logger.error(f"目录不存在: {directory_path}")
            return []
        
        # 查找所有DXF文件（包括子目录）
        dxf_files = list(directory.glob('**/*.dxf'))
        logger.info(f"在 {directory_path} 中找到 {len(dxf_files)} 个DXF文件")
        
        all_texts = set()
        
        for i, dxf_file in enumerate(dxf_files, 1):
            logger.info(f"处理文件 {i}/{len(dxf_files)}: {dxf_file.name}")
            texts = self.extract_from_file(str(dxf_file))
            all_texts.update(texts)
        
        final_texts = list(all_texts)
        logger.info(f"总共提取到 {len(final_texts)} 条唯一文本")
        
        return final_texts
    
    def process_and_export(self, input_path: str, output_path: str) -> bool:
        """处理输入并导出结果"""
        input_path_obj = Path(input_path)
        
        if input_path_obj.is_file():
            texts = self.extract_from_file(input_path)
        elif input_path_obj.is_dir():
            texts = self.extract_from_directory(input_path)
        else:
            logger.error(f"输入路径无效: {input_path}")
            return False
        
        if texts:
            return self.excel_exporter.export_to_excel(texts, output_path)
        else:
            logger.warning("未提取到任何文本")
            return False


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='DXF文本提取工具')
    parser.add_argument('input', help='输入DXF文件或目录路径')
    parser.add_argument('--output', '-o', default='extracted_texts.xlsx', help='输出Excel文件路径')
    parser.add_argument('--verbose', '-v', action='store_true', help='显示详细日志')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # 创建提取引擎
    engine = TextExtractionEngine()
    
    # 处理并导出
    success = engine.process_and_export(args.input, args.output)
    
    if success:
        print(f"文本提取完成，结果已保存到: {args.output}")
        return 0
    else:
        print("文本提取失败")
        return 1


if __name__ == '__main__':
    main()