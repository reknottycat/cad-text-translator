#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DXF文本提取脚本

功能：
1. 从DXF文件中提取所有文本内容
2. 支持多种文本实体类型（TEXT、MTEXT、DIMENSION等）
3. 导出到Excel文件供翻译使用
4. 提供多种提取方法和过滤选项

作者: AI Assistant
日期: 2024
"""

import os
import sys
import argparse
from pathlib import Path
import pandas as pd
from typing import List, Dict, Set, Optional, Tuple
import re
from collections import defaultdict

try:
    import ezdxf
except ImportError:
    print("错误: 未安装ezdxf库")
    print("请运行: pip install ezdxf")
    sys.exit(1)

# 导入日志配置
try:
    from logger_config import get_logger
    logger = get_logger("extract_texts")
except ImportError:
    import logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger("extract_texts")

class DXFTextExtractor:
    """DXF文本提取器"""
    
    def __init__(self, dxf_file: str):
        """初始化提取器
        
        Args:
            dxf_file: DXF文件路径
        """
        self.dxf_file = Path(dxf_file)
        self.doc = None
        self.texts = []
        self.stats = defaultdict(int)
        
        # 文本过滤配置
        self.min_length = 1
        self.max_length = 1000
        self.exclude_patterns = [
            r'^\s*$',  # 空白字符
            r'^[\d\.\-\+\s]*$',  # 纯数字
            r'^[A-Za-z]$',  # 单个字母
        ]
        self.exclude_layers = set()
        
    def load_dxf(self) -> bool:
        """加载DXF文件
        
        Returns:
            bool: 是否加载成功
        """
        try:
            logger.info(f"正在加载DXF文件: {self.dxf_file}")
            self.doc = ezdxf.readfile(str(self.dxf_file))
            logger.info(f"DXF文件加载成功，版本: {self.doc.dxfversion}")
            return True
        except Exception as e:
            logger.error(f"加载DXF文件失败: {e}")
            return False
    
    def should_exclude_text(self, text: str, layer: str = "") -> bool:
        """判断是否应该排除文本
        
        Args:
            text: 文本内容
            layer: 图层名称
            
        Returns:
            bool: 是否应该排除
        """
        # 检查长度
        if len(text) < self.min_length or len(text) > self.max_length:
            return True
            
        # 检查图层
        if layer in self.exclude_layers:
            return True
            
        # 检查模式
        for pattern in self.exclude_patterns:
            if re.match(pattern, text):
                return True
                
        return False
    
    def extract_text_entities(self) -> List[Dict]:
        """提取TEXT实体
        
        Returns:
            List[Dict]: 文本信息列表
        """
        texts = []
        
        for entity in self.doc.modelspace().query('TEXT'):
            try:
                text_content = entity.dxf.text.strip()
                layer = entity.dxf.layer
                
                if self.should_exclude_text(text_content, layer):
                    continue
                    
                text_info = {
                    'type': 'TEXT',
                    'handle': entity.dxf.handle,
                    'text': text_content,
                    'layer': layer,
                    'position': f"({entity.dxf.insert.x:.2f}, {entity.dxf.insert.y:.2f})",
                    'height': entity.dxf.height,
                    'rotation': entity.dxf.rotation,
                    'style': getattr(entity.dxf, 'style', ''),
                }
                texts.append(text_info)
                self.stats['TEXT'] += 1
                
            except Exception as e:
                logger.warning(f"处理TEXT实体时出错: {e}")
                
        return texts
    
    def extract_mtext_entities(self) -> List[Dict]:
        """提取MTEXT实体
        
        Returns:
            List[Dict]: 文本信息列表
        """
        texts = []
        
        for entity in self.doc.modelspace().query('MTEXT'):
            try:
                # MTEXT可能包含格式代码，需要清理
                raw_text = entity.text
                # 移除常见的格式代码
                clean_text = re.sub(r'\\[A-Za-z][^;]*;', '', raw_text)
                clean_text = re.sub(r'\{[^}]*\}', '', clean_text)
                clean_text = clean_text.strip()
                
                layer = entity.dxf.layer
                
                if self.should_exclude_text(clean_text, layer):
                    continue
                    
                text_info = {
                    'type': 'MTEXT',
                    'handle': entity.dxf.handle,
                    'text': clean_text,
                    'layer': layer,
                    'position': f"({entity.dxf.insert.x:.2f}, {entity.dxf.insert.y:.2f})",
                    'height': entity.dxf.char_height,
                    'rotation': entity.dxf.rotation,
                    'style': getattr(entity.dxf, 'style', ''),
                }
                texts.append(text_info)
                self.stats['MTEXT'] += 1
                
            except Exception as e:
                logger.warning(f"处理MTEXT实体时出错: {e}")
                
        return texts
    
    def extract_dimension_texts(self) -> List[Dict]:
        """提取标注文本
        
        Returns:
            List[Dict]: 文本信息列表
        """
        texts = []
        dimension_types = ['DIMENSION', 'ALIGNED_DIMENSION', 'LINEAR_DIMENSION', 
                          'RADIAL_DIMENSION', 'ANGULAR_DIMENSION']
        
        for dim_type in dimension_types:
            for entity in self.doc.modelspace().query(dim_type):
                try:
                    # 获取标注文本
                    dim_text = getattr(entity.dxf, 'text', '')
                    if not dim_text:
                        # 如果没有自定义文本，可能使用默认测量值
                        dim_text = '<>'  # AutoCAD中表示使用测量值
                    
                    layer = entity.dxf.layer
                    
                    if dim_text and not self.should_exclude_text(dim_text, layer):
                        text_info = {
                            'type': f'DIM_{dim_type}',
                            'handle': entity.dxf.handle,
                            'text': dim_text,
                            'layer': layer,
                            'position': 'N/A',  # 标注位置较复杂
                            'height': getattr(entity.dxf, 'dimtxt', 0),
                            'rotation': 0,
                            'style': getattr(entity.dxf, 'dimstyle', ''),
                        }
                        texts.append(text_info)
                        self.stats[f'DIM_{dim_type}'] += 1
                        
                except Exception as e:
                    logger.warning(f"处理{dim_type}实体时出错: {e}")
                    
        return texts
    
    def extract_attribute_texts(self) -> List[Dict]:
        """提取属性文本（块属性）
        
        Returns:
            List[Dict]: 文本信息列表
        """
        texts = []
        
        for entity in self.doc.modelspace().query('INSERT'):
            try:
                if entity.has_attrib:
                    for attrib in entity.attribs:
                        text_content = attrib.dxf.text.strip()
                        layer = attrib.dxf.layer
                        
                        if self.should_exclude_text(text_content, layer):
                            continue
                            
                        text_info = {
                            'type': 'ATTRIB',
                            'handle': attrib.dxf.handle,
                            'text': text_content,
                            'layer': layer,
                            'position': f"({attrib.dxf.insert.x:.2f}, {attrib.dxf.insert.y:.2f})",
                            'height': attrib.dxf.height,
                            'rotation': attrib.dxf.rotation,
                            'style': getattr(attrib.dxf, 'style', ''),
                            'tag': attrib.dxf.tag,  # 属性标签
                        }
                        texts.append(text_info)
                        self.stats['ATTRIB'] += 1
                        
            except Exception as e:
                logger.warning(f"处理INSERT实体属性时出错: {e}")
                
        return texts
    
    def extract_all_texts(self) -> List[Dict]:
        """提取所有文本
        
        Returns:
            List[Dict]: 所有文本信息列表
        """
        logger.info("开始提取文本")
        all_texts = []
        
        # 提取各种类型的文本
        all_texts.extend(self.extract_text_entities())
        all_texts.extend(self.extract_mtext_entities())
        all_texts.extend(self.extract_dimension_texts())
        all_texts.extend(self.extract_attribute_texts())
        
        # 去重（基于handle）
        seen_handles = set()
        unique_texts = []
        for text in all_texts:
            if text['handle'] not in seen_handles:
                unique_texts.append(text)
                seen_handles.add(text['handle'])
            else:
                logger.debug(f"跳过重复文本: {text['handle']}")
        
        self.texts = unique_texts
        logger.info(f"提取完成，共找到 {len(unique_texts)} 个文本")
        
        return unique_texts
    
    def filter_chinese_texts(self, texts: List[Dict]) -> List[Dict]:
        """过滤出包含中文的文本
        
        Args:
            texts: 文本列表
            
        Returns:
            List[Dict]: 包含中文的文本列表
        """
        chinese_texts = []
        chinese_pattern = re.compile(r'[\u4e00-\u9fff]+')
        
        for text in texts:
            if chinese_pattern.search(text['text']):
                chinese_texts.append(text)
                
        logger.info(f"过滤出包含中文的文本: {len(chinese_texts)} 个")
        return chinese_texts
    
    def export_to_excel(self, texts: List[Dict], output_file: str) -> bool:
        """导出到Excel文件
        
        Args:
            texts: 文本列表
            output_file: 输出文件路径
            
        Returns:
            bool: 是否导出成功
        """
        try:
            if not texts:
                logger.warning("没有文本可导出")
                return False
                
            # 准备数据
            data = []
            for i, text in enumerate(texts, 1):
                row = {
                    '序号': i,
                    '类型': text['type'],
                    '句柄': text['handle'],
                    '原文': text['text'],
                    '译文': '',  # 空白列供填写翻译
                    '图层': text['layer'],
                    '位置': text['position'],
                    '高度': text.get('height', 0),
                    '旋转角度': text.get('rotation', 0),
                    '样式': text.get('style', ''),
                }
                
                # 添加属性特有字段
                if 'tag' in text:
                    row['属性标签'] = text['tag']
                    
                data.append(row)
            
            # 创建DataFrame
            df = pd.DataFrame(data)
            
            # 导出到Excel
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='文本提取', index=False)
                
                # 获取工作表以设置列宽
                worksheet = writer.sheets['文本提取']
                
                # 设置列宽
                column_widths = {
                    'A': 8,   # 序号
                    'B': 12,  # 类型
                    'C': 15,  # 句柄
                    'D': 30,  # 原文
                    'E': 30,  # 译文
                    'F': 15,  # 图层
                    'G': 20,  # 位置
                    'H': 10,  # 高度
                    'I': 12,  # 旋转角度
                    'J': 15,  # 样式
                }
                
                for col, width in column_widths.items():
                    worksheet.column_dimensions[col].width = width
            
            logger.info(f"成功导出到Excel文件: {output_file}")
            return True
            
        except Exception as e:
            logger.error(f"导出Excel文件失败: {e}")
            return False
    
    def print_statistics(self):
        """打印统计信息"""
        print("\n" + "="*50)
        print("           提取统计")
        print("="*50)
        
        total = sum(self.stats.values())
        print(f"总计文本数量: {total}")
        
        if self.stats:
            print("\n按类型分布:")
            for text_type, count in sorted(self.stats.items()):
                percentage = (count / total * 100) if total > 0 else 0
                print(f"  {text_type:15}: {count:4d} ({percentage:5.1f}%)")
        
        print("="*50)

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='DXF文本提取工具')
    parser.add_argument('input', nargs='?', help='输入DXF文件路径')
    parser.add_argument('-o', '--output', help='输出Excel文件路径')
    parser.add_argument('--chinese-only', action='store_true', help='只提取包含中文的文本')
    parser.add_argument('--min-length', type=int, default=1, help='最小文本长度')
    parser.add_argument('--max-length', type=int, default=1000, help='最大文本长度')
    parser.add_argument('--exclude-layers', nargs='*', help='要排除的图层名称')
    
    args = parser.parse_args()
    
    # 如果没有指定输入文件，自动查找
    if not args.input:
        dxf_files = list(Path('.').glob('*.dxf')) + list(Path('.').glob('*.DXF'))
        if not dxf_files:
            print("错误: 未找到DXF文件")
            print("请指定DXF文件路径或将DXF文件放在当前目录")
            return 1
        elif len(dxf_files) == 1:
            args.input = str(dxf_files[0])
            print(f"自动选择文件: {args.input}")
        else:
            print("找到多个DXF文件:")
            for i, f in enumerate(dxf_files, 1):
                print(f"  {i}. {f.name}")
            try:
                choice = int(input("请选择文件编号: ")) - 1
                if 0 <= choice < len(dxf_files):
                    args.input = str(dxf_files[choice])
                else:
                    print("无效选择")
                    return 1
            except (ValueError, KeyboardInterrupt):
                print("\n操作取消")
                return 1
    
    # 设置输出文件名
    if not args.output:
        input_path = Path(args.input)
        args.output = f"{input_path.stem}_texts.xlsx"
    
    print(f"\n输入文件: {args.input}")
    print(f"输出文件: {args.output}")
    
    # 创建提取器
    extractor = DXFTextExtractor(args.input)
    
    # 设置过滤参数
    extractor.min_length = args.min_length
    extractor.max_length = args.max_length
    if args.exclude_layers:
        extractor.exclude_layers = set(args.exclude_layers)
    
    # 加载DXF文件
    if not extractor.load_dxf():
        return 1
    
    # 提取文本
    texts = extractor.extract_all_texts()
    
    if not texts:
        print("未找到任何文本")
        return 1
    
    # 过滤中文文本（如果指定）
    if args.chinese_only:
        texts = extractor.filter_chinese_texts(texts)
        if not texts:
            print("未找到包含中文的文本")
            return 1
    
    # 导出到Excel
    if extractor.export_to_excel(texts, args.output):
        print(f"\n✓ 成功导出 {len(texts)} 个文本到: {args.output}")
    else:
        print("✗ 导出失败")
        return 1
    
    # 打印统计信息
    extractor.print_statistics()
    
    return 0

if __name__ == '__main__':
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n操作被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n程序发生错误: {e}")
        logger.error(f"程序发生错误: {e}")
        sys.exit(1)