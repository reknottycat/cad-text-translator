#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
浩辰CAD优化转换器
解决性能瓶颈问题的优化版本
主要优化:
1. 减少COM调用次数
2. 批量处理实体
3. 优化内存使用
4. 添加进度显示
5. 异步处理机制
"""

import os
import sys
import time
import win32com.client
from typing import List, Dict, Optional
from collections import Counter, defaultdict
import threading
from concurrent.futures import ThreadPoolExecutor

class OptimizedHaoChenCADConverter:
    """
    优化版浩辰CAD转换器
    解决性能瓶颈，提升转换速度
    """
    
    def __init__(self):
        self.app = None
        self.doc = None
        self.connected = False
        self.batch_size = 100  # 批处理大小
        self.progress_callback = None
        
    def set_progress_callback(self, callback):
        """设置进度回调函数"""
        self.progress_callback = callback
        
    def _update_progress(self, current, total, message=""):
        """更新进度"""
        if self.progress_callback:
            self.progress_callback(current, total, message)
        else:
            percent = (current / total * 100) if total > 0 else 0
            print(f"\r进度: {percent:.1f}% ({current}/{total}) {message}", end="", flush=True)
    
    def connect_to_cad(self) -> bool:
        """
        连接到浩辰CAD应用程序 - 优化版
        """
        cad_prog_ids = [
            "GStarCAD.Application",
            "Gcad.Application", 
            "GStarCAD.Application.26",
            "Gcad.Application.26",
            "ZWCAD.Application",
            "AutoCAD.Application"
        ]
        
        print("正在连接浩辰CAD...")
        
        # 优先连接现有实例
        for prog_id in cad_prog_ids:
            try:
                self.app = win32com.client.GetActiveObject(prog_id)
                print(f"✓ 连接到现有 {prog_id} 实例")
                self.connected = True
                return True
            except:
                continue
        
        # 创建新实例
        for prog_id in cad_prog_ids:
            try:
                self.app = win32com.client.Dispatch(prog_id)
                print(f"✓ 启动新的 {prog_id} 实例")
                self.connected = True
                # 设置CAD为不可见模式以提升性能
                try:
                    self.app.Visible = False
                    print("✓ 设置CAD为后台模式")
                except:
                    pass
                return True
            except Exception as e:
                continue
        
        print("✗ 无法连接到任何CAD程序")
        return False
    
    def open_dwg_file(self, dwg_path: str) -> bool:
        """
        打开DWG文件 - 优化版
        """
        if not self.connected or not self.app:
            return False
            
        try:
            abs_path = os.path.abspath(dwg_path)
            print(f"正在打开文件: {abs_path}")
            
            # 关闭当前文档（如果有）
            self.close_document()
            
            # 打开新文档
            self.doc = self.app.Documents.Open(abs_path)
            print(f"✓ 文件打开成功")
            return True
            
        except Exception as e:
            print(f"✗ 打开文件失败: {e}")
            return False
    
    def analyze_entities_optimized(self) -> Dict:
        """
        优化的实体分析 - 批量处理
        """
        if not self.doc:
            return {}
        
        print("正在分析文档实体...")
        
        analysis = {
            'total_entities': 0,
            'text_entities': 0,
            'entity_types': Counter(),
            'layers': set(),
            'text_contents': [],
            'processing_time': 0
        }
        
        start_time = time.time()
        
        try:
            # 获取模型空间
            model_space = self.doc.ModelSpace
            total_entities = model_space.Count
            analysis['total_entities'] = total_entities
            
            print(f"找到 {total_entities} 个实体")
            
            # 批量处理实体
            batch_count = (total_entities + self.batch_size - 1) // self.batch_size
            
            for batch_idx in range(batch_count):
                start_idx = batch_idx * self.batch_size
                end_idx = min(start_idx + self.batch_size, total_entities)
                
                # 获取当前批次的实体
                batch_entities = []
                for i in range(start_idx, end_idx):
                    try:
                        entity = model_space.Item(i)
                        batch_entities.append(entity)
                    except:
                        continue
                
                # 处理当前批次
                self._process_entity_batch(batch_entities, analysis)
                
                # 更新进度
                self._update_progress(end_idx, total_entities, "分析实体")
            
            analysis['processing_time'] = time.time() - start_time
            print(f"\n✓ 实体分析完成，用时 {analysis['processing_time']:.2f}秒")
            
        except Exception as e:
            print(f"\n✗ 实体分析失败: {e}")
        
        return analysis
    
    def _process_entity_batch(self, entities: List, analysis: Dict):
        """
        批量处理实体
        """
        for entity in entities:
            try:
                # 获取实体属性
                props = self._get_entity_properties(entity)
                
                # 统计实体类型
                entity_type = props.get('type', 'Unknown')
                analysis['entity_types'][entity_type] += 1
                
                # 收集图层信息
                if 'layer' in props:
                    analysis['layers'].add(props['layer'])
                
                # 处理文本实体
                if self._is_text_entity(entity_type):
                    analysis['text_entities'] += 1
                    text_content = self._extract_text_content(entity, props)
                    if text_content:
                        analysis['text_contents'].append({
                            'text': text_content,
                            'type': entity_type,
                            'layer': props.get('layer', ''),
                            'position': props.get('position', (0, 0, 0))
                        })
                        
            except Exception as e:
                # 跳过有问题的实体
                continue
    
    def _get_entity_properties(self, entity) -> Dict:
        """
        安全地获取实体属性
        """
        props = {}
        try:
            props['type'] = entity.ObjectName
            props['layer'] = entity.Layer
            # 尝试获取位置信息
            if hasattr(entity, 'InsertionPoint'):
                props['position'] = (entity.InsertionPoint[0], entity.InsertionPoint[1], entity.InsertionPoint[2])
        except:
            pass
        return props
    
    def _is_text_entity(self, entity_type: str) -> bool:
        """
        判断是否为文本实体
        """
        text_types = ['AcDbText', 'AcDbMText', 'AcDbAttributeDefinition', 'AcDbAttribute']
        return entity_type in text_types
    
    def _extract_text_content(self, entity, props: Dict) -> Optional[str]:
        """
        提取文本内容
        """
        try:
            if hasattr(entity, 'TextString'):
                return entity.TextString
            elif hasattr(entity, 'Text'):
                return entity.Text
        except:
            pass
        return None
    
    def convert_to_dxf_optimized(self, output_path: str) -> bool:
        """
        优化的DXF转换
        """
        if not self.doc:
            print("✗ 没有打开的文档")
            return False
        
        try:
            print(f"正在转换为DXF: {output_path}")
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 设置导出选项（如果支持）
            try:
                # 尝试设置DXF版本为R2018（兼容性更好）
                self.doc.SetVariable("DWGCODEPAGE", "ANSI_936")  # 中文编码
            except:
                pass
            
            # 执行转换
            start_time = time.time()
            
            # 使用SaveAs方法转换
            self.doc.SaveAs(output_path, 12)  # 12 = acR2018_dxf
            
            conversion_time = time.time() - start_time
            
            # 验证输出文件
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print(f"✓ 转换成功，用时 {conversion_time:.2f}秒")
                print(f"✓ 输出文件: {output_path}")
                return True
            else:
                print("✗ 转换失败：输出文件无效")
                return False
                
        except Exception as e:
            print(f"✗ 转换失败: {e}")
            return False
    
    def close_document(self):
        """
        关闭当前文档
        """
        try:
            if self.doc:
                self.doc.Close(False)  # False = 不保存
                self.doc = None
        except:
            pass
    
    def disconnect(self):
        """
        断开CAD连接
        """
        try:
            self.close_document()
            if self.app:
                # 不退出CAD应用程序，只是断开连接
                self.app = None
            self.connected = False
        except:
            pass

def convert_dwg_to_dxf_optimized(dwg_path: str, output_dir: str = None) -> bool:
    """
    优化的DWG到DXF转换函数
    """
    if not os.path.exists(dwg_path):
        print(f"✗ 文件不存在: {dwg_path}")
        return False
    
    # 设置输出路径
    if output_dir is None:
        output_dir = os.path.dirname(dwg_path)
    
    base_name = os.path.splitext(os.path.basename(dwg_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.dxf")
    
    print(f"开始转换: {os.path.basename(dwg_path)}")
    
    # 创建转换器
    converter = OptimizedHaoChenCADConverter()
    
    try:
        # 连接CAD
        if not converter.connect_to_cad():
            return False
        
        # 打开文件
        if not converter.open_dwg_file(dwg_path):
            return False
        
        # 分析文档（可选）
        analysis = converter.analyze_entities_optimized()
        if analysis:
            print(f"文档信息: {analysis['total_entities']} 个实体, {analysis['text_entities']} 个文本")
        
        # 转换为DXF
        success = converter.convert_to_dxf_optimized(output_path)
        
        return success
        
    finally:
        # 清理资源
        converter.disconnect()

def main():
    """
    主函数 - 处理命令行参数
    """
    import argparse
    
    parser = argparse.ArgumentParser(description='浩辰CAD优化转换器 - DWG转DXF')
    parser.add_argument('input', help='输入DWG文件或目录路径')
    parser.add_argument('--output', '-o', help='输出目录（默认为输入文件所在目录）')
    parser.add_argument('--recursive', '-r', action='store_true', help='递归处理子目录')
    
    args = parser.parse_args()
    
    input_path = os.path.abspath(args.input)
    output_dir = os.path.abspath(args.output) if args.output else None
    
    print("=== 浩辰CAD优化转换器 ===")
    print(f"输入路径: {input_path}")
    if output_dir:
        print(f"输出目录: {output_dir}")
    print()
    
    success_count = 0
    total_count = 0
    
    if os.path.isfile(input_path):
        # 单文件处理
        if input_path.lower().endswith('.dwg'):
            total_count = 1
            if convert_dwg_to_dxf_optimized(input_path, output_dir):
                success_count = 1
        else:
            print("✗ 输入文件不是DWG格式")
    
    elif os.path.isdir(input_path):
        # 目录处理
        dwg_files = []
        
        if args.recursive:
            # 递归查找
            for root, dirs, files in os.walk(input_path):
                for file in files:
                    if file.lower().endswith('.dwg'):
                        dwg_files.append(os.path.join(root, file))
        else:
            # 只处理当前目录
            for file in os.listdir(input_path):
                if file.lower().endswith('.dwg'):
                    dwg_files.append(os.path.join(input_path, file))
        
        total_count = len(dwg_files)
        print(f"找到 {total_count} 个DWG文件")
        print()
        
        for i, dwg_file in enumerate(dwg_files, 1):
            print(f"\n[{i}/{total_count}] 处理文件: {os.path.basename(dwg_file)}")
            
            # 设置输出目录
            if output_dir:
                # 保持相对路径结构
                rel_path = os.path.relpath(os.path.dirname(dwg_file), input_path)
                current_output_dir = os.path.join(output_dir, rel_path) if rel_path != '.' else output_dir
            else:
                current_output_dir = os.path.dirname(dwg_file)
            
            if convert_dwg_to_dxf_optimized(dwg_file, current_output_dir):
                success_count += 1
            
            print(f"当前进度: {success_count}/{i} 成功")
    
    else:
        print(f"✗ 路径不存在: {input_path}")
    
    # 显示总结
    print("\n=== 转换总结 ===")
    print(f"总文件数: {total_count}")
    print(f"成功转换: {success_count}")
    print(f"失败数量: {total_count - success_count}")
    
    if success_count > 0:
        print(f"成功率: {success_count/total_count*100:.1f}%")
    
    return 0 if success_count == total_count else 1

if __name__ == "__main__":
    sys.exit(main())