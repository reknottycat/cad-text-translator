#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DXF文件清理工具
用于修复DXF文件中的无效数据，特别是SEQEND实体的layer属性问题

作者: AI Assistant
日期: 2024
"""

import ezdxf
import re
from pathlib import Path
import shutil
import os

def clean_layer_name(layer_name):
    """
    清理图层名称，移除无效字符
    
    Args:
        layer_name: 原始图层名称
        
    Returns:
        清理后的图层名称
    """
    if not layer_name or not isinstance(layer_name, str):
        return "0"  # 默认图层
    
    # 移除非ASCII字符和问号
    cleaned = re.sub(r'[^\w\-_.]', '', layer_name)
    
    # 如果清理后为空或只包含无效字符，使用默认图层
    if not cleaned or cleaned.isspace():
        return "0"
    
    # 确保图层名不以数字开头（DXF规范）
    if cleaned[0].isdigit():
        cleaned = "Layer_" + cleaned
    
    # 限制长度（DXF图层名最大255字符）
    return cleaned[:255]

def fix_seqend_entities(doc):
    """
    修复SEQEND实体的无效属性
    
    Args:
        doc: ezdxf文档对象
        
    Returns:
        修复的实体数量
    """
    fixed_count = 0
    
    # 遍历所有布局（模型空间和图纸空间）
    for layout in doc.layouts:
        for entity in layout:
            if entity.dxftype() == 'SEQEND':
                try:
                    # 检查并修复layer属性
                    if hasattr(entity.dxf, 'layer'):
                        original_layer = entity.dxf.layer
                        cleaned_layer = clean_layer_name(original_layer)
                        
                        if original_layer != cleaned_layer:
                            print(f"修复SEQEND实体图层: '{original_layer}' -> '{cleaned_layer}'")
                            entity.dxf.layer = cleaned_layer
                            fixed_count += 1
                            
                    # 确保SEQEND实体有必要的属性
                    if not hasattr(entity.dxf, 'layer') or not entity.dxf.layer:
                        entity.dxf.layer = "0"
                        fixed_count += 1
                        
                except Exception as e:
                    print(f"修复SEQEND实体时出错: {e}")
                    # 尝试设置默认值
                    try:
                        entity.dxf.layer = "0"
                        fixed_count += 1
                    except:
                        pass
    
    return fixed_count

def fix_all_entities(doc):
    """
    修复所有实体的无效图层属性
    
    Args:
        doc: ezdxf文档对象
        
    Returns:
        修复的实体数量
    """
    fixed_count = 0
    
    # 遍历所有布局
    for layout in doc.layouts:
        for entity in layout:
            try:
                # 检查并修复layer属性
                if hasattr(entity.dxf, 'layer'):
                    original_layer = entity.dxf.layer
                    cleaned_layer = clean_layer_name(original_layer)
                    
                    if original_layer != cleaned_layer:
                        print(f"修复实体图层: {entity.dxftype()} '{original_layer}' -> '{cleaned_layer}'")
                        entity.dxf.layer = cleaned_layer
                        fixed_count += 1
                        
                # 确保实体有图层属性
                elif hasattr(entity, 'dxf'):
                    entity.dxf.layer = "0"
                    fixed_count += 1
                    
            except Exception as e:
                print(f"修复实体 {entity.dxftype()} 时出错: {e}")
                # 尝试设置默认值
                try:
                    if hasattr(entity, 'dxf'):
                        entity.dxf.layer = "0"
                        fixed_count += 1
                except:
                    pass
    
    return fixed_count

def clean_dxf_file(input_path, output_path=None, backup=True):
    """
    清理DXF文件
    
    Args:
        input_path: 输入文件路径
        output_path: 输出文件路径（如果为None，则覆盖原文件）
        backup: 是否创建备份
        
    Returns:
        tuple: (是否成功, 修复的实体数量, 错误信息)
    """
    input_path = Path(input_path)
    
    if not input_path.exists():
        return False, 0, f"文件不存在: {input_path}"
    
    if output_path is None:
        output_path = input_path
    else:
        output_path = Path(output_path)
    
    # 创建备份
    if backup and output_path == input_path:
        backup_path = input_path.with_suffix(input_path.suffix + '.backup')
        try:
            shutil.copy2(input_path, backup_path)
            print(f"已创建备份: {backup_path}")
        except Exception as e:
            print(f"创建备份失败: {e}")
    
    try:
        print(f"正在清理文件: {input_path}")
        
        # 读取DXF文件
        doc = ezdxf.readfile(str(input_path))
        
        # 修复SEQEND实体
        seqend_fixed = fix_seqend_entities(doc)
        print(f"修复了 {seqend_fixed} 个SEQEND实体")
        
        # 修复所有实体的图层属性
        all_fixed = fix_all_entities(doc)
        print(f"修复了 {all_fixed} 个实体的图层属性")
        
        total_fixed = seqend_fixed + all_fixed
        
        # 清理文档头部信息
        try:
            # 设置DXF版本
            doc.dxfversion = 'AC1021'  # AutoCAD 2007/2008/2009
            
            # 清理无效的系统变量
            header = doc.header
            
            # 设置基本的头部变量
            header['$ACADVER'] = 'AC1021'
            header['$DWGCODEPAGE'] = 'ANSI_936'  # 中文编码
            
            print("已清理文档头部信息")
            
        except Exception as e:
            print(f"清理头部信息时出错: {e}")
        
        # 验证文档
        try:
            doc.audit()
            print("文档验证通过")
        except Exception as e:
            print(f"文档验证警告: {e}")
        
        # 保存文件
        doc.saveas(str(output_path))
        print(f"文件已保存: {output_path}")
        
        return True, total_fixed, None
        
    except ezdxf.DXFStructureError as e:
        error_msg = f"DXF结构错误: {e}"
        print(error_msg)
        return False, 0, error_msg
        
    except ezdxf.DXFVersionError as e:
        error_msg = f"DXF版本错误: {e}"
        print(error_msg)
        return False, 0, error_msg
        
    except Exception as e:
        error_msg = f"处理文件时出错: {e}"
        print(error_msg)
        return False, 0, error_msg

def clean_directory(directory_path, pattern="*.dxf"):
    """
    清理目录中的所有DXF文件
    
    Args:
        directory_path: 目录路径
        pattern: 文件匹配模式
        
    Returns:
        dict: 处理结果统计
    """
    directory_path = Path(directory_path)
    
    if not directory_path.exists() or not directory_path.is_dir():
        print(f"目录不存在: {directory_path}")
        return {'success': 0, 'failed': 0, 'total_fixed': 0}
    
    # 查找所有DXF文件
    dxf_files = list(directory_path.glob(pattern))
    
    if not dxf_files:
        print(f"在目录 {directory_path} 中未找到匹配的文件")
        return {'success': 0, 'failed': 0, 'total_fixed': 0}
    
    print(f"找到 {len(dxf_files)} 个文件")
    
    results = {'success': 0, 'failed': 0, 'total_fixed': 0}
    
    for i, file_path in enumerate(dxf_files, 1):
        print(f"\n[{i}/{len(dxf_files)}] 处理文件: {file_path.name}")
        
        success, fixed_count, error = clean_dxf_file(file_path)
        
        if success:
            results['success'] += 1
            results['total_fixed'] += fixed_count
            print(f"✓ 成功处理，修复了 {fixed_count} 个问题")
        else:
            results['failed'] += 1
            print(f"✗ 处理失败: {error}")
    
    return results

def main():
    """
    主函数 - 命令行接口
    """
    import argparse
    
    parser = argparse.ArgumentParser(description='DXF文件清理工具')
    parser.add_argument('input', help='输入文件或目录路径')
    parser.add_argument('--output', '-o', help='输出文件路径（仅对单文件有效）')
    parser.add_argument('--no-backup', action='store_true', help='不创建备份文件')
    parser.add_argument('--pattern', default='*.dxf', help='文件匹配模式（仅对目录有效）')
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    
    print("=== DXF文件清理工具 ===")
    print(f"输入路径: {input_path}")
    
    if input_path.is_file():
        # 处理单个文件
        success, fixed_count, error = clean_dxf_file(
            input_path, 
            args.output, 
            not args.no_backup
        )
        
        if success:
            print(f"\n✓ 处理成功，修复了 {fixed_count} 个问题")
            return 0
        else:
            print(f"\n✗ 处理失败: {error}")
            return 1
            
    elif input_path.is_dir():
        # 处理目录
        results = clean_directory(input_path, args.pattern)
        
        print("\n=== 处理结果 ===")
        print(f"成功处理: {results['success']} 个文件")
        print(f"处理失败: {results['failed']} 个文件")
        print(f"总共修复: {results['total_fixed']} 个问题")
        
        return 0 if results['failed'] == 0 else 1
        
    else:
        print(f"\n✗ 路径不存在: {input_path}")
        return 1

if __name__ == "__main__":
    main()