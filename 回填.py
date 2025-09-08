import os
import pandas as pd
import ezdxf
import re
from pathlib import Path
from datetime import datetime
from logger_config import get_logger

# 初始化日志记录器
logger = get_logger("dxf_backfill")

def smart_translate(text, translation_map):
    """
    智能翻译函数，支持多种空格处理策略
    
    Args:
        text: 要翻译的文本
        translation_map: 翻译映射表
    
    Returns:
        tuple: (翻译结果, 使用的方法) 或 (None, 原因) 表示跳过
    """
    if not text or not isinstance(text, str):
        return None, '无效文本'
    
    # 首先尝试直接匹配
    if text in translation_map:
        translated = translation_map[text]
        # 检查翻译是否为空
        if not translated or not translated.strip():
            return None, '翻译为空'
        return translated, '直接匹配'
    
    # 定义标准化方法
    normalization_methods = [
        ('移除空格', lambda x: re.sub(r'\s+', '', x)),
        ('单空格', lambda x: re.sub(r'\s+', ' ', x.strip())),
        ('去首尾空格', lambda x: x.strip())
    ]
    
    # 尝试各种标准化方法
    for method_name, method_func in normalization_methods:
        normalized_text = method_func(text)
        
        # 在翻译映射表中查找标准化后的文本
        for original, translated in translation_map.items():
            if method_func(original) == normalized_text:
                # 检查翻译是否为空
                if not translated or not translated.strip():
                    return None, '翻译为空'
                logger.debug(f"智能翻译匹配: '{text}' -> '{translated}' (通过{method_name}匹配'{original}')")
                return translated, f'{method_name}匹配({original})'
    
    return None, '未找到匹配'

def load_translation_map(excel_path):
    """从Excel加载翻译映射表"""
    logger.info(f"开始加载翻译映射表: {excel_path}")
    try:
        df = pd.read_excel(excel_path)
        logger.info(f"Excel文件读取成功，共 {len(df)} 行数据")
        
        # 创建翻译映射，过滤掉空翻译
        translation_map = {}
        for _, row in df.iterrows():
            # 智能检测列索引，支持不同的Excel格式
            if len(row) >= 3:
                # 3列格式：可能是序号、原文、译文
                original = str(row.iloc[1]).strip()  # 原文（第2列）
                translated = row.iloc[2]  # 译文（第3列）
            elif len(row) >= 2:
                # 2列格式：原文、译文
                original = str(row.iloc[0]).strip()  # 原文（第1列）
                translated = row.iloc[1]  # 译文（第2列）
            else:
                logger.warning(f"跳过格式不正确的行: {row.values}")
                continue
            
            # 检查翻译是否有效（不为空、不为NaN、不为None等无效值）
            if pd.notna(translated):
                translated_str = str(translated).strip()
                # 检查是否为有效翻译（不为空且不是常见的无效值）
                invalid_values = ['', 'nan', 'none', 'null', 'n/a', 'na']
                if translated_str and translated_str.lower() not in invalid_values:
                    translation_map[original] = translated_str
                else:
                    logger.debug(f"跳过无效翻译: '{original}' -> '{translated}'")
            else:
                logger.debug(f"跳过空翻译: '{original}' -> '{translated}'")
        
        logger.info(f"翻译映射表加载成功，共 {len(translation_map)} 条有效翻译映射")
        return translation_map
    except FileNotFoundError:
        logger.error(f"翻译文件 '{excel_path}' 未找到")
        print(f"错误：翻译文件 '{excel_path}' 未找到。")
        return {}
    except Exception as e:
        logger.error(f"读取Excel文件时出错: {str(e)}")
        print(f"读取Excel文件时出错: {e}")
        return {}

def translate_text_entity(owner, entity, translation_map, font_name="Times New Roman", replace_mode=False, font_size_reduction=4):
    """
    翻译文本实体
    
    Args:
        owner: 实体所有者（模型空间、图纸空间或块定义）
        entity: 文本实体
        translation_map: 翻译映射表
        font_name: 字体名称
        replace_mode: 是否使用替换模式（True）还是新建模式（False）
        font_size_reduction: 字体大小减少值
    
    Returns:
        dict: 翻译结果统计
    """
    result = {
        'processed': 0,
        'translated': 0,
        'skipped': 0,
        'errors': 0
    }
    
    try:
        result['processed'] += 1
        
        # 获取原始文本
        original_text = entity.dxf.text if hasattr(entity.dxf, 'text') else ''
        
        if not original_text or not original_text.strip():
            result['skipped'] += 1
            return result
        
        # 尝试智能翻译
        translated_text, method = smart_translate(original_text, translation_map)
        
        if translated_text is None:
            logger.debug(f"跳过文本 '{original_text}': {method}")
            result['skipped'] += 1
            return result
        
        logger.info(f"翻译文本: '{original_text}' -> '{translated_text}' ({method})")
        
        if replace_mode:
            # 替换模式：直接修改原实体
            entity.dxf.text = translated_text
            
            # 设置字体
            if hasattr(entity.dxf, 'style'):
                # 查找或创建文本样式
                style_name = f"TranslatedStyle_{font_name.replace(' ', '_')}"
                
                # 检查样式是否存在
                if style_name not in owner.doc.styles:
                    # 创建新样式
                    style = owner.doc.styles.new(style_name)
                    style.dxf.font = font_name
                    style.dxf.width = 0.8  # 字体宽度因子
                
                entity.dxf.style = style_name
            
            # 调整字体大小
            if hasattr(entity.dxf, 'height'):
                original_height = entity.dxf.height
                new_height = max(original_height - font_size_reduction, 1.0)
                entity.dxf.height = new_height
                logger.debug(f"字体大小调整: {original_height} -> {new_height}")
        
        else:
            # 新建模式：创建新的文本实体
            # 获取原实体的位置和属性
            insert_point = entity.dxf.insert if hasattr(entity.dxf, 'insert') else (0, 0, 0)
            height = entity.dxf.height if hasattr(entity.dxf, 'height') else 2.5
            rotation = entity.dxf.rotation if hasattr(entity.dxf, 'rotation') else 0
            
            # 调整字体大小
            new_height = max(height - font_size_reduction, 1.0)
            
            # 创建新的文本实体
            new_entity = owner.add_text(
                text=translated_text,
                dxfattribs={
                    'insert': insert_point,
                    'height': new_height,
                    'rotation': rotation,
                    'layer': entity.dxf.layer if hasattr(entity.dxf, 'layer') else '0'
                }
            )
            
            # 设置字体样式
            style_name = f"TranslatedStyle_{font_name.replace(' ', '_')}"
            if style_name not in owner.doc.styles:
                style = owner.doc.styles.new(style_name)
                style.dxf.font = font_name
                style.dxf.width = 0.8
            
            new_entity.dxf.style = style_name
            
            # 删除原实体
            owner.delete_entity(entity)
            
            logger.debug(f"创建新文本实体: '{translated_text}' at {insert_point}")
        
        result['translated'] += 1
        
    except Exception as e:
        logger.error(f"翻译文本实体时出错: {e}")
        result['errors'] += 1
    
    return result

def translate_dwg(dwg_path, translation_map, font_name="Times New Roman", replace_mode=False, font_size_reduction=4):
    """
    翻译DWG文件中的文本
    
    Args:
        dwg_path: DWG文件路径
        translation_map: 翻译映射表
        font_name: 字体名称
        replace_mode: 是否使用替换模式
        font_size_reduction: 字体大小减少值
    
    Returns:
        dict: 处理结果统计
    """
    logger.info(f"开始处理文件: {dwg_path}")
    
    result = {
        'file': dwg_path,
        'processed': 0,
        'translated': 0,
        'skipped': 0,
        'errors': 0,
        'success': False
    }
    
    try:
        # 打开DXF文件
        doc = ezdxf.readfile(dwg_path)
        logger.info(f"成功打开文件: {dwg_path}")
        
        # 处理模型空间
        msp = doc.modelspace()
        logger.info("开始处理模型空间")
        
        # 查找所有文本实体
        text_entities = msp.query('TEXT MTEXT')
        logger.info(f"在模型空间找到 {len(text_entities)} 个文本实体")
        
        for entity in text_entities:
            entity_result = translate_text_entity(
                msp, entity, translation_map, font_name, replace_mode, font_size_reduction
            )
            
            # 累加统计结果
            for key in ['processed', 'translated', 'skipped', 'errors']:
                result[key] += entity_result[key]
        
        # 处理图纸空间
        for layout in doc.layouts:
            if layout.name != 'Model':  # 跳过模型空间
                logger.info(f"开始处理图纸空间: {layout.name}")
                
                text_entities = layout.query('TEXT MTEXT')
                logger.info(f"在图纸空间 {layout.name} 找到 {len(text_entities)} 个文本实体")
                
                for entity in text_entities:
                    entity_result = translate_text_entity(
                        layout, entity, translation_map, font_name, replace_mode, font_size_reduction
                    )
                    
                    # 累加统计结果
                    for key in ['processed', 'translated', 'skipped', 'errors']:
                        result[key] += entity_result[key]
        
        # 处理块定义中的文本
        logger.info("开始处理块定义")
        for block_name, block in doc.blocks.items():
            if not block_name.startswith('*'):  # 跳过匿名块
                text_entities = block.query('TEXT MTEXT')
                if text_entities:
                    logger.info(f"在块 {block_name} 中找到 {len(text_entities)} 个文本实体")
                    
                    for entity in text_entities:
                        entity_result = translate_text_entity(
                            block, entity, translation_map, font_name, replace_mode, font_size_reduction
                        )
                        
                        # 累加统计结果
                        for key in ['processed', 'translated', 'skipped', 'errors']:
                            result[key] += entity_result[key]
        
        # 保存文件
        output_path = dwg_path.replace('.dxf', '_translated.dxf')
        doc.saveas(output_path)
        logger.info(f"翻译完成，已保存到: {output_path}")
        
        result['success'] = True
        result['output_path'] = output_path
        
    except Exception as e:
        logger.error(f"处理文件 {dwg_path} 时出错: {e}")
        result['errors'] += 1
        result['error_message'] = str(e)
    
    return result

def process_directory(directory, translation_map, output_folder, font_name="Times New Roman", replace_mode=False, font_size_reduction=4):
    """
    处理目录中的所有DXF文件
    
    Args:
        directory: 输入目录
        translation_map: 翻译映射表
        output_folder: 输出目录
        font_name: 字体名称
        replace_mode: 是否使用替换模式
        font_size_reduction: 字体大小减少值
    
    Returns:
        list: 处理结果列表
    """
    logger.info(f"开始处理目录: {directory}")
    
    # 确保输出目录存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 查找所有DXF文件
    dxf_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.dxf'):
                dxf_files.append(os.path.join(root, file))
    
    logger.info(f"找到 {len(dxf_files)} 个DXF文件")
    
    results = []
    for i, dxf_file in enumerate(dxf_files, 1):
        print(f"\n处理文件 {i}/{len(dxf_files)}: {os.path.basename(dxf_file)}")
        
        result = translate_dwg(dxf_file, translation_map, font_name, replace_mode, font_size_reduction)
        results.append(result)
        
        # 显示处理结果
        if result['success']:
            print(f"✓ 成功 - 处理: {result['processed']}, 翻译: {result['translated']}, 跳过: {result['skipped']}")
        else:
            print(f"✗ 失败 - {result.get('error_message', '未知错误')}")
    
    return results

def main():
    """
    主函数
    """
    import argparse
    
    parser = argparse.ArgumentParser(description='DXF文件翻译回填工具')
    parser.add_argument('directory', nargs='?', default='.', help='要处理的目录（默认为当前目录）')
    parser.add_argument('--excel', '-e', default='extracted_texts.xlsx', help='翻译Excel文件路径')
    parser.add_argument('--font', '-f', default='Times New Roman', help='目标字体名称')
    parser.add_argument('--output', '-o', help='输出目录（默认为输入目录下的translated子目录）')
    parser.add_argument('--replace', '-r', action='store_true', help='使用替换模式（直接修改原文本）')
    parser.add_argument('--font-reduction', type=int, default=4, help='字体大小减少值（默认4）')
    
    args = parser.parse_args()
    
    # 设置工作目录
    work_dir = os.path.abspath(args.directory)
    if not os.path.exists(work_dir):
        print(f"错误：目录 '{work_dir}' 不存在")
        return 1
    
    # 设置输出目录
    if args.output:
        output_dir = os.path.abspath(args.output)
    else:
        output_dir = os.path.join(work_dir, 'translated')
    
    # 设置Excel文件路径
    excel_path = args.excel
    if not os.path.isabs(excel_path):
        excel_path = os.path.join(work_dir, excel_path)
    
    print(f"工作目录: {work_dir}")
    print(f"翻译文件: {excel_path}")
    print(f"输出目录: {output_dir}")
    print(f"字体设置: {args.font}")
    print(f"处理模式: {'替换模式' if args.replace else '新建模式'}")
    print(f"字体减少: {args.font_reduction}")
    print()
    
    # 加载翻译映射表
    translation_map = load_translation_map(excel_path)
    if not translation_map:
        print("错误：无法加载翻译映射表")
        return 1
    
    print(f"加载了 {len(translation_map)} 条翻译映射")
    print()
    
    # 处理目录
    results = process_directory(
        work_dir, 
        translation_map, 
        output_dir, 
        args.font, 
        args.replace, 
        args.font_reduction
    )
    
    # 显示总结
    print("\n=== 处理总结 ===")
    total_processed = sum(r['processed'] for r in results)
    total_translated = sum(r['translated'] for r in results)
    total_skipped = sum(r['skipped'] for r in results)
    total_errors = sum(r['errors'] for r in results)
    successful_files = sum(1 for r in results if r['success'])
    
    print(f"文件总数: {len(results)}")
    print(f"成功文件: {successful_files}")
    print(f"处理文本: {total_processed}")
    print(f"翻译文本: {total_translated}")
    print(f"跳过文本: {total_skipped}")
    print(f"错误数量: {total_errors}")
    
    if successful_files > 0:
        print(f"\n翻译文件已保存到: {output_dir}")
    
    logger.info(f"处理完成 - 成功: {successful_files}/{len(results)}, 翻译: {total_translated}, 错误: {total_errors}")
    
    return 0 if total_errors == 0 else 1

if __name__ == "__main__":
    main()