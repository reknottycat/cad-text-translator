#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化版CAD文件翻译处理脚本

适用于以下情况：
1. 无法安装ODA File Converter
2. 已经有DXF文件
3. 需要手动转换DWG文件

作者: AI Assistant
日期: 2024
"""

import os
import sys
from pathlib import Path
import subprocess
from logger_config import get_logger

# 初始化日志记录器
logger = get_logger("simple_processor")

def print_banner():
    """打印程序横幅"""
    print("="*60)
    print("        简化版CAD文件翻译处理系统")
    print("="*60)
    print()

def check_dependencies():
    """检查必要的依赖
    
    Returns:
        tuple: (bool, list) - (是否所有依赖都满足, 缺失的依赖列表)
    """
    logger.info("开始检查依赖库")
    print("检查依赖库...")
    
    required_modules = ['ezdxf', 'pandas', 'openpyxl']
    missing_modules = []
    
    for module in required_modules:
        try:
            __import__(module)
            logger.debug(f"依赖库 {module} 检查通过")
            print(f"✓ {module}")
        except ImportError:
            logger.warning(f"依赖库 {module} 缺失")
            print(f"✗ {module} (缺失)")
            missing_modules.append(module)
    
    if missing_modules:
        logger.error(f"缺少依赖库: {missing_modules}")
        print(f"\n缺少以下依赖库: {', '.join(missing_modules)}")
        print("请运行以下命令安装:")
        print(f"pip install {' '.join(missing_modules)}")
        return False, missing_modules
    
    logger.info("所有依赖库检查完成")
    print("✓ 所有依赖库已安装")
    return True, []

def check_files():
    """检查文件情况"""
    logger.info("开始检查文件")
    print("\n检查文件...")
    
    # 检查DXF文件
    dxf_files = list(Path('.').glob('*.dxf')) + list(Path('.').glob('*.DXF'))
    dwg_files = list(Path('.').glob('*.dwg')) + list(Path('.').glob('*.DWG'))
    
    logger.debug(f"扫描到 {len(dxf_files)} 个DXF文件，{len(dwg_files)} 个DWG文件")
    
    if dxf_files:
        logger.info(f"找到 {len(dxf_files)} 个DXF文件")
        print(f"✓ 找到 {len(dxf_files)} 个DXF文件:")
        for dxf in dxf_files[:5]:  # 只显示前5个
            print(f"  - {dxf.name}")
        if len(dxf_files) > 5:
            print(f"  ... 还有 {len(dxf_files) - 5} 个文件")
        return True, 'dxf'
    elif dwg_files:
        logger.info(f"找到 {len(dwg_files)} 个DWG文件")
        print(f"✓ 找到 {len(dwg_files)} 个DWG文件:")
        for dwg in dwg_files[:5]:  # 只显示前5个
            print(f"  - {dwg.name}")
        if len(dwg_files) > 5:
            print(f"  ... 还有 {len(dwg_files) - 5} 个文件")
        return True, 'dwg'
    else:
        logger.warning("未找到DXF或DWG文件")
        print("✗ 未找到DXF或DWG文件")
        return False, None

def provide_dwg_conversion_guide():
    """提供DWG转换指导"""
    print("\n" + "="*60)
    print("           DWG文件转换指导")
    print("="*60)
    print("\n由于检测到DWG文件，您需要先将其转换为DXF格式。")
    print("\n推荐的转换方法：")
    print("\n1. 使用AutoCAD或浩辰CAD：")
    print("   - 打开DWG文件")
    print("   - 选择 文件 > 另存为")
    print("   - 选择DXF格式保存")
    print("\n2. 使用免费的DWG查看器：")
    print("   - DWG TrueView (Autodesk官方)")
    print("   - LibreCAD (开源)")
    print("   - FreeCAD (开源)")
    print("\n3. 在线转换工具：")
    print("   - CloudConvert")
    print("   - Zamzar")
    print("   - 注意：在线工具可能有文件大小限制")
    print("\n4. 命令行工具：")
    print("   - ODA File Converter (免费，需注册)")
    print("   - 下载地址: https://www.opendesign.com/guestfiles/oda_file_converter")
    print("\n转换完成后，将DXF文件放在当前目录，然后重新运行此程序。")
    print("="*60)

def run_extraction():
    """运行文本提取"""
    logger.info("开始运行文本提取")
    print("\n" + "="*60)
    print("           开始文本提取")
    print("="*60)
    
    # 检查提取脚本
    extractor_script = "提取.py"
    if not Path(extractor_script).exists():
        # 尝试其他可能的脚本名
        possible_scripts = ["extract_texts.py", "dxf_text_extractor.py", "text_extractor.py"]
        for script in possible_scripts:
            if Path(script).exists():
                extractor_script = script
                break
        else:
            logger.error("未找到文本提取脚本")
            print("✗ 未找到文本提取脚本")
            print("请确保以下文件之一存在：")
            print("  - 提取.py")
            print("  - extract_texts.py")
            print("  - dxf_text_extractor.py")
            return False
    
    logger.info(f"使用提取脚本: {extractor_script}")
    print(f"使用提取脚本: {extractor_script}")
    
    try:
        logger.debug(f"开始执行提取脚本: {extractor_script}")
        result = subprocess.run([sys.executable, extractor_script], 
                              capture_output=False, 
                              text=True)
        
        if result.returncode == 0:
            logger.info("文本提取成功")
            print("\n✓ 文本提取完成")
            return True
        else:
            logger.error(f"文本提取失败，退出码: {result.returncode}")
            print(f"\n✗ 文本提取失败 (退出码: {result.returncode})")
            return False
            
    except Exception as e:
        logger.error(f"运行提取脚本时发生异常: {e}")
        print(f"\n✗ 运行提取脚本时发生错误: {e}")
        return False

def check_excel_output():
    """检查Excel输出文件"""
    logger.info("检查Excel输出文件")
    excel_files = list(Path('.').glob('*.xlsx'))
    
    if excel_files:
        logger.info(f"找到Excel文件: {[f.name for f in excel_files]}")
        print(f"\n✓ 生成了Excel文件: {', '.join([f.name for f in excel_files])}")
        return True, excel_files
    else:
        logger.warning("未找到Excel输出文件")
        print("\n✗ 未找到Excel输出文件")
        return False, []

def provide_translation_guide(excel_files):
    """提供翻译指导"""
    print("\n" + "="*60)
    print("           翻译填写指导")
    print("="*60)
    print("\n请按照以下步骤填写翻译：")
    print("\n1. 打开Excel文件：")
    for excel_file in excel_files:
        print(f"   - {excel_file.name}")
    print("\n2. 在'译文'列中填写对应的翻译内容")
    print("\n3. 翻译注意事项：")
    print("   - 保持技术术语的准确性")
    print("   - 注意尺寸标注的格式")
    print("   - 如果某些内容不需要翻译，可以留空")
    print("\n4. 保存Excel文件")
    print("\n5. 返回此程序继续下一步")
    print("="*60)

def wait_for_translation():
    """等待用户完成翻译"""
    logger.info("等待用户完成翻译")
    print("\n请完成翻译后按回车键继续...")
    try:
        input()
        logger.info("用户确认翻译完成")
        return True
    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print("\n程序被用户中断")
        return False

def run_backfill():
    """运行翻译回填"""
    logger.info("开始运行翻译回填")
    print("\n" + "="*60)
    print("           开始翻译回填")
    print("="*60)
    
    # 检查回填脚本
    backfill_script = "回填.py"
    if not Path(backfill_script).exists():
        # 尝试其他可能的脚本名
        possible_scripts = ["backfill.py", "translate_back.py", "fill_translation.py"]
        for script in possible_scripts:
            if Path(script).exists():
                backfill_script = script
                break
        else:
            logger.error("未找到翻译回填脚本")
            print("✗ 未找到翻译回填脚本")
            print("请确保以下文件之一存在：")
            print("  - 回填.py")
            print("  - backfill.py")
            print("  - translate_back.py")
            return False
    
    logger.info(f"使用回填脚本: {backfill_script}")
    print(f"使用回填脚本: {backfill_script}")
    
    try:
        logger.debug(f"开始执行回填脚本: {backfill_script}")
        result = subprocess.run([sys.executable, backfill_script], 
                              capture_output=False, 
                              text=True)
        
        if result.returncode == 0:
            logger.info("翻译回填成功")
            print("\n✓ 翻译回填完成")
            return True
        else:
            logger.error(f"翻译回填失败，退出码: {result.returncode}")
            print(f"\n✗ 翻译回填失败 (退出码: {result.returncode})")
            return False
            
    except Exception as e:
        logger.error(f"运行回填脚本时发生异常: {e}")
        print(f"\n✗ 运行回填脚本时发生错误: {e}")
        return False

def main():
    """主函数"""
    logger.info("简化版CAD文件翻译处理系统启动")
    print_banner()
    
    # 1. 检查依赖
    deps_ok, missing = check_dependencies()
    if not deps_ok:
        logger.error("依赖检查失败，程序退出")
        return 1
    
    # 2. 检查文件
    files_ok, file_type = check_files()
    if not files_ok:
        logger.error("未找到可处理的文件")
        print("\n请将DXF或DWG文件放在当前目录中，然后重新运行程序。")
        return 1
    
    # 3. 如果是DWG文件，提供转换指导
    if file_type == 'dwg':
        provide_dwg_conversion_guide()
        logger.info("提供DWG转换指导后退出")
        return 0
    
    # 4. 运行文本提取
    if not run_extraction():
        logger.error("文本提取失败")
        return 1
    
    # 5. 检查Excel输出
    excel_ok, excel_files = check_excel_output()
    if not excel_ok:
        logger.error("未生成Excel文件")
        return 1
    
    # 6. 提供翻译指导
    provide_translation_guide(excel_files)
    
    # 7. 等待用户完成翻译
    if not wait_for_translation():
        logger.info("用户中断程序")
        return 0
    
    # 8. 运行翻译回填
    if not run_backfill():
        logger.error("翻译回填失败")
        return 1
    
    # 9. 完成
    print("\n" + "="*60)
    print("           处理完成！")
    print("="*60)
    print("\n翻译后的DXF文件已生成，请检查结果。")
    
    logger.info("简化版CAD文件翻译处理系统完成")
    return 0

if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n程序被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n程序发生未预期的错误: {e}")
        logger.error(f"程序发生未预期的错误: {e}")
        sys.exit(1)