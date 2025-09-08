#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD文件翻译处理主控制脚本

功能:
1. 将DWG文件转换为DXF文件
2. 从DXF文件中提取文本
3. 等待用户填写翻译
4. 将翻译回填到DXF文件中

作者: AI Assistant
日期: 2024
"""

import os
import sys
from pathlib import Path
import subprocess
import time
from logger_config import get_logger

# 初始化日志记录器
logger = get_logger("main_processor")

def print_banner():
    """打印程序横幅"""
    print("="*60)
    print("           CAD文件翻译处理系统")
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
            logger.debug(f"依赖库检查通过: {module}")
            print(f"✓ {module}")
        except ImportError:
            logger.warning(f"依赖库缺失: {module}")
            print(f"✗ {module} (缺失)")
            missing_modules.append(module)
    
    if missing_modules:
        logger.error(f"缺少依赖库: {', '.join(missing_modules)}")
        print(f"\n缺少以下依赖库: {', '.join(missing_modules)}")
        print("请运行以下命令安装:")
        print(f"pip install {' '.join(missing_modules)}")
        return False, missing_modules
    
    logger.info("所有依赖库检查完成")
    print("✓ 所有依赖库已安装")
    return True, []

def run_script(script_name, description):
    """运行指定的Python脚本"""
    logger.info(f"准备运行脚本: {script_name} - {description}")
    script_path = Path(script_name)
    
    if not script_path.exists():
        logger.error(f"脚本文件不存在: {script_name}")
        print(f"✗ 错误: 脚本文件 {script_name} 不存在")
        return False
    
    print(f"\n{description}...")
    print(f"运行脚本: {script_name}")
    print("-" * 40)
    
    try:
        logger.debug(f"开始执行脚本: {script_name}")
        # 使用subprocess运行脚本
        result = subprocess.run([sys.executable, script_name], 
                              capture_output=False, 
                              text=True, 
                              cwd=Path.cwd())
        
        if result.returncode == 0:
            logger.info(f"脚本执行成功: {script_name}")
            print("-" * 40)
            print(f"✓ {description}完成")
            return True
        else:
            logger.error(f"脚本执行失败: {script_name}, 退出码: {result.returncode}")
            print("-" * 40)
            print(f"✗ {description}失败 (退出码: {result.returncode})")
            return False
            
    except Exception as e:
        logger.error(f"运行脚本时发生异常: {script_name}, 错误: {e}")
        print("-" * 40)
        print(f"✗ {description}失败: {e}")
        return False

def wait_for_user_input(message, timeout=None):
    """等待用户输入
    
    Args:
        message: 提示信息
        timeout: 超时时间（秒），None表示无限等待
    
    Returns:
        str: 用户输入的内容
    """
    logger.info(f"等待用户输入: {message}")
    print(f"\n{message}")
    
    if timeout:
        print(f"(将在 {timeout} 秒后自动继续)")
        
    try:
        if timeout:
            # 简单的超时实现
            import signal
            
            def timeout_handler(signum, frame):
                raise TimeoutError("输入超时")
            
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout)
            
            try:
                user_input = input("请输入 (回车继续): ")
                signal.alarm(0)  # 取消超时
                return user_input
            except TimeoutError:
                print("\n输入超时，自动继续...")
                return ""
        else:
            return input("请输入 (回车继续): ")
            
    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print("\n用户中断程序")
        sys.exit(0)
    except Exception as e:
        logger.error(f"等待用户输入时发生错误: {e}")
        return ""

def check_file_exists(file_path, description):
    """检查文件是否存在
    
    Args:
        file_path: 文件路径
        description: 文件描述
    
    Returns:
        bool: 文件是否存在
    """
    path = Path(file_path)
    if path.exists():
        logger.info(f"{description}文件存在: {file_path}")
        print(f"✓ {description}文件存在: {file_path}")
        return True
    else:
        logger.warning(f"{description}文件不存在: {file_path}")
        print(f"✗ {description}文件不存在: {file_path}")
        return False

def main():
    """主函数"""
    logger.info("CAD文件翻译处理系统启动")
    print_banner()
    
    # 1. 检查依赖
    deps_ok, missing = check_dependencies()
    if not deps_ok:
        logger.error("依赖检查失败，程序退出")
        return 1
    
    print("\n" + "="*60)
    print("开始处理流程")
    print("="*60)
    
    # 2. 步骤1: DWG转DXF
    print("\n步骤 1: DWG文件转换为DXF文件")
    print("-" * 30)
    
    # 检查是否有转换脚本
    converter_script = "haochen_optimized_converter.py"
    if check_file_exists(converter_script, "DWG转换脚本"):
        success = run_script(converter_script, "DWG转DXF转换")
        if not success:
            logger.warning("DWG转换失败，但继续执行后续步骤")
            print("⚠ DWG转换失败，请手动转换或直接使用DXF文件")
    else:
        logger.info("未找到DWG转换脚本，跳过转换步骤")
        print("ℹ 未找到DWG转换脚本，请确保有DXF文件可用")
    
    # 3. 步骤2: 提取文本
    print("\n步骤 2: 从DXF文件提取文本")
    print("-" * 30)
    
    extractor_script = "提取.py"
    if check_file_exists(extractor_script, "文本提取脚本"):
        success = run_script(extractor_script, "文本提取")
        if not success:
            logger.error("文本提取失败")
            print("✗ 文本提取失败，无法继续")
            return 1
    else:
        logger.error("未找到文本提取脚本")
        print("✗ 未找到文本提取脚本")
        return 1
    
    # 4. 等待用户填写翻译
    print("\n步骤 3: 等待用户填写翻译")
    print("-" * 30)
    
    # 检查是否生成了Excel文件
    excel_files = list(Path('.').glob('*.xlsx'))
    if excel_files:
        logger.info(f"找到Excel文件: {[f.name for f in excel_files]}")
        print(f"✓ 找到Excel文件: {', '.join([f.name for f in excel_files])}")
        print("\n请打开Excel文件，在'译文'列中填写翻译内容")
        print("填写完成后保存文件，然后回到这里继续")
        
        wait_for_user_input("翻译填写完成后，按回车键继续...")
    else:
        logger.warning("未找到Excel文件")
        print("⚠ 未找到Excel文件，请检查文本提取是否成功")
        
        user_choice = wait_for_user_input("是否继续执行回填步骤？(y/n): ")
        if user_choice.lower() not in ['y', 'yes', '']:
            logger.info("用户选择不继续")
            print("程序结束")
            return 0
    
    # 5. 步骤4: 回填翻译
    print("\n步骤 4: 将翻译回填到DXF文件")
    print("-" * 30)
    
    backfill_script = "回填.py"
    if check_file_exists(backfill_script, "翻译回填脚本"):
        success = run_script(backfill_script, "翻译回填")
        if success:
            logger.info("翻译回填完成")
            print("\n" + "="*60)
            print("           处理完成！")
            print("="*60)
            print("\n翻译后的DXF文件已生成，请检查结果")
        else:
            logger.error("翻译回填失败")
            print("✗ 翻译回填失败")
            return 1
    else:
        logger.error("未找到翻译回填脚本")
        print("✗ 未找到翻译回填脚本")
        return 1
    
    logger.info("CAD文件翻译处理系统完成")
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