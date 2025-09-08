#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD翻译字体配置工具
用于快速修改回填.py中的默认字体设置
"""

import os
import re
from pathlib import Path
from logger_config import get_logger

# 初始化日志记录器
logger = get_logger("font_config")

# 可用字体列表
AVAILABLE_FONTS = [
    "Times New Roman",
    "Arial", 
    "SimSun",  # 宋体
    "SimHei",  # 黑体
    "Microsoft YaHei",  # 微软雅黑
    "Calibri",
    "Verdana",
    "Tahoma",
    "Georgia",
    "Courier New"
]

def get_current_font():
    """获取当前设置的字体"""
    logger.info("开始获取当前字体设置")
    script_path = Path("回填.py")
    if not script_path.exists():
        logger.error("找不到回填.py文件")
        print("错误：找不到回填.py文件")
        return None
        
    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
        logger.info("成功读取回填.py文件")
    except Exception as e:
        logger.error(f"读取回填.py文件失败: {e}")
        print(f"错误：读取文件失败 - {e}")
        return None
        
    # 查找字体设置行
    pattern = r'font_name = "([^"]+)"\s*# 默认字体'
    match = re.search(pattern, content)
    
    if match:
        current_font = match.group(1)
        logger.info(f"找到当前字体设置: {current_font}")
        return current_font
    else:
        logger.warning("无法找到字体设置行")
        print("警告：无法找到字体设置行")
        return None

def set_font(new_font):
    """设置新的默认字体"""
    logger.info(f"开始设置字体为: {new_font}")
    script_path = Path("回填.py")
    if not script_path.exists():
        logger.error("找不到回填.py文件")
        print("错误：找不到回填.py文件")
        return False
        
    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
        logger.info("成功读取回填.py文件")
    except Exception as e:
        logger.error(f"读取回填.py文件失败: {e}")
        print(f"错误：读取文件失败 - {e}")
        return False
        
    # 替换字体设置
    pattern = r'font_name = "[^"]+"(\s*# 默认字体)'
    replacement = f'font_name = "{new_font}"\\1'
    
    new_content = re.sub(pattern, replacement, content)
    
    if new_content != content:
        try:
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            logger.info(f"字体设置成功更改为: {new_font}")
            print(f"✅ 字体已更改为: {new_font}")
            return True
        except Exception as e:
            logger.error(f"写入文件失败: {e}")
            print(f"❌ 写入文件失败: {e}")
            return False
    else:
        logger.warning("字体设置失败，未找到匹配的字体设置行")
        print("❌ 字体设置失败")
        return False

def main():
    """主函数 - 交互式字体设置"""
    print("=== CAD翻译字体配置工具 ===")
    print()
    
    # 显示当前字体
    current = get_current_font()
    if current:
        print(f"当前字体: {current}")
    else:
        print("无法获取当前字体设置")
    
    print()
    print("可用字体列表:")
    for i, font in enumerate(AVAILABLE_FONTS, 1):
        marker = " ← 当前" if font == current else ""
        print(f"{i:2d}. {font}{marker}")
    
    print()
    try:
        choice = input("请选择字体编号 (1-{}) 或按回车键退出: ".format(len(AVAILABLE_FONTS)))
        
        if not choice.strip():
            print("已取消")
            return
            
        choice_num = int(choice)
        if 1 <= choice_num <= len(AVAILABLE_FONTS):
            selected_font = AVAILABLE_FONTS[choice_num - 1]
            
            if selected_font == current:
                print(f"字体 '{selected_font}' 已经是当前设置")
                return
                
            print(f"正在设置字体为: {selected_font}")
            if set_font(selected_font):
                print("字体设置成功！")
            else:
                print("字体设置失败！")
        else:
            print("无效的选择")
            
    except ValueError:
        print("请输入有效的数字")
    except KeyboardInterrupt:
        print("\n已取消")

if __name__ == "__main__":
    try:
        main()
        logger.info("CAD翻译字体配置工具正常结束")
    except Exception as e:
        logger.error(f"程序运行出现异常: {e}")
        print(f"程序出现错误: {e}")
    except KeyboardInterrupt:
        logger.info("用户中断程序执行")
        print("\n程序被用户中断")
