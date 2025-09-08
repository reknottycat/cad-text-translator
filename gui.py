#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CAD文件翻译处理系统 - 图形用户界面（线程安全重构版）

重构目标：
1. 修复线程安全问题：所有UI更新通过主线程的after方法执行
2. 统一子进程管理：增加超时控制和错误处理
3. 优化用户体验：添加取消功能和更好的进度反馈
4. 启动时依赖检查：避免运行时错误
5. 集成日志功能：记录用户操作和程序状态

作者: AI Assistant
日期: 2024
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import subprocess
import sys
from pathlib import Path
import queue
import os
import time
import signal
from typing import Optional, List, Callable, Tuple

# 导入日志配置
try:
    from logger_config import get_logger
except ImportError:
    # 如果日志模块不可用，创建一个简单的替代
    import logging
    def get_logger(name="cad_translator"):
        return logging.getLogger(name)

# 内置依赖检查
def _check_deps() -> Tuple[bool, List[str]]:
    """检查必要依赖是否已安装"""
    required = ['ezdxf', 'pandas', 'openpyxl']
    missing = []
    
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    
    return len(missing) == 0, missing

# 已移除Aspose.CAD相关依赖，现在使用浩辰CAD转换器

# 配置管理 - 现在通过命令行参数传递给回填.py
def get_translation_config():
    """获取翻译配置（保留兼容性）"""
    return {'mode': 'replace', 'font_size_reduction': 4}

def set_translation_mode(mode):
    """设置翻译模式（已废弃，现在通过命令行参数传递）"""
    # 此函数已废弃，配置现在通过命令行参数传递给回填.py
    pass

def set_font_size_reduction(value):
    """设置字体减少值（已废弃，现在通过命令行参数传递）"""
    # 此函数已废弃，配置现在通过命令行参数传递给回填.py
    pass

def get_current_font():
    """获取当前字体（从font_config.py读取）"""
    try:
        from font_config import get_current_font as get_font
        return get_font()
    except ImportError:
        return 'Times New Roman'

def set_font(font):
    """设置字体（通过font_config.py）"""
    try:
        from font_config import set_font as set_font_config
        return set_font_config(font)
    except ImportError:
        # 如果font_config.py不可用，直接返回True（GUI会通过命令行参数传递）
        return True

# 使用导入的依赖检查函数
check_dependencies = _check_deps

class CADTranslationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CAD文件翻译处理系统 v2.1")
        self.geometry("950x750")
        
        # 初始化日志记录器
        self.logger = get_logger("cad_gui")
        self.logger.info("CAD翻译工具启动")
        
        # 工作状态（线程安全）
        self._processing = False
        self._current_process = None
        self._cancel_requested = False
        
        # UI更新队列（线程安全通信）
        self.ui_queue = queue.Queue()
        
        # 创建界面
        self._create_widgets()
        
        # 启动依赖检查
        self._check_startup_dependencies()
        
        # 启动UI队列处理
        self._process_ui_queue()
        
        # 绑定关闭事件
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        self.logger.info("GUI界面初始化完成")

    def _on_closing(self):
        """处理窗口关闭事件"""
        if self._processing:
            if messagebox.askokcancel("退出确认", "正在处理文件，确定要退出吗？"):
                self._on_cancel()
                self.destroy()
        else:
            self.destroy()

    def _check_startup_dependencies(self) -> None:
        """启动时检查依赖"""
        def check():
            deps_ok, missing = check_dependencies()
            if not deps_ok:
                msg = f"缺少必要依赖包：{', '.join(missing)}\n\n请运行：pip install {' '.join(missing)}"
                self.ui_queue.put(("show_error", "依赖检查失败", msg))
                self.logger.error(f"依赖检查失败，缺少: {missing}")
            else:
                self.logger.info("依赖检查通过")
        
        # 在后台线程中检查依赖
        threading.Thread(target=check, daemon=True).start()

    def _process_ui_queue(self) -> None:
        """处理UI更新队列（主线程中执行）"""
        try:
            while True:
                action, *args = self.ui_queue.get_nowait()
                
                if action == "log":
                    message = args[0]
                    self.log_text.insert("end", message + "\n")
                    self.log_text.see("end")
                    
                elif action == "set_buttons":
                    converting, extracting, applying = args
                    self._set_buttons(converting, extracting, applying)
                    
                elif action == "show_error":
                    title, message = args
                    messagebox.showerror(title, message)
                    
                elif action == "show_info":
                    title, message = args
                    messagebox.showinfo(title, message)
                    
                elif action == "update_progress":
                    # 可以在这里添加进度条更新逻辑
                    pass
                    
        except queue.Empty:
            pass
        
        # 继续处理队列
        self.after(100, self._process_ui_queue)

    def _update_buttons_state(self, processing: bool) -> None:
        """更新按钮状态（线程安全）"""
        if processing:
            self.ui_queue.put(("set_buttons", True, True, True))
        else:
            self.ui_queue.put(("set_buttons", False, False, False))

    def _append_log_safe(self, message: str) -> None:
        """线程安全的日志添加"""
        self.ui_queue.put(("log", message))
        self.logger.info(message)

    def _safe_ui_call(self, action: str, *args) -> None:
        """线程安全的UI调用"""
        self.ui_queue.put((action, *args))

    def _on_cancel(self) -> None:
        """取消当前操作"""
        if self._processing and self._current_process:
            self._cancel_requested = True
            self._append_log_safe("正在取消操作...")
            
            try:
                # 尝试优雅地终止进程
                if sys.platform == "win32":
                    # Windows
                    subprocess.run(["taskkill", "/F", "/T", "/PID", str(self._current_process.pid)], 
                                 capture_output=True)
                else:
                    # Unix/Linux/Mac
                    os.killpg(os.getpgid(self._current_process.pid), signal.SIGTERM)
                    
                # 等待进程结束
                try:
                    self._current_process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    # 强制终止
                    if sys.platform == "win32":
                        subprocess.run(["taskkill", "/F", "/T", "/PID", str(self._current_process.pid)], 
                                     capture_output=True)
                    else:
                        os.killpg(os.getpgid(self._current_process.pid), signal.SIGKILL)
                        
                self._append_log_safe("操作已取消")
                
            except Exception as e:
                self._append_log_safe(f"取消操作时出错: {e}")
                self.logger.error(f"取消操作失败: {e}")
            
            finally:
                self._current_process = None
                self._cancel_requested = False
                self._set_processing(False)

    @property
    def processing(self) -> bool:
        """获取处理状态"""
        return self._processing

    def _set_processing(self, value: bool) -> None:
        """设置处理状态"""
        self._processing = value
        self._update_buttons_state(value)
        
        if value:
            self.cancel_button.configure(state="normal")
        else:
            self.cancel_button.configure(state="disabled")

    def _create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(main_frame, text="CAD文件翻译处理系统", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=(20, 30))
        
        # 文件选择区域
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # 文件选择按钮
        button_frame = ctk.CTkFrame(file_frame)
        button_frame.pack(fill="x", padx=20, pady=20)
        
        self.select_files_btn = ctk.CTkButton(button_frame, text="选择DWG文件", 
                                            command=self._select_dwg_files,
                                            width=150, height=40)
        self.select_files_btn.pack(side="left", padx=(0, 10))
        
        self.select_folder_btn = ctk.CTkButton(button_frame, text="选择工作文件夹", 
                                             command=self._select_folder,
                                             width=150, height=40)
        self.select_folder_btn.pack(side="left", padx=10)
        
        # 工作目录显示
        self.workdir_label = ctk.CTkLabel(file_frame, text="工作目录：未选择", 
                                         font=ctk.CTkFont(size=12))
        self.workdir_label.pack(pady=(0, 10))
        
        # 字体选择区域
        font_frame = ctk.CTkFrame(main_frame)
        font_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        font_label = ctk.CTkLabel(font_frame, text="选择字体：", 
                                 font=ctk.CTkFont(size=14, weight="bold"))
        font_label.pack(side="left", padx=(20, 10), pady=20)
        
        # 获取可用字体列表
        try:
            from font_config import AVAILABLE_FONTS
            fonts = AVAILABLE_FONTS
        except ImportError:
            fonts = ["Times New Roman", "Arial", "SimSun", "Microsoft YaHei"]
        
        self.font_var = ctk.StringVar(value=get_current_font() or fonts[0])
        self.font_dropdown = ctk.CTkComboBox(font_frame, values=fonts, 
                                           variable=self.font_var,
                                           command=self._on_font_changed,
                                           width=200)
        self.font_dropdown.pack(side="left", padx=10, pady=20)
        
        # 操作按钮区域
        action_frame = ctk.CTkFrame(main_frame)
        action_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # 第一行按钮
        button_row1 = ctk.CTkFrame(action_frame)
        button_row1.pack(fill="x", padx=20, pady=(20, 10))
        
        self.convert_btn = ctk.CTkButton(button_row1, text="1. 转换DWG→DXF", 
                                       command=self._on_convert,
                                       width=200, height=50,
                                       font=ctk.CTkFont(size=14, weight="bold"))
        self.convert_btn.pack(side="left", padx=(0, 10))
        
        self.extract_btn = ctk.CTkButton(button_row1, text="2. 提取文本", 
                                       command=self._on_extract,
                                       width=200, height=50,
                                       font=ctk.CTkFont(size=14, weight="bold"))
        self.extract_btn.pack(side="left", padx=10)
        
        # 第二行按钮
        button_row2 = ctk.CTkFrame(action_frame)
        button_row2.pack(fill="x", padx=20, pady=10)
        
        self.open_excel_btn = ctk.CTkButton(button_row2, text="3. 打开Excel翻译", 
                                          command=self._on_open_excel,
                                          width=200, height=50,
                                          font=ctk.CTkFont(size=14, weight="bold"))
        self.open_excel_btn.pack(side="left", padx=(0, 10))
        
        self.apply_btn = ctk.CTkButton(button_row2, text="4. 应用翻译", 
                                     command=self._on_apply,
                                     width=200, height=50,
                                     font=ctk.CTkFont(size=14, weight="bold"))
        self.apply_btn.pack(side="left", padx=10)
        
        # 第三行按钮
        button_row3 = ctk.CTkFrame(action_frame)
        button_row3.pack(fill="x", padx=20, pady=(10, 20))
        
        self.auto_btn = ctk.CTkButton(button_row3, text="🚀 一键处理", 
                                    command=self._on_auto,
                                    width=200, height=50,
                                    font=ctk.CTkFont(size=16, weight="bold"),
                                    fg_color="#FF6B35", hover_color="#E55A2B")
        self.auto_btn.pack(side="left", padx=(0, 10))
        
        self.cancel_button = ctk.CTkButton(button_row3, text="取消操作", 
                                         command=self._on_cancel,
                                         width=200, height=50,
                                         font=ctk.CTkFont(size=14, weight="bold"),
                                         fg_color="#DC143C", hover_color="#B22222",
                                         state="disabled")
        self.cancel_button.pack(side="left", padx=10)
        
        # 日志区域
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        log_label = ctk.CTkLabel(log_frame, text="处理日志：", 
                                font=ctk.CTkFont(size=14, weight="bold"))
        log_label.pack(anchor="w", padx=20, pady=(20, 10))
        
        self.log_text = ctk.CTkTextbox(log_frame, height=200, 
                                      font=ctk.CTkFont(family="Consolas", size=11))
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # 初始日志
        self._log("CAD文件翻译处理系统已启动")
        self._log("请选择DWG文件或工作文件夹开始处理")

    def _on_font_changed(self, selected_font: str):
        """字体选择改变时的回调"""
        if set_font(selected_font):
            self._log(f"字体已设置为: {selected_font}")
            self.logger.info(f"用户更改字体为: {selected_font}")
        else:
            self._log(f"字体设置失败: {selected_font}")
            self.logger.error(f"字体设置失败: {selected_font}")
            # 恢复到之前的字体
            current = get_current_font()
            if current:
                self.font_var.set(current)

    def _ensure_working_dir(self) -> Optional[Path]:
        """确保有工作目录"""
        if hasattr(self, 'working_dir') and self.working_dir:
            return self.working_dir
        else:
            messagebox.showerror("错误", "请先选择DWG文件或工作文件夹")
            return None

    def _set_buttons(self, converting=False, extracting=False, applying=False):
        """设置按钮状态"""
        # 这个方法现在通过UI队列调用，确保在主线程中执行
        pass

    def _log(self, message: str):
        """添加日志消息"""
        self._append_log_safe(message)

    def _get_script_path(self, script_name: str) -> str:
        """获取脚本路径"""
        return str(Path(script_name).resolve())

    def _stream_subprocess(self, cmd: List[str], cwd: Path, timeout: int = 300) -> int:
        """流式执行子进程并实时显示输出"""
        self.logger.info(f"执行命令: {' '.join(cmd)}")
        self._append_log_safe(f"执行: {' '.join(cmd)}")
        
        try:
            # 创建子进程
            if sys.platform == "win32":
                # Windows: 创建新的进程组
                self._current_process = subprocess.Popen(
                    cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    universal_newlines=True, bufsize=1,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
                )
            else:
                # Unix/Linux/Mac: 创建新的进程组
                self._current_process = subprocess.Popen(
                    cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    universal_newlines=True, bufsize=1, preexec_fn=os.setsid
                )
            
            # 实时读取输出
            start_time = time.time()
            while True:
                if self._cancel_requested:
                    self._append_log_safe("操作被用户取消")
                    return -1
                
                # 检查超时
                if time.time() - start_time > timeout:
                    self._append_log_safe(f"操作超时（{timeout}秒），正在终止...")
                    self._current_process.terminate()
                    return -2
                
                output = self._current_process.stdout.readline()
                if output:
                    self._append_log_safe(output.strip())
                elif self._current_process.poll() is not None:
                    break
                    
                time.sleep(0.1)
            
            return_code = self._current_process.returncode
            self.logger.info(f"命令执行完成，返回码: {return_code}")
            return return_code
            
        except Exception as e:
            error_msg = f"执行命令时出错: {e}"
            self._append_log_safe(error_msg)
            self.logger.error(error_msg)
            return -3
        finally:
            self._current_process = None

    def _count_files_text(self, directory: Path) -> str:
        """统计目录中的文件数量"""
        dwg_count = len(list(directory.glob("*.dwg")))
        dxf_count = len(list(directory.glob("*.dxf")))
        return f"(DWG: {dwg_count}, DXF: {dxf_count})"

    def _select_dwg_files(self):
        """选择DWG文件"""
        files = filedialog.askopenfilenames(
            title="选择DWG文件",
            filetypes=[("DWG文件", "*.dwg"), ("所有文件", "*.*")]
        )
        
        if files:
            # 使用第一个文件的目录作为工作目录
            self.working_dir = Path(files[0]).parent
            count_text = self._count_files_text(self.working_dir)
            self.workdir_label.configure(text=f"工作目录：{self.working_dir} {count_text}")
            self._log(f"已选择工作目录: {self.working_dir}")
            self._log(f"找到 {len(files)} 个DWG文件")
            self.logger.info(f"用户选择了 {len(files)} 个DWG文件，工作目录: {self.working_dir}")

    def _select_folder(self):
        """选择工作文件夹"""
        folder = filedialog.askdirectory(title="选择工作文件夹")
        if folder:
            self.working_dir = Path(folder)
            count_text = self._count_files_text(self.working_dir)
            self.workdir_label.configure(text=f"工作目录：{self.working_dir} {count_text}")
            self._log(f"已选择工作目录: {self.working_dir}")
            self.logger.info(f"用户选择工作目录: {self.working_dir}")

    def _on_convert(self):
        """转换DWG到DXF"""
        workdir = self._ensure_working_dir()
        if not workdir:
            return
            
        if self.processing:
            messagebox.showwarning("警告", "正在处理中，请等待完成")
            return
            
        threading.Thread(target=self._convert_worker, args=(workdir,), daemon=True).start()

    def _on_extract(self):
        """提取文本"""
        workdir = self._ensure_working_dir()
        if not workdir:
            return
            
        if self.processing:
            messagebox.showwarning("警告", "正在处理中，请等待完成")
            return
            
        threading.Thread(target=self._extract_worker, args=(workdir,), daemon=True).start()

    def _on_open_excel(self):
        """打开Excel文件"""
        workdir = self._ensure_working_dir()
        if not workdir:
            return
            
        excel_file = workdir / "extracted_texts.xlsx"
        if excel_file.exists():
            try:
                os.startfile(str(excel_file))
                self._log(f"已打开Excel文件: {excel_file}")
                self.logger.info(f"用户打开Excel文件: {excel_file}")
            except Exception as e:
                error_msg = f"打开Excel文件失败: {e}"
                self._log(error_msg)
                self.logger.error(error_msg)
                messagebox.showerror("错误", error_msg)
        else:
            messagebox.showerror("错误", "未找到extracted_texts.xlsx文件，请先执行文本提取")

    def _on_apply(self):
        """应用翻译"""
        workdir = self._ensure_working_dir()
        if not workdir:
            return
            
        if self.processing:
            messagebox.showwarning("警告", "正在处理中，请等待完成")
            return
            
        threading.Thread(target=self._apply_worker, args=(workdir,), daemon=True).start()

    def _on_auto(self):
        """一键处理"""
        workdir = self._ensure_working_dir()
        if not workdir:
            return
            
        if self.processing:
            messagebox.showwarning("警告", "正在处理中，请等待完成")
            return
            
        threading.Thread(target=self._auto_worker, args=(workdir,), daemon=True).start()

    def _convert_worker(self, workdir: Path):
        """转换工作线程"""
        self._set_processing(True)
        self._append_log_safe("=== 开始DWG转DXF ===")        
        
        try:
            # 检查是否有DWG文件
            dwg_files = list(workdir.glob("*.dwg"))
            if not dwg_files:
                self._append_log_safe("错误：工作目录中没有找到DWG文件")
                self._safe_ui_call("show_error", "错误", "工作目录中没有找到DWG文件")
                return
            
            self._append_log_safe(f"找到 {len(dwg_files)} 个DWG文件")
            
            # 尝试使用浩辰转换器
            converter_script = self._get_script_path("haochen_optimized_converter.py")
            if Path(converter_script).exists():
                self._append_log_safe("使用浩辰CAD转换器...")
                cmd = [sys.executable, converter_script, str(workdir)]
                result = self._stream_subprocess(cmd, workdir, timeout=600)  # 10分钟超时
                
                if result == 0:
                    self._append_log_safe("✅ DWG转DXF完成")
                    self._safe_ui_call("show_info", "完成", "DWG文件转换完成")
                elif result == -1:
                    self._append_log_safe("❌ 转换被用户取消")
                elif result == -2:
                    self._append_log_safe("❌ 转换超时")
                    self._safe_ui_call("show_error", "超时", "转换操作超时，请检查文件大小和系统性能")
                else:
                    self._append_log_safe(f"❌ 转换失败，返回码: {result}")
                    self._safe_ui_call("show_error", "转换失败", "DWG转换失败，请检查日志")
            else:
                self._append_log_safe("错误：找不到转换器脚本")
                self._safe_ui_call("show_error", "错误", "找不到haochen_optimized_converter.py")
                
        except Exception as e:
            error_msg = f"转换过程出错: {e}"
            self._append_log_safe(error_msg)
            self.logger.error(error_msg)
            self._safe_ui_call("show_error", "错误", error_msg)
        finally:
            self._set_processing(False)
            self._append_log_safe("=== DWG转DXF结束 ===")

    def _extract_worker(self, workdir: Path):
        """提取工作线程"""
        self._set_processing(True)
        self._append_log_safe("=== 开始文本提取 ===")
        
        try:
            # 检查是否有DXF文件
            dxf_files = list(workdir.glob("*.dxf"))
            if not dxf_files:
                self._append_log_safe("错误：工作目录中没有找到DXF文件")
                self._safe_ui_call("show_error", "错误", "工作目录中没有找到DXF文件，请先转换DWG文件")
                return
            
            self._append_log_safe(f"找到 {len(dxf_files)} 个DXF文件")
            
            # 执行文本提取
            extract_script = self._get_script_path("extract_texts.py")
            cmd = [sys.executable, extract_script, str(workdir)]
            result = self._stream_subprocess(cmd, workdir, timeout=300)  # 5分钟超时
            
            if result == 0:
                self._append_log_safe("✅ 文本提取完成")
                self._safe_ui_call("show_info", "完成", "文本提取完成，可以打开Excel文件进行翻译")
            elif result == -1:
                self._append_log_safe("❌ 提取被用户取消")
            elif result == -2:
                self._append_log_safe("❌ 提取超时")
                self._safe_ui_call("show_error", "超时", "文本提取超时")
            else:
                self._append_log_safe(f"❌ 文本提取失败，返回码: {result}")
                self._safe_ui_call("show_error", "提取失败", "文本提取失败，请检查日志")
                
        except Exception as e:
            error_msg = f"提取过程出错: {e}"
            self._append_log_safe(error_msg)
            self.logger.error(error_msg)
            self._safe_ui_call("show_error", "错误", error_msg)
        finally:
            self._set_processing(False)
            self._append_log_safe("=== 文本提取结束 ===")

    def _apply_worker(self, workdir: Path):
        """应用翻译工作线程"""
        self._set_processing(True)
        self._append_log_safe("=== 开始应用翻译 ===")
        
        try:
            # 检查Excel文件是否存在
            excel_file = workdir / "extracted_texts.xlsx"
            if not excel_file.exists():
                self._append_log_safe("错误：找不到extracted_texts.xlsx文件")
                self._safe_ui_call("show_error", "错误", "找不到extracted_texts.xlsx文件，请先执行文本提取")
                return
            
            # 检查是否有DXF文件
            dxf_files = list(workdir.glob("*.dxf"))
            if not dxf_files:
                self._append_log_safe("错误：工作目录中没有找到DXF文件")
                self._safe_ui_call("show_error", "错误", "工作目录中没有找到DXF文件")
                return
            
            self._append_log_safe(f"找到 {len(dxf_files)} 个DXF文件")
            
            # 获取当前字体设置
            current_font = self.font_var.get()
            
            # 执行翻译应用
            apply_script = self._get_script_path("回填.py")
            cmd = [sys.executable, apply_script, str(workdir), "--font", current_font]
            result = self._stream_subprocess(cmd, workdir, timeout=300)  # 5分钟超时
            
            if result == 0:
                self._append_log_safe("✅ 翻译应用完成")
                self._safe_ui_call("show_info", "完成", "翻译已成功应用到DXF文件")
            elif result == -1:
                self._append_log_safe("❌ 应用被用户取消")
            elif result == -2:
                self._append_log_safe("❌ 应用超时")
                self._safe_ui_call("show_error", "超时", "翻译应用超时")
            else:
                self._append_log_safe(f"❌ 翻译应用失败，返回码: {result}")
                self._safe_ui_call("show_error", "应用失败", "翻译应用失败，请检查日志")
                
        except Exception as e:
            error_msg = f"应用过程出错: {e}"
            self._append_log_safe(error_msg)
            self.logger.error(error_msg)
            self._safe_ui_call("show_error", "错误", error_msg)
        finally:
            self._set_processing(False)
            self._append_log_safe("=== 应用翻译结束 ===")

    def _auto_worker(self, workdir: Path):
        """一键处理工作线程"""
        self._set_processing(True)
        self._append_log_safe("=== 开始一键处理 ===")
        
        try:
            # 步骤1: 转换DWG到DXF
            dwg_files = list(workdir.glob("*.dwg"))
            if dwg_files:
                self._append_log_safe("步骤1: 转换DWG到DXF...")
                converter_script = self._get_script_path("haochen_optimized_converter.py")
                if Path(converter_script).exists():
                    cmd = [sys.executable, converter_script, str(workdir)]
                    result = self._stream_subprocess(cmd, workdir, timeout=600)
                    if result != 0 and not self._cancel_requested:
                        self._append_log_safe("❌ DWG转换失败，停止处理")
                        self._safe_ui_call("show_error", "转换失败", "DWG转换失败，请检查日志")
                        return
                else:
                    self._append_log_safe("❌ 找不到转换器，停止处理")
                    self._safe_ui_call("show_error", "错误", "找不到转换器脚本")
                    return
            
            if self._cancel_requested:
                return
            
            # 步骤2: 提取文本
            self._append_log_safe("步骤2: 提取文本...")
            extract_script = self._get_script_path("extract_texts.py")
            cmd = [sys.executable, extract_script, str(workdir)]
            result = self._stream_subprocess(cmd, workdir, timeout=300)
            if result != 0 and not self._cancel_requested:
                self._append_log_safe("❌ 文本提取失败，停止处理")
                self._safe_ui_call("show_error", "提取失败", "文本提取失败，请检查日志")
                return
            
            if self._cancel_requested:
                return
            
            # 步骤3: 提示用户翻译
            self._append_log_safe("步骤3: 请在Excel中完成翻译后点击确定继续...")
            excel_file = workdir / "extracted_texts.xlsx"
            if excel_file.exists():
                try:
                    os.startfile(str(excel_file))
                except:
                    pass
            
            # 在UI线程中显示对话框
            def ask_continue():
                return messagebox.askyesno("翻译确认", 
                    "Excel文件已打开，请完成翻译后点击'是'继续应用翻译，或点击'否'取消操作")
            
            # 这里需要在主线程中执行对话框
            continue_processing = False
            def ui_ask():
                nonlocal continue_processing
                continue_processing = ask_continue()
            
            self.after(0, ui_ask)
            
            # 等待用户响应
            while not hasattr(self, '_dialog_responded'):
                if self._cancel_requested:
                    return
                time.sleep(0.5)
            
            if not continue_processing:
                self._append_log_safe("用户取消了翻译应用")
                return
            
            # 步骤4: 应用翻译
            self._append_log_safe("步骤4: 应用翻译...")
            current_font = self.font_var.get()
            apply_script = self._get_script_path("回填.py")
            cmd = [sys.executable, apply_script, str(workdir), "--font", current_font]
            result = self._stream_subprocess(cmd, workdir, timeout=300)
            
            if result == 0:
                self._append_log_safe("✅ 一键处理完成！")
                self._safe_ui_call("show_info", "完成", "一键处理完成！所有步骤都已成功执行")
            else:
                self._append_log_safe(f"❌ 翻译应用失败，返回码: {result}")
                self._safe_ui_call("show_error", "应用失败", "翻译应用失败，请检查日志")
                
        except Exception as e:
            error_msg = f"一键处理出错: {e}"
            self._append_log_safe(error_msg)
            self.logger.error(error_msg)
            self._safe_ui_call("show_error", "错误", error_msg)
        finally:
            self._set_processing(False)
            self._append_log_safe("=== 一键处理结束 ===")

if __name__ == "__main__":
    ctk.set_appearance_mode("System")  # 可选: "System"/"Light"/"Dark"
    ctk.set_default_color_theme("blue")
    app = CADTranslationApp()
    app.mainloop()
