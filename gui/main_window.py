"""
GUI 界面 - 主窗口
使用 tkinter 实现，兼容 Windows 离线环境
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os
from pathlib import Path
from datetime import datetime

# 添加父目录到路径
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.parser import SLXParser, parse_slx
from core.analyzer import ModelAnalyzer
from core.generator import DocGenerator


class MainWindow:
    """主窗口"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Simulink 文档生成器 v1.0")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # 设置样式
        self._setup_style()
        
        # 创建界面
        self._create_widgets()
        
        # 绑定事件
        self._bind_events()
        
        # 状态变量
        self.model = None
        self.analyzer = None
        self.is_processing = False
        
    def _setup_style(self):
        """设置界面样式"""
        style = ttk.Style()
        style.theme_use('clam')  # 使用 clam 主题，跨平台一致
        
        # 自定义样式
        style.configure('Title.TLabel', font=('Microsoft YaHei', 14, 'bold'))
        style.configure('Info.TLabel', font=('Microsoft YaHei', 10))
        style.configure('Action.TButton', font=('Microsoft YaHei', 10))
        
    def _create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ===== 文件选择区域 =====
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        # SLX 文件
        ttk.Label(file_frame, text="模型文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.slx_path_var = tk.StringVar()
        self.slx_path_entry = ttk.Entry(file_frame, textvariable=self.slx_path_var, width=50)
        self.slx_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(file_frame, text="浏览...", command=self._browse_slx).grid(row=0, column=2, pady=5)
        
        # 输出目录
        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_dir_var = tk.StringVar()
        self.output_dir_entry = ttk.Entry(file_frame, textvariable=self.output_dir_var, width=50)
        self.output_dir_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(file_frame, text="浏览...", command=self._browse_output).grid(row=1, column=2, pady=5)
        
        file_frame.columnconfigure(1, weight=1)
        
        # ===== 文档选项区域 =====
        options_frame = ttk.LabelFrame(main_frame, text="文档选项", padding="10")
        options_frame.pack(fill=tk.X, pady=5)
        
        # 复选框
        self.include_signal_flow_var = tk.BooleanVar(value=True)
        self.include_block_list_var = tk.BooleanVar(value=True)
        self.include_data_dict_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(
            options_frame, 
            text="包含信号流描述", 
            variable=self.include_signal_flow_var
        ).grid(row=0, column=0, sticky=tk.W, pady=2)
        
        ttk.Checkbutton(
            options_frame, 
            text="包含模块清单", 
            variable=self.include_block_list_var
        ).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Checkbutton(
            options_frame, 
            text="包含数据字典", 
            variable=self.include_data_dict_var
        ).grid(row=0, column=2, sticky=tk.W, pady=2)
        
        # ===== 操作按钮区域 =====
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            button_frame, 
            text="解析模型", 
            command=self._parse_model,
            style='Action.TButton'
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="生成文档", 
            command=self._generate_doc,
            style='Action.TButton'
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="预览信息", 
            command=self._preview_info,
            style='Action.TButton'
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="清除日志", 
            command=self._clear_log,
            style='Action.TButton'
        ).pack(side=tk.RIGHT, padx=5)
        
        # ===== 日志/预览区域 =====
        log_frame = ttk.LabelFrame(main_frame, text="日志 / 预览", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # ===== 状态栏 =====
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            status_frame, 
            variable=self.progress_var, 
            maximum=100,
            length=200
        )
        self.progress_bar.pack(side=tk.RIGHT)
        
    def _bind_events(self):
        """绑定事件"""
        # 回车键触发解析
        self.slx_path_entry.bind('<Return>', lambda e: self._parse_model())
        
    def _browse_slx(self):
        """浏览 SLX 文件"""
        file_path = filedialog.askopenfilename(
            title="选择 Simulink 模型文件",
            filetypes=[
                ("Simulink 模型", "*.slx *.mdl"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.slx_path_var.set(file_path)
            # 自动设置输出目录
            output_dir = os.path.dirname(file_path)
            self.output_dir_var.set(output_dir)
            
    def _browse_output(self):
        """浏览输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir_var.set(dir_path)
            
    def _log(self, message: str, level: str = "INFO"):
        """记录日志"""
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 颜色标记
        color_map = {
            "INFO": "black",
            "SUCCESS": "green",
            "WARNING": "orange",
            "ERROR": "red"
        }
        color = color_map.get(level, "black")
        
        # 添加标签
        tag_name = f"{level}_{timestamp}"
        self.log_text.tag_config(tag_name, foreground=color)
        
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", tag_name)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def _clear_log(self):
        """清除日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def _update_status(self, status: str, progress: float = None):
        """更新状态"""
        self.status_var.set(status)
        if progress is not None:
            self.progress_var.set(progress)
            
    def _parse_model(self):
        """解析模型"""
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请稍候...")
            return
            
        slx_path = self.slx_path_var.get()
        if not slx_path:
            messagebox.showerror("错误", "请选择 Simulink 模型文件")
            return
            
        if not os.path.exists(slx_path):
            messagebox.showerror("错误", f"文件不存在: {slx_path}")
            return
            
        # 在线程中执行
        self.is_processing = True
        thread = threading.Thread(target=self._do_parse, args=(slx_path,))
        thread.start()
        
    def _do_parse(self, slx_path: str):
        """执行解析（在后台线程）"""
        try:
            self._update_status("正在解析模型...", 10)
            self._log(f"开始解析: {slx_path}")
            
            # 解析
            parser = SLXParser(slx_path)
            self.model = parser.parse()
            
            self._update_status("正在分析模型...", 50)
            self._log("解析完成，开始分析...")
            
            # 分析
            self.analyzer = ModelAnalyzer(self.model)
            
            self._update_status("就绪", 100)
            self._log("模型分析完成！", "SUCCESS")
            
            # 显示概览
            overview = self.analyzer.get_model_overview()
            self._log(f"模型名称: {overview['name']}")
            self._log(f"模块总数: {overview['total_blocks']}")
            self._log(f"子系统数量: {overview['total_subsystems']}")
            self._log(f"最大层级: {overview['max_depth']}")
            
        except Exception as e:
            self._update_status("解析失败", 0)
            self._log(f"解析错误: {str(e)}", "ERROR")
            messagebox.showerror("错误", f"解析失败:\n{str(e)}")
            
        finally:
            self.is_processing = False
            
    def _generate_doc(self):
        """生成文档"""
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请稍候...")
            return
            
        if self.model is None:
            messagebox.showerror("错误", "请先解析模型")
            return
            
        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showerror("错误", "请选择输出目录")
            return
            
        # 构建输出文件名
        model_name = self.model.name
        output_path = os.path.join(output_dir, f"{model_name}_设计文档.docx")
        
        # 获取选项
        options = {
            'include_signal_flow': self.include_signal_flow_var.get(),
            'include_block_list': self.include_block_list_var.get(),
            'include_data_dict': self.include_data_dict_var.get()
        }
        
        # 在线程中执行
        self.is_processing = True
        thread = threading.Thread(target=self._do_generate, args=(output_path, options))
        thread.start()
        
    def _do_generate(self, output_path: str, options: dict):
        """执行生成（在后台线程）"""
        try:
            self._update_status("正在生成文档...", 20)
            self._log(f"开始生成文档: {output_path}")
            
            # 生成
            generator = DocGenerator(self.model, self.analyzer)
            generator.generate(output_path, options)
            
            self._update_status("完成", 100)
            self._log(f"文档生成成功！", "SUCCESS")
            self._log(f"保存位置: {output_path}")
            
            # 询问是否打开
            if messagebox.askyesno("成功", f"文档已生成:\n{output_path}\n\n是否打开？"):
                os.startfile(output_path)
                
        except Exception as e:
            self._update_status("生成失败", 0)
            self._log(f"生成错误: {str(e)}", "ERROR")
            messagebox.showerror("错误", f"文档生成失败:\n{str(e)}")
            
        finally:
            self.is_processing = False
            
    def _preview_info(self):
        """预览模型信息"""
        if self.model is None:
            messagebox.showerror("错误", "请先解析模型")
            return
            
        self._log("\n" + "=" * 50, "INFO")
        self._log("模型详细信息", "INFO")
        self._log("=" * 50, "INFO")
        
        overview = self.analyzer.get_model_overview()
        
        self._log(f"\n【基本信息】")
        self._log(f"  名称: {overview['name']}")
        self._log(f"  版本: {overview['version'] or 'N/A'}")
        self._log(f"  作者: {overview['author'] or 'N/A'}")
        self._log(f"  创建时间: {self.model.created or 'N/A'}")
        self._log(f"  修改时间: {self.model.modified or 'N/A'}")
        
        self._log(f"\n【统计信息】")
        self._log(f"  模块总数: {overview['total_blocks']}")
        self._log(f"  子系统数量: {overview['total_subsystems']}")
        self._log(f"  最大层级深度: {overview['max_depth']}")
        
        self._log(f"\n【模块类型分布】")
        for stat in overview['block_type_stats'][:10]:
            self._log(f"  {stat.block_type}: {stat.count} 个")
            
        self._log(f"\n【子系统列表】")
        for level, subsystem in self.analyzer.get_all_subsystems():
            indent = "  " * level
            self._log(f"{indent}- {subsystem.name} ({len(subsystem.blocks)} 个模块)")
            
    def run(self):
        """运行应用"""
        self._log("欢迎使用 Simulink 文档生成器！")
        self._log("请选择一个 .slx 或 .mdl 文件开始。")
        self.root.mainloop()


def main():
    """主函数"""
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    main()