import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from docx import Document
import traceback
from collections import defaultdict

from cleaner import StyleCategory, clean_document

# 清理结果的中性类别 → 中文展示名；展示词汇只存在于 GUI adapter
CATEGORY_LABELS = {
    StyleCategory.PARAGRAPH: '段落',
    StyleCategory.CHARACTER: '字符',
    StyleCategory.TABLE: '表格',
    StyleCategory.OTHER: '其他',
}

class WordStyleCleaner:
    def __init__(self, root):
        self.root = root
        self.root.title("Word样式清理工具")
        self.root.geometry("700x350")  # 增加窗口高度以容纳样式列表区域
        
        # 创建UI组件
        self.create_widgets()
        
        # 存储删除的样式信息
        self.deleted_styles_info = defaultdict(list)
    
    def create_widgets(self):
        # 创建选择文件或文件夹的框架
        selection_frame = tk.Frame(self.root)
        selection_frame.pack(pady=10, fill=tk.X, padx=10)
        
        file_path_label = tk.Label(selection_frame, text='选择文件或文件夹：')
        file_path_label.pack(side=tk.LEFT)
        
        self.file_path_var = tk.StringVar()
        file_path_entry = tk.Entry(selection_frame, textvariable=self.file_path_var)
        file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        choose_file_button = tk.Button(selection_frame, text='选择文件', command=self.choose_file)
        choose_file_button.pack(side=tk.LEFT, padx=2)
        
        choose_folder_button = tk.Button(selection_frame, text='选择文件夹', command=self.choose_folder)
        choose_folder_button.pack(side=tk.LEFT, padx=2)
        
        # 创建处理按钮和进度条
        action_frame = tk.Frame(self.root)
        action_frame.pack(pady=10, fill=tk.X, padx=10)
        
        self.remove_button = tk.Button(action_frame, text='删除未使用的样式', command=self.remove_unused_styles)
        self.remove_button.pack(side=tk.LEFT)
        
        # 创建进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(action_frame, variable=self.progress_var, length=300)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        # 创建统计信息框架
        stats_frame = tk.Frame(self.root)
        stats_frame.pack(pady=5, fill=tk.X, padx=10)
        
        self.stats_label = tk.Label(stats_frame, text="统计信息：共处理 0 个文件，删除 0 个样式")
        self.stats_label.pack()
        
        # 创建样式列表区域
        style_list_frame = tk.LabelFrame(self.root, text="已删除的样式详情")
        style_list_frame.pack(pady=5, fill=tk.BOTH, expand=True, padx=10)
        
        # 创建文本显示区域（移除标签页）
        self.result_text = scrolledtext.ScrolledText(style_list_frame, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_text.config(state=tk.DISABLED)
        
        # 创建状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_label = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[('Word 文档', '*.docx')])
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set(f"已选择文件: {os.path.basename(file_path)}")
            # 清空之前的删除样式信息
            self.deleted_styles_info.clear()
            self._update_style_list_display()
    
    def choose_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.file_path_var.set(folder_path)
            # 统计文件夹中的docx文件数量
            docx_count = len([f for f in os.listdir(folder_path) if f.endswith('.docx')])
            self.status_var.set(f"已选择文件夹，包含 {docx_count} 个Word文档")
            # 清空之前的删除样式信息
            self.deleted_styles_info.clear()
            self._update_style_list_display()
    
    def remove_unused_styles(self):
        # 获取选择的文件或文件夹路径
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件或文件夹")
            return
        
        try:
            # 禁用按钮防止重复点击
            self.remove_button.config(state=tk.DISABLED)
            self.status_var.set("正在处理...")
            self.root.update()
            
            # 清空之前的删除样式信息
            self.deleted_styles_info.clear()
            
            # 判断是单个文件还是文件夹
            if os.path.isfile(file_path):
                # 单个文件处理
                self._process_single_file(file_path)
            elif os.path.isdir(file_path):
                # 文件夹处理
                self._process_folder(file_path)
            else:
                messagebox.showerror("错误", "无效的文件或文件夹路径")
                return
            
            # 更新样式列表显示
            self._update_style_list_display()
            
            messagebox.showinfo("完成", "样式清理完成！")
            self.status_var.set("处理完成")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误：{str(e)}")
            self.status_var.set(f"处理失败: {str(e)}")
            traceback.print_exc()
        finally:
            # 恢复按钮状态
            self.remove_button.config(state=tk.NORMAL)
            self.progress_var.set(0)
    
    def _process_single_file(self, input_file_path):
        """处理单个Word文件"""
        # 自动在原文件夹保存，添加后缀
        base_name = os.path.splitext(os.path.basename(input_file_path))[0]
        output_dir = os.path.dirname(input_file_path)
        save_file_path = os.path.join(output_dir, f"{base_name}_Q.docx")
        
        # 如果文件已存在，添加数字后缀
        counter = 1
        while os.path.exists(save_file_path):
            save_file_path = os.path.join(output_dir, f"{base_name}_Q({counter}).docx")
            counter += 1
        
        self.status_var.set(f"正在处理: {os.path.basename(input_file_path)}")
        self.root.update()
        
        # 清理样式并保存，同时获取删除的样式信息
        deleted_styles = self._remove_unused_styles_from_file(input_file_path, save_file_path)
        
        # 存储删除的样式信息
        file_name = os.path.basename(input_file_path)
        self.deleted_styles_info[file_name] = deleted_styles
        
        self.status_var.set(f"已保存: {os.path.basename(save_file_path)}")
    
    def _process_folder(self, folder_path):
        """处理文件夹中的所有Word文件"""
        # 获取所有docx文件
        docx_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]
        if not docx_files:
            messagebox.showinfo("提示", "所选文件夹中没有找到Word文档(.docx)")
            return
        
        # 直接使用原文件夹，添加后缀（不再询问保存位置）
        output_folder = folder_path
        
        # 处理每个文件
        for i, file_name in enumerate(docx_files):
            try:
                input_file_path = os.path.join(folder_path, file_name)
                
                # 在原文件夹中，添加后缀
                base_name = os.path.splitext(file_name)[0]
                output_file_path = os.path.join(output_folder, f"{base_name}_Q.docx")
                
                # 如果文件已存在，添加数字后缀
                counter = 1
                while os.path.exists(output_file_path):
                    output_file_path = os.path.join(output_folder, f"{base_name}_Q({counter}).docx")
                    counter += 1
                
                # 更新状态
                self.status_var.set(f"正在处理 ({i+1}/{len(docx_files)}): {file_name}")
                self.progress_var.set((i+1) / len(docx_files) * 100)
                self.root.update()
                
                # 清理样式并保存，同时获取删除的样式信息
                deleted_styles = self._remove_unused_styles_from_file(input_file_path, output_file_path)
                
                # 存储删除的样式信息
                self.deleted_styles_info[file_name] = deleted_styles
                
            except Exception as e:
                messagebox.showwarning("警告", f"处理文件 {file_name} 时出错：{str(e)}")
                continue
    
    def _remove_unused_styles_from_file(self, input_file_path, output_file_path):
        """从单个Word文件中删除未使用的样式"""
        # 加载 Word 文档
        document = Document(input_file_path)

        # 清理未使用的自定义样式（核心逻辑在 cleaner module）
        deleted = clean_document(document)

        # 保存新的 Word 文档
        document.save(output_file_path)

        return [
            {'name': d.name, 'type': CATEGORY_LABELS[d.category]}
            for d in deleted
        ]
    
    def _update_style_list_display(self):
        """更新已删除样式的显示列表"""
        # 计算统计信息
        total_files = len(self.deleted_styles_info)
        total_styles = sum(len(styles) for styles in self.deleted_styles_info.values())
        
        # 更新统计标签
        self.stats_label.config(text=f"统计信息：共处理 {total_files} 个文件，删除 {total_styles} 个样式")
        
        # 更新结果显示区域
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        
        if not self.deleted_styles_info:
            self.result_text.insert(tk.END, "暂无删除的样式信息")
        else:
            for file_name, deleted_styles in self.deleted_styles_info.items():
                if deleted_styles:
                    self.result_text.insert(tk.END, f"📄 文件: {file_name}\n")
                    self.result_text.insert(tk.END, f"   删除样式数量: {len(deleted_styles)}\n")
                    self.result_text.insert(tk.END, "   删除的样式：\n")
                    
                    # 按样式类型分类
                    styles_by_type = {"段落": [], "字符": [], "表格": [], "其他": [], "未知": []}
                    for style_info in deleted_styles:
                        styles_by_type[style_info['type']].append(style_info['name'])
                    
                    # 显示分类后的样式
                    for style_type in ["段落", "字符", "表格", "其他", "未知"]:
                        if styles_by_type[style_type]:
                            self.result_text.insert(tk.END, f"     {style_type}样式：\n")
                            for style_name in sorted(styles_by_type[style_type]):
                                self.result_text.insert(tk.END, f"       • {style_name}\n")
                    
                    self.result_text.insert(tk.END, "\n")
                else:
                    self.result_text.insert(tk.END, f"📄 文件: {file_name}\n")
                    self.result_text.insert(tk.END, "   未删除任何样式\n\n")
        
        self.result_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = WordStyleCleaner(root)
    root.mainloop()
