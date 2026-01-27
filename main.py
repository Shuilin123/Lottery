import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import json
import os
import shutil
from pathlib import Path
import sys
import webbrowser
import threading
# 假设flask_server目录下有app.py文件
from flask_server.webapp import WebApp  


class StudentManager:
    def __init__(self, parent, data_path):
        self.parent = parent
        self.data_path = data_path
        self.excel_path = os.path.join(data_path, "users.xlsx")
        self.setup_ui()
        self.load_students()
    
    def setup_ui(self):
        # 学生信息管理框架
        student_frame = ttk.LabelFrame(self.parent, text="学生信息管理", padding=10)
        student_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 说明标签
        ttk.Label(student_frame, text="第一步：配置学生信息", 
                 font=("Microsoft YaHei", 10, "bold"), foreground="blue").pack(anchor=tk.W, pady=(0, 10))
        
        # 按钮框架
        btn_frame = ttk.Frame(student_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(btn_frame, text="📤 上传Excel文件", 
                  command=self.upload_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="➕ 添加学生", 
                  command=self.add_student).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✏️ 编辑选中", 
                  command=self.edit_student).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除选中", 
                  command=self.delete_student).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 保存修改", 
                  command=self.save_students).pack(side=tk.LEFT, padx=5)
        
        # 表格框架
        table_frame = ttk.Frame(student_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格
        self.tree = ttk.Treeview(table_frame, columns=("ID", "姓名", "学号", "班级"), 
                                 show="headings", height=15)
        
        # 设置列标题
        self.tree.heading("ID", text="序号")
        self.tree.heading("姓名", text="姓名")
        self.tree.heading("学号", text="学号")
        self.tree.heading("班级", text="班级")
        
        # 设置列宽
        self.tree.column("ID", width=50, anchor=tk.CENTER)
        self.tree.column("姓名", width=100, anchor=tk.CENTER)
        self.tree.column("学号", width=120, anchor=tk.CENTER)
        self.tree.column("班级", width=100, anchor=tk.CENTER)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定双击编辑事件
        self.tree.bind("<Double-1>", lambda e: self.edit_student())
    
    def upload_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            try:
                # 读取Excel文件
                df = pd.read_excel(file_path)
                required_columns = ["姓名", "学号", "班级"]
                
                # 检查必要列是否存在
                missing_cols = [col for col in required_columns if col not in df.columns]
                if missing_cols:
                    messagebox.showerror("错误", f"Excel文件中缺少必要列: {missing_cols}")
                    return
                
                # 保存到指定位置
                df.to_excel(self.excel_path, index=False)
                self.load_students()
                messagebox.showinfo("成功", f"已上传 {len(df)} 条学生记录")
            except Exception as e:
                messagebox.showerror("错误", f"上传失败: {str(e)}")
    
    def load_students(self):
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 加载Excel数据
        if os.path.exists(self.excel_path):
            try:
                df = pd.read_excel(self.excel_path)
                for idx, row in df.iterrows():
                    self.tree.insert("", tk.END, values=(
                        idx + 1,
                        str(row.get("姓名", "")),
                        str(row.get("学号", "")),
                        str(row.get("班级", ""))
                    ))
            except Exception as e:
                print(f"加载学生数据失败: {e}")
    
    def add_student(self):
        # 创建添加对话框
        dialog = tk.Toplevel(self.parent)
        dialog.title("添加学生")
        dialog.geometry("350x250")
        dialog.resizable(False, False)
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # 居中对话框
        dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 输入框框架
        input_frame = ttk.Frame(dialog, padding=20)
        input_frame.pack(fill=tk.BOTH, expand=True)
        
        # 姓名输入
        ttk.Label(input_frame, text="姓名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_entry = ttk.Entry(input_frame, width=25)
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 学号输入
        ttk.Label(input_frame, text="学号:").grid(row=1, column=0, sticky=tk.W, pady=5)
        id_entry = ttk.Entry(input_frame, width=25)
        id_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 班级输入
        ttk.Label(input_frame, text="班级:").grid(row=2, column=0, sticky=tk.W, pady=5)
        class_entry = ttk.Entry(input_frame, width=25)
        class_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 按钮框架
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        def save():
            name = name_entry.get().strip()
            student_id = id_entry.get().strip()
            class_name = class_entry.get().strip()
            
            if not all([name, student_id, class_name]):
                messagebox.showwarning("警告", "请填写所有字段")
                return
            
            # 添加到表格
            new_id = len(self.tree.get_children()) + 1
            self.tree.insert("", tk.END, values=(new_id, name, student_id, class_name))
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
        
        name_entry.focus_set()
    
    def delete_student(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的学生")
            return
        
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selected)} 名学生吗？"):
            for item in selected:
                self.tree.delete(item)
    
    def edit_student(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要编辑的学生")
            return
        
        item = selected[0]
        values = self.tree.item(item)["values"]
        
        # 创建编辑对话框
        dialog = tk.Toplevel(self.parent)
        dialog.title("编辑学生")
        dialog.geometry("350x250")
        dialog.resizable(False, False)
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # 居中对话框
        dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 输入框框架
        input_frame = ttk.Frame(dialog, padding=20)
        input_frame.pack(fill=tk.BOTH, expand=True)
        
        # 姓名输入
        ttk.Label(input_frame, text="姓名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_entry = ttk.Entry(input_frame, width=25)
        name_entry.insert(0, values[1])
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 学号输入
        ttk.Label(input_frame, text="学号:").grid(row=1, column=0, sticky=tk.W, pady=5)
        id_entry = ttk.Entry(input_frame, width=25)
        id_entry.insert(0, values[2])
        id_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 班级输入
        ttk.Label(input_frame, text="班级:").grid(row=2, column=0, sticky=tk.W, pady=5)
        class_entry = ttk.Entry(input_frame, width=25)
        class_entry.insert(0, values[3])
        class_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 按钮框架
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        def save():
            self.tree.item(item, values=(
                values[0],
                name_entry.get().strip(),
                id_entry.get().strip(),
                class_entry.get().strip()
            ))
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
        
        name_entry.focus_set()
        name_entry.select_range(0, tk.END)
    
    def save_students(self):
        # 从表格中获取数据并保存到Excel
        try:
            data = []
            for item in self.tree.get_children():
                values = self.tree.item(item)["values"]
                data.append({
                    "姓名": values[1],
                    "学号": values[2],
                    "班级": values[3]
                })
            
            if not data:
                messagebox.showwarning("警告", "没有学生数据可以保存")
                return
            
            df = pd.DataFrame(data)
            df.to_excel(self.excel_path, index=False)
            messagebox.showinfo("成功", f"已保存 {len(data)} 条学生信息")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")


class PrizeManager:
    def __init__(self, parent, data_path, img_path):
        self.parent = parent
        self.data_path = data_path
        self.img_path = img_path
        self.json_path = os.path.join(data_path, "prizes.json")
        self.prizes = []
        # 新增：维护每次抽取数量的数组
        self.prizes_count = []
        self.current_image_path = None  # 用于临时存储图片路径
        self.setup_ui()
        self.load_prizes()
    
    def setup_ui(self):
        # 奖项管理框架
        prize_frame = ttk.LabelFrame(self.parent, text="奖项配置管理", padding=10)
        prize_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 说明标签
        ttk.Label(prize_frame, text="第二步：配置奖项信息", 
                 font=("Microsoft YaHei", 10, "bold"), foreground="blue").pack(anchor=tk.W, pady=(0, 10))
        
        # 按钮框架
        btn_frame = ttk.Frame(prize_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(btn_frame, text="➕ 添加奖项", 
                  command=self.add_prize).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✏️ 编辑奖项", 
                  command=self.edit_prize).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ 删除奖项", 
                  command=self.delete_prize).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 刷新列表", 
                  command=self.load_prizes).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 保存配置", 
                  command=self.save_prizes).pack(side=tk.LEFT, padx=5)
        
        # 表格框架
        table_frame = ttk.Frame(prize_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格：新增「每次抽取数量」列
        columns = ("等级", "奖项名称", "奖品名称", "总数量", "每次抽取数量", "图片")
        self.tree = ttk.Treeview(table_frame, columns=columns, 
                                 show="headings", height=10)
        
        # 设置列标题和宽度
        column_widths = {
            "等级": 60,
            "奖项名称": 100,
            "奖品名称": 150,
            "总数量": 80,
            "每次抽取数量": 100,  # 新增列宽度
            "图片": 150
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], anchor=tk.CENTER)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def load_prizes(self):
        """加载奖项数据（修复图片名称和每次抽取数量列错位问题）"""
        try:
            if os.path.exists(self.json_path):
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    # 读取JSON数据
                    prize_data = json.load(f)
                    
                    # 兼容旧格式（仅prizes列表）和新格式（包含prizes和prizes_count）
                    if isinstance(prize_data, dict) and "prizes" in prize_data:
                        self.prizes = prize_data.get("prizes", [])
                        self.prizes_count = prize_data.get("prizes_count", [])
                    elif isinstance(prize_data, list):
                        self.prizes = prize_data
                        self.prizes_count = [1] * len(self.prizes)  # 旧数据默认每次抽1个
                    else:
                        raise ValueError("奖项数据格式错误，必须是列表或包含prizes的字典")
                    
                    # 数据校验：过滤非字典元素
                    self.prizes = [p for p in self.prizes if isinstance(p, dict)]
                    # 补全prizes_count（确保数量与prizes一致）
                    while len(self.prizes_count) < len(self.prizes):
                        self.prizes_count.append(1)
                    self.prizes_count = self.prizes_count[:len(self.prizes)]  # 截断多余的
                    
                    # 清空表格后重新加载
                    for item in self.tree.get_children():
                        self.tree.delete(item)
                    
                    # 遍历有效奖项，插入表格（修正字段顺序）
                    for idx, prize in enumerate(self.prizes):
                        # 安全获取各字段
                        prize_type = prize.get("type", "")
                        text = prize.get("text", "")
                        title = prize.get("title", "")
                        count = prize.get("count", "")
                        draw_count = self.prizes_count[idx]  # 每次抽取数量
                        img = os.path.basename(prize.get("img", "")) if prize.get("img") else ""
                        
                        # 关键修复：values顺序与表格列（等级、奖项名称、奖品名称、总数量、每次抽取数量、图片）严格对应
                        self.tree.insert("", tk.END, values=(
                            prize_type, text, title, count, draw_count, img
                        ))
                print("奖项数据加载成功")
            else:
                # 文件不存在则初始化空列表
                self.prizes = []
                self.prizes_count = []
                print("奖项配置文件不存在，初始化空列表")
        except Exception as e:
            # 捕获并明确错误信息
            messagebox.showerror("错误", f"加载奖项数据失败: {str(e)}")
            self.prizes = []  # 兜底：初始化空列表
            self.prizes_count = []
    
    def create_prize_dialog(self, title, prize_data=None, draw_count=None):
        """创建添加/编辑奖项的对话框（新增每次抽取数量参数）"""
        dialog = tk.Toplevel(self.parent)
        dialog.title(title)
        dialog.geometry("500x450")  # 调整高度以容纳新输入框
        dialog.resizable(False, False)
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # 居中对话框
        dialog.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - dialog.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 输入字段
        ttk.Label(main_frame, text="奖项等级 (数字):").grid(row=0, column=0, sticky=tk.W, pady=5)
        type_entry = ttk.Entry(main_frame, width=30)
        type_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(main_frame, text="奖项名称 (如: 一等奖):").grid(row=1, column=0, sticky=tk.W, pady=5)
        text_entry = ttk.Entry(main_frame, width=30)
        text_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(main_frame, text="奖品名称:").grid(row=2, column=0, sticky=tk.W, pady=5)
        title_entry = ttk.Entry(main_frame, width=30)
        title_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(main_frame, text="总数量:").grid(row=3, column=0, sticky=tk.W, pady=5)
        count_entry = ttk.Entry(main_frame, width=30)
        count_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 新增：每次抽取数量输入框
        ttk.Label(main_frame, text="每次抽取数量:").grid(row=4, column=0, sticky=tk.W, pady=5)
        draw_count_entry = ttk.Entry(main_frame, width=30)
        draw_count_entry.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 图片上传部分
        ttk.Label(main_frame, text="奖品图片:").grid(row=5, column=0, sticky=tk.W, pady=5)
        
        img_frame = ttk.Frame(main_frame)
        img_frame.grid(row=5, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        img_label = ttk.Label(img_frame, text="未选择图片", width=25, relief=tk.SUNKEN)
        img_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.current_image_path = None
        
        def select_image():
            file_path = filedialog.askopenfilename(
                title="选择奖品图片",
                filetypes=[("图片文件", "*.jpg *.jpeg *.png *.gif *.bmp"), ("所有文件", "*.*")]
            )
            if file_path:
                self.current_image_path = file_path
                img_label.config(text=os.path.basename(file_path))
        
        ttk.Button(img_frame, text="选择图片", command=select_image, width=10).pack(side=tk.LEFT)
        
        # 如果有现有数据，填充字段
        if prize_data:
            type_entry.insert(0, prize_data.get("type", ""))
            text_entry.insert(0, prize_data.get("text", ""))
            title_entry.insert(0, prize_data.get("title", ""))
            count_entry.insert(0, prize_data.get("count", ""))
            # 填充每次抽取数量
            draw_count_entry.insert(0, draw_count if draw_count is not None else 1)
        
        # 按钮框架
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        return dialog, type_entry, text_entry, title_entry, count_entry, draw_count_entry, img_label
    
    def add_prize(self):
        # 新增：传入空的draw_count
        dialog, type_entry, text_entry, title_entry, count_entry, draw_count_entry, img_label = self.create_prize_dialog("添加奖项")
        
        def save():
            try:
                prize_type = int(type_entry.get().strip())
                count = int(count_entry.get().strip())
                # 新增：校验每次抽取数量
                draw_count = int(draw_count_entry.get().strip())
                if draw_count < 1:
                    messagebox.showwarning("警告", "每次抽取数量必须大于0")
                    return
                elif draw_count > count:
                    messagebox.showwarning("警告", "每次抽取数量不能大于总数量")
                    return
            except ValueError:
                messagebox.showwarning("警告", "奖项等级、总数量、每次抽取数量必须是数字")
                return
            
            text = text_entry.get().strip()
            title = title_entry.get().strip()
            
            if not all([text, title]):
                messagebox.showwarning("警告", "请填写奖项名称和奖品名称")
                return
            
            # 检查奖项等级是否重复
            for prize in self.prizes:
                if prize.get("type") == prize_type:
                    messagebox.showwarning("警告", f"奖项等级 {prize_type} 已存在")
                    return
            
            # 保存图片
            img_filename = ""
            if self.current_image_path:
                img_filename = os.path.basename(self.current_image_path)
                dest_path = os.path.join(self.img_path, img_filename)
                try:
                    shutil.copy2(self.current_image_path, dest_path)
                except Exception as e:
                    messagebox.showerror("错误", f"图片保存失败: {str(e)}")
                    return
            
            # 创建奖项对象
            prize = {
                "type": prize_type,
                "text": text,
                "title": title,
                "count": count
            }
            if img_filename:
                prize["img"] = f"../img/{img_filename}"
            
            # 添加到列表和表格
            self.prizes.append(prize)
            # 新增：添加每次抽取数量到数组
            self.prizes_count.append(draw_count)
            self.tree.insert("", tk.END, values=(
                prize_type, text, title, count, draw_count, img_filename
            ))
            dialog.destroy()
            messagebox.showinfo("成功", "奖项添加成功")
        
        ttk.Button(dialog, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(dialog, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
        
        type_entry.focus_set()
    
    def edit_prize(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要编辑的奖项")
            return
        
        item = selected[0]
        index = self.tree.index(item)
        prize = self.prizes[index]
        # 新增：获取该奖项的每次抽取数量
        draw_count = self.prizes_count[index] if index < len(self.prizes_count) else 1
        
        # 创建编辑对话框（传入draw_count）
        dialog, type_entry, text_entry, title_entry, count_entry, draw_count_entry, img_label = self.create_prize_dialog("编辑奖项", prize, draw_count)
        
        def save():
            try:
                prize_type = int(type_entry.get().strip())
                count = int(count_entry.get().strip())
                # 新增：校验每次抽取数量
                draw_count = int(draw_count_entry.get().strip())
                if draw_count < 1:
                    messagebox.showwarning("警告", "每次抽取数量必须大于0")
                    return
                elif draw_count > count:
                    messagebox.showwarning("警告", "每次抽取数量不能大于总数量")
                    return
            except ValueError:
                messagebox.showwarning("警告", "奖项等级、总数量、每次抽取数量必须是数字")
                return
            
            text = text_entry.get().strip()
            title = title_entry.get().strip()
            
            if not all([text, title]):
                messagebox.showwarning("警告", "请填写奖项名称和奖品名称")
                return
            
            # 检查奖项等级是否重复（排除自身）
            for i, p in enumerate(self.prizes):
                if p.get("type") == prize_type and i != index:
                    messagebox.showwarning("警告", f"奖项等级 {prize_type} 已存在")
                    return
            
            # 处理图片
            img_filename = ""
            if self.current_image_path:
                # 上传新图片
                img_filename = os.path.basename(self.current_image_path)
                dest_path = os.path.join(self.img_path, img_filename)
                try:
                    shutil.copy2(self.current_image_path, dest_path)
                except Exception as e:
                    messagebox.showerror("错误", f"图片保存失败: {str(e)}")
                    return
            elif prize.get("img"):
                # 使用原有图片
                img_filename = os.path.basename(prize.get("img", ""))
            
            # 更新数据
            self.prizes[index].update({
                "type": prize_type,
                "text": text,
                "title": title,
                "count": count
            })
            
            if img_filename:
                self.prizes[index]["img"] = f"../img/{img_filename}"
            elif "img" in self.prizes[index] and not self.current_image_path and img_label.cget("text") == "未选择图片":
                # 如果原本有图片但现在删除了
                del self.prizes[index]["img"]
            
            # 新增：更新每次抽取数量
            self.prizes_count[index] = draw_count
            
            # 更新表格
            self.tree.item(item, values=(
                prize_type, text, title, count, draw_count, img_filename
            ))
            
            dialog.destroy()
            messagebox.showinfo("成功", "奖项修改成功")
        
        ttk.Button(dialog, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(dialog, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)
        
        type_entry.focus_set()
        type_entry.select_range(0, tk.END)
    
    def delete_prize(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的奖项")
            return
        
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selected)} 个奖项吗？"):
            # 需要从后往前删除，避免索引变化
            items_to_delete = list(selected)
            items_to_delete.sort(reverse=True)  # 从后往前排序
            
            for item in items_to_delete:
                index = self.tree.index(item)
                self.prizes.pop(index)
                # 新增：删除对应位置的每次抽取数量
                if index < len(self.prizes_count):
                    self.prizes_count.pop(index)
                self.tree.delete(item)
    
    def save_prizes(self):
        try:
            # 确保prizes按type排序
            self.prizes.sort(key=lambda x: x.get("type", 0))
            # 同步prizes_count的排序（按prize的type排序后重新整理）
            sorted_pairs = sorted(zip(self.prizes, self.prizes_count), key=lambda x: x[0].get("type", 0))
            self.prizes, self.prizes_count = zip(*sorted_pairs)
            self.prizes = list(self.prizes)
            self.prizes_count = list(self.prizes_count)
            
            # 补全prizes_count（确保数量一致）
            while len(self.prizes_count) < len(self.prizes):
                self.prizes_count.append(1)
            
            # 构造新的JSON结构（包含prizes和prizes_count）
            save_data = {
                "prizes": self.prizes,
                "prizes_count": self.prizes_count
            }
            
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "奖项配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")


class LotteryManager:
    def __init__(self, root):
        self.root = root
        self.root.title("抽奖系统管理工具")
        self.root.geometry("900x750")
        
        # 设置窗口图标
        try:
            self.root.iconbitmap(default="")  # 可以设置图标文件路径
        except:
            pass
        
        # 设置样式
        self.setup_style()
        
        # 设置路径
        self.base_dir = Path(__file__).parent
        self.product_dir = self.base_dir / "product"
        self.flask_server_dir = self.base_dir / "flask_server"
        
        # 数据路径
        self.student_data_dir = self.flask_server_dir / "data"
        self.prize_data_dir = self.student_data_dir
        self.prize_img_dir = self.product_dir / "dist" / "img"
        
        # 创建必要的目录
        self.student_data_dir.mkdir(parents=True, exist_ok=True)
        self.prize_data_dir.mkdir(parents=True, exist_ok=True)
        self.prize_img_dir.mkdir(parents=True, exist_ok=True)
        
        # 创建主框架
        self.setup_main_frame()
        
        # 初始化管理器
        self.student_manager = StudentManager(self.main_frame, str(self.student_data_dir))
        self.prize_manager = PrizeManager(self.main_frame, str(self.prize_data_dir), str(self.prize_img_dir))
        
        # 抽奖按钮
        self.setup_lottery_button()
        
        # 状态栏
        self.setup_status_bar()
    
    def setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        # 配置颜色
        style.configure("TLabel", font=("Microsoft YaHei", 10))
        style.configure("TButton", font=("Microsoft YaHei", 10))
        style.configure("TLabelframe", font=("Microsoft YaHei", 11, "bold"))
        style.configure("Treeview", font=("Microsoft YaHei", 10))
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 10, "bold"))
        
        # 配置大按钮样式
        style.configure("Lottery.TButton", 
                       font=("Microsoft YaHei", 16, "bold"),
                       padding=20)
    
    def setup_main_frame(self):
        # 标题
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill=tk.X, padx=10, pady=10)
        
        title_label = ttk.Label(title_frame, text="🎯 抽奖系统管理工具 🎯", 
                               font=("Microsoft YaHei", 18, "bold"))
        title_label.pack()
        
        # 说明标签
        desc_label = ttk.Label(title_frame, 
                              text="请按照以下步骤配置抽奖系统：1. 配置学生信息 → 2. 配置奖项信息 → 3. 开始抽奖",
                              font=("Microsoft YaHei", 10))
        desc_label.pack(pady=5)
        
        # 主框架（带滚动条）
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建Canvas和滚动条
        canvas = tk.Canvas(main_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        self.main_frame = ttk.Frame(canvas)
        
        # 配置Canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # 布局
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 绑定事件
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        self.main_frame.bind("<Configure>", configure_canvas)
        
        def configure_window(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        canvas.bind("<Configure>", configure_window)
        
        # 绑定鼠标滚轮
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    def setup_lottery_button(self):
        # 抽奖按钮框架
        lottery_frame = ttk.Frame(self.main_frame)
        lottery_frame.pack(fill=tk.X, padx=10, pady=20)
        
        # 说明标签
        ttk.Label(lottery_frame, text="第三步：开始抽奖", 
                 font=("Microsoft YaHei", 10, "bold"), foreground="blue").pack(pady=(0, 10))
        
        # 大号抽奖按钮
        lottery_btn = ttk.Button(
            lottery_frame, 
            text="🎉 开始抽奖 🎉", 
            command=self.start_lottery,
            style="Lottery.TButton"
        )
        lottery_btn.pack(expand=True, fill=tk.X)
    
    def setup_status_bar(self):
        # 状态栏
        self.status_bar = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def start_lottery(self):
        # 检查学生数据
        student_file = self.student_data_dir / "users.xlsx"  # 修正原代码的笔误：user.xlsx → users.xlsx
        if not student_file.exists():
            messagebox.showwarning("警告", "请先上传学生信息")
            return
        
        # 检查奖项配置
        prize_file = self.prize_data_dir / "prizes.json"
        if not prize_file.exists():
            messagebox.showwarning("警告", "请先配置奖项信息")
            return
        
        # 启动抽奖界面
        self.show_lottery_interface()
    
    def show_lottery_interface(self):
        # 弹出对话框提示用户
        web_app = WebApp()
        flask_thread = threading.Thread(target=web_app.run, daemon=True)  # 修正：target=web_app.run() → target=web_app.run
        flask_thread.start()
        
        # 等待一小段时间确保服务器启动
        import time
        time.sleep(1)
        
        # 弹出对话框提示用户
        messagebox.showinfo(
            "抽奖界面已启动",
            "请在浏览器中输入以下地址访问抽奖界面：\n\n"
            "http://127.0.0.1:8090/\n\n"
            "点击确定后在浏览器中打开。"
        )
        # 自动打开浏览器
        webbrowser.open("http://127.0.0.1:8090/")

    def create_lottery_ui(self, window):
        # 标题
        title_label = ttk.Label(window, text="🎯 抽奖进行中 🎯", 
                               font=("Microsoft YaHei", 24, "bold"))
        title_label.pack(pady=20)
        
        # 结果显示区域
        result_frame = ttk.Frame(window)
        result_frame.pack(expand=True, fill=tk.BOTH, padx=50, pady=20)
        
        result_label = ttk.Label(
            result_frame, 
            text="等待开始...", 
            font=("Microsoft YaHei", 48, "bold"),
            foreground="#FF6B6B",
            anchor=tk.CENTER
        )
        result_label.pack(expand=True, fill=tk.BOTH)
        
        # 当前奖项显示
        current_prize_frame = ttk.LabelFrame(window, text="当前奖项", padding=10)
        current_prize_frame.pack(fill=tk.X, padx=50, pady=10)
        
        prize_label = ttk.Label(current_prize_frame, text="未设置", 
                               font=("Microsoft YaHei", 14))
        prize_label.pack()
        
        # 控制按钮
        control_frame = ttk.Frame(window)
        control_frame.pack(fill=tk.X, padx=50, pady=20)
        
        def start_drawing():
            result_label.config(text="抽奖中...", foreground="#FF6B6B")
            prize_label.config(text="特等奖 - 神秘大礼")
            # 这里可以添加实际的抽奖逻辑
        
        def stop_drawing():
            result_label.config(text="中奖者：张三\n班级：计算机1班\n学号：20230001", 
                              foreground="#1E90FF", font=("Microsoft YaHei", 36, "bold"))
        
        def reset_drawing():
            result_label.config(text="等待开始...", foreground="#FF6B6B", 
                              font=("Microsoft YaHei", 48, "bold"))
            prize_label.config(text="未设置")
        
        ttk.Button(control_frame, text="开始抽奖", 
                  command=start_drawing).pack(side=tk.LEFT, padx=10, ipadx=20)
        ttk.Button(control_frame, text="停止抽奖", 
                  command=stop_drawing).pack(side=tk.LEFT, padx=10, ipadx=20)
        ttk.Button(control_frame, text="重置", 
                  command=reset_drawing).pack(side=tk.LEFT, padx=10, ipadx=20)
        ttk.Button(control_frame, text="关闭", 
                  command=window.destroy).pack(side=tk.LEFT, padx=10, ipadx=20)
    
    def update_status(self, message):
        self.status_bar.config(text=message)


def main():
    root = tk.Tk()
    app = LotteryManager(root)
    
    # 设置窗口居中
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # 设置最小窗口大小
    root.minsize(900, 600)
    
    root.mainloop()


if __name__ == "__main__":
    main()