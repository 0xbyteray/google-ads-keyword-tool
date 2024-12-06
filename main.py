import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import yaml
import os
import csv
from keyword_ideas_service import KeywordIdeasService
from google.ads.googleads.errors import GoogleAdsException
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
import platform
import requests
from bs4 import BeautifulSoup
import time
import random
from kgr_calculator import KGRCalculator

# 根据操作系统设置matplotlib中文字体支持
if platform.system() == 'Windows':
    matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # Windows的中文字体
else:  # macOS
    matplotlib.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # macOS 系统自带的字体
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

class GoogleAdsKeywordTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Ads Keyword Tool")
        
        # 设置窗口大小和位置
        window_width = 1600
        window_height = 900
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        self.service = None
        
        # 创建左右分隔的主框架
        self.main_paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))  # 修改下边距为0
        
        # 左侧框架
        self.left_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.left_frame, weight=1)
        
        # 右侧框架
        self.right_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.right_frame, weight=2)
        
        # 创建状态栏
        self.status_frame = ttk.Frame(root)
        self.status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_text = tk.Text(self.status_frame, height=6, wrap=tk.WORD)
        self.status_text.pack(fill=tk.X)
        self.status_text.config(state=tk.DISABLED)
        
        # 加载配置
        self.load_config()
        
        # 初始化服务
        self.keyword_service = None
        self.initialize_service()
        
        # 初始化KGR计算器
        self.kgr_calculator = KGRCalculator()
        
        # 创建输入区域
        self.create_input_area()
        
        # 创建结果展示区域
        self.create_result_area()
        
        # 创建右侧月度数据展示区域
        self.create_monthly_trend_area()
        
        # 绑定单元格点击事件
        self.result_table.bind('<ButtonRelease-1>', self.handle_cell_click)
        
        # 添加User-Agent池
        self.user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
        ]

    def load_refresh_token(self) -> str:
        """从refresh_token.txt加载refresh token"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            token_path = os.path.join(current_dir, '.refresh_token')
            
            if not os.path.exists(token_path):
                raise FileNotFoundError(f"找不到refresh token文件: {token_path}")
                
            with open(token_path, 'r') as f:
                refresh_token = f.read().strip()
                
            if not refresh_token:
                raise ValueError("refresh token文件为空")
                
            return refresh_token
                
        except Exception as e:
            raise Exception(f"加载refresh token失败: {str(e)}")

    def load_config(self) -> dict:
        """加载所有必要的配置"""
        try:
            # 加载YAML配置
            current_dir = os.path.dirname(os.path.abspath(__file__))
            yaml_path = os.path.join(current_dir, 'config.yaml')
            
            if not os.path.exists(yaml_path):
                raise FileNotFoundError(f"找不到配置文件: {yaml_path}")
                
            with open(yaml_path, 'r') as f:
                yaml_config = yaml.safe_load(f)

            # 加载refresh token
            refresh_token = self.load_refresh_token()
             
            # 合并配置
            config_dict = {
                'client_id': yaml_config.get('client_id'),
                'client_secret': yaml_config.get('client_secret'),
                'developer_token': yaml_config.get('developer_token'),
                'login_customer_id': yaml_config.get('login_customer_id'),
                'refresh_token': refresh_token
            }
            
            # 验证必要字段
            missing_keys = [k for k, v in config_dict.items() if not v]
            if missing_keys:
                raise ValueError(f"配置文件中缺少必要字段: {', '.join(missing_keys)}")
                
            return config_dict
            
        except Exception as e:
            raise Exception(f"加载配置失败: {str(e)}")

    def initialize_service(self):
        """初始化关键词服务"""
        try:
            # 检查必要的配置文件是否存在
            current_dir = os.path.dirname(os.path.abspath(__file__))
            required_files = {
                '.refresh_token': 'Refresh Token文件',
                'config.yaml': 'YAML配置文件'
            }
            
            missing_files = []
            for file_name, desc in required_files.items():
                if not os.path.exists(os.path.join(current_dir, file_name)):
                    missing_files.append(f"{desc} ({file_name})")
            
            if missing_files:
                error_msg = "缺少以下必要文件：\n" + "\n".join(missing_files)
                self.update_status(error_msg)
                messagebox.showerror("初始化失败", error_msg)
                return
                
            config = self.load_config()
            self.keyword_service = KeywordIdeasService(config)
            self.update_status("Google Ads API 服务初始化成功")
        except Exception as e:
            error_msg = f"初始化服务失败: {str(e)}"
            self.update_status(error_msg)
            messagebox.showerror("错误", error_msg)

    def create_input_area(self):
        """创建输入区域"""
        # 输入区域框架
        input_frame = ttk.LabelFrame(self.left_frame, text="搜索条件", padding=(10, 5, 10, 10))
        input_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        
        # 关键词输入区域
        keyword_label = ttk.Label(input_frame, text="关键词（每行一个）:")
        keyword_label.pack(fill=tk.X, pady=(5, 0))
        
        self.keyword_input = tk.Text(input_frame, height=4)  
        self.keyword_input.pack(fill=tk.X, pady=(2, 10))  
        
        # URL输入区域
        url_label = ttk.Label(input_frame, text="网址（可选）:")
        url_label.pack(fill=tk.X)
        
        self.url_input = ttk.Entry(input_frame)
        self.url_input.pack(fill=tk.X, pady=(0, 10))
        
        # 按钮区域
        button_frame = ttk.Frame(input_frame)
        button_frame.pack(fill=tk.X)
        
        search_button = ttk.Button(button_frame, text="搜索关键词", command=self.search_keywords)
        search_button.pack(side=tk.LEFT, padx=5)
        
        clear_button = ttk.Button(button_frame, text="清空关键词", command=self.clear_keywords)
        clear_button.pack(side=tk.LEFT)

        export_button = ttk.Button(button_frame, text="导出结果", command=self.export_results)
        export_button.pack(side=tk.LEFT, padx=5)
        
    def create_result_area(self):
        """创建结果展示区域"""
        # 结果区域框架
        result_frame = ttk.LabelFrame(self.left_frame, text="搜索结果", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # 创建表格容器，用于放置表格和滚动条
        table_container = ttk.Frame(result_frame)
        table_container.pack(fill=tk.BOTH, expand=True)
        
        # 配置网格权重，使表格可以跟随窗口调整大小
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)
        
        # 创建表格
        columns = ('keyword', 'avg_monthly_searches', 'competition', 'competition_index',
                  'recent_growth', 'growth', 'low_cpc', 'high_cpc', 'kgr')
        self.result_table = ttk.Treeview(table_container, columns=columns, show='headings', height=20)
        
        # 创建自定义样式
        style = ttk.Style()
        style.configure('Custom.Treeview', rowheight=30)  # 设置更大的行高
        
        # 设置列标题和排序事件
        for col in columns:
            if col != 'kgr':  # KGR列不需要排序功能
                self.result_table.heading(col, text=self.get_column_title(col),
                                    command=lambda c=col: self.treeview_sort_column(self.result_table, c, False))
            else:
                self.result_table.heading(col, text='KGR(avg/latest)')  # KGR列的标题
        
        # 设置列宽
        self.result_table.column('keyword', width=200, minwidth=150)
        self.result_table.column('avg_monthly_searches', width=120, minwidth=100)
        self.result_table.column('competition', width=80, minwidth=80)
        self.result_table.column('competition_index', width=80, minwidth=80)
        self.result_table.column('recent_growth', width=100, minwidth=100)
        self.result_table.column('growth', width=80, minwidth=80)
        self.result_table.column('low_cpc', width=100, minwidth=100)
        self.result_table.column('high_cpc', width=100, minwidth=100)
        self.result_table.column('kgr', width=80, minwidth=80)  # KGR列的宽度
        
        # 添加垂直滚动条
        vsb = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.result_table.yview)
        self.result_table.configure(yscrollcommand=vsb.set)
        
        # 添加水平滚动条
        hsb = ttk.Scrollbar(table_container, orient=tk.HORIZONTAL, command=self.result_table.xview)
        self.result_table.configure(xscrollcommand=hsb.set)
        
        # 使用网格布局放置表格和滚动条
        self.result_table.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # 绑定选择事件
        self.result_table.bind('<<TreeviewSelect>>', self.on_item_select)
        self.result_table.bind('<Button-1>', self.on_click)
        
        # 存储Entry组件的字典
        self.keyword_entries = {}
        self.current_entry = None  # 添加当前显示的Entry引用

    def hide_current_entry(self, event=None):
        """隐藏当前显示的Entry"""
        if self.current_entry:
            self.current_entry.place_forget()
            self.current_entry = None

    def on_click(self, event):
        """处理单击事件"""
        # 先隐藏当前显示的Entry
        self.hide_current_entry()
        
        region = self.result_table.identify('region', event.x, event.y)
        if region != "cell":
            return
            
        # 获取点击的行和列
        row_id = self.result_table.identify_row(event.y)
        col = self.result_table.identify_column(event.x)
        
        # 只处理关键词列
        if col != '#1':  # 第一列
            return
            
        # 获取单元格的坐标
        bbox = self.result_table.bbox(row_id, 'keyword')
        if not bbox:
            return
            
        # 如果这个单元格还没有Entry
        if row_id not in self.keyword_entries:
            # 获取关键词
            keyword = self.result_table.item(row_id)['values'][0]
            
            # 创建Entry，使用自定义样式
            entry = ttk.Entry(self.result_table)
            entry.insert(0, keyword)
            entry.configure(state='readonly')  # 设置为只读
            
            # 绑定失去焦点事件
            entry.bind('<FocusOut>', self.hide_current_entry)
            
            # 存储Entry
            self.keyword_entries[row_id] = entry
        
        # 显示Entry，调整位置和大小以完全显示文本
        entry = self.keyword_entries[row_id]
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=28)  # 固定高度为28像素
        entry.select_range(0, tk.END)  # 全选文本
        entry.focus_set()  # 设置焦点
        
        # 更新当前显示的Entry引用
        self.current_entry = entry

    def create_monthly_trend_area(self):
        """创建右侧月度数据展示区域"""
        # 标题
        self.trend_title_label = ttk.Label(self.right_frame, text="月度搜索趋势", font=('Helvetica', 12, 'bold'))
        self.trend_title_label.pack(fill=tk.X, pady=(0, 10))
        self.trend_title_label.pack_forget()  # 初始隐藏
        
        # 关键词标签
        self.trend_keyword_label = ttk.Label(self.right_frame, text="", font=('Helvetica', 11))
        self.trend_keyword_label.pack(fill=tk.X, pady=(0, 10))
        self.trend_keyword_label.pack_forget()  # 初始隐藏
        
        # 增长率标签
        self.growth_label = ttk.Label(self.right_frame, text="", font=('Helvetica', 11))
        self.growth_label.pack(fill=tk.X, pady=(0, 10))
        self.growth_label.pack_forget()  # 初始隐藏
        
        # 创建图表区域
        self.fig = Figure(figsize=(6, 3), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.right_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.canvas_widget.pack_forget()  # 初始隐藏
        
        # 月度数据表格
        columns = ('month', 'searches', 'growth')
        self.trend_table = ttk.Treeview(self.right_frame, columns=columns, show='headings', height=8)
        
        # 设置列标题
        self.trend_table.heading('month', text='月份')
        self.trend_table.heading('searches', text='搜索量')
        self.trend_table.heading('growth', text='环比增长')
       
        # 设置列宽
        self.trend_table.column('month', width=100)
        self.trend_table.column('searches', width=100)
        self.trend_table.column('growth', width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.right_frame, orient=tk.VERTICAL, command=self.trend_table.yview)
        self.trend_table.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.trend_table.pack(fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始隐藏表格和滚动条
        self.trend_table.pack_forget()
        scrollbar.pack_forget()
        
    def show_trend_widgets(self):
        """显示趋势相关的所有组件"""
        self.trend_title_label.pack(fill=tk.X, pady=(0, 10))
        self.trend_keyword_label.pack(fill=tk.X, pady=(0, 10))
        self.growth_label.pack(fill=tk.X, pady=(0, 10))
        self.canvas_widget.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.trend_table.pack(fill=tk.BOTH, expand=True)
        
    def on_item_select(self, event):
        """处理表格项目选择事件"""
        selection = self.result_table.selection()
        if not selection:
            return
            
        # 获取选中项
        item = selection[0]
        keyword = self.result_table.item(item)['values'][0]
        
        # 更新月度趋势数据
        self.update_monthly_trend(keyword)
        
    def update_monthly_trend(self, keyword):
        """更新月度趋势数据显示"""
        # 显示所有趋势相关的组件
        self.show_trend_widgets()
        
        # 清空现有数据
        for item in self.trend_table.get_children():
            self.trend_table.delete(item)
            
        # 获取选中关键词的数据
        keyword_data = next((data for data in self.search_results if data.text == keyword), None)
        if not keyword_data:
            return
            
        # 更新关键词标签
        self.trend_keyword_label.config(text=f"关键词: {keyword}")
        
        # 更新增长率标签
        growth_text = f"年增长: {self.format_growth_rate(keyword_data.growth_percentage)}"
        if keyword_data.growth_percentage > 0:
            growth_text += " ↑"
        elif keyword_data.growth_percentage < 0:
            growth_text += " ↓"
            
        recent_growth_text = f"近3月增长: {self.format_growth_rate(keyword_data.recent_growth_percentage)}"
        if keyword_data.recent_growth_percentage > 0:
            recent_growth_text += " ↑"
        elif keyword_data.recent_growth_percentage < 0:
            recent_growth_text += " ↓"
            
        self.growth_label.config(text=f"{growth_text}    {recent_growth_text}")
        
        # 按时间倒序排列月度数据
        monthly_data = sorted(keyword_data.monthly_searches, key=lambda x: x.year_month, reverse=True)
        
        # 填充表格
        for i in range(len(monthly_data)):
            current_volume = monthly_data[i].monthly_searches
            if i < len(monthly_data) - 1:  # 还有下一个月的数据（较早的月份）
                prev_volume = monthly_data[i+1].monthly_searches  # 上个月（较早的月份）
                if prev_volume == 0:
                    growth = float('inf') if current_volume > 0 else 0
                else:
                    growth = ((current_volume - prev_volume) / prev_volume) * 100
                growth_text = "∞" if growth == float('inf') else f"{growth:.1f}%"
            else:
                growth_text = "-"
                
            self.trend_table.insert('', tk.END, values=(
                monthly_data[i].year_month,
                self.format_number(monthly_data[i].monthly_searches),
                growth_text
            ))
            
        # 更新趋势图
        self.update_trend_chart(monthly_data)
        
    def update_trend_chart(self, monthly_data):
        """更新趋势图"""
        # 清空现有图表
        self.ax.clear()
        
        # 准备数据
        dates = [data.year_month for data in reversed(monthly_data)]  # 按时间正序
        volumes = [data.monthly_searches for data in reversed(monthly_data)]
        
        # 绘制折线图
        self.ax.plot(dates, volumes, marker='o')
        
        # 设置标签和标题
        self.ax.set_xlabel('月份')
        self.ax.set_ylabel('搜索量')
        
        # 旋转x轴标签以防重叠
        self.ax.tick_params(axis='x', rotation=45)
        
        # 自动调整布局
        self.fig.tight_layout()
        
        # 刷新画布
        self.canvas.draw()

    def search_keywords(self):
        """搜索关键词"""
        if not self.keyword_service:
            messagebox.showerror("错误", "Google Ads API 服务未初始化")
            return
            
        # 获取输入的关键词
        keywords = [kw.strip() for kw in self.keyword_input.get("1.0", tk.END).splitlines() if kw.strip()]
        # 获取输入的URL
        url = self.url_input.get().strip()
        
        if not keywords and not url:
            messagebox.showwarning("提示", "请输入关键词或网址")
            return
            
        try:
            # 清空现有结果
            for item in self.result_table.get_children():
                self.result_table.delete(item)
                
            self.update_status("正在搜索关键词创意...")
            
            # 调用服务获取关键词创意
            self.search_results = self.keyword_service.generate_keyword_ideas(
                keywords=keywords if keywords else None,
                url=url if url else None
            )

            # 显示结果
            for idea in self.search_results:
                self.result_table.insert('', tk.END, values=(
                    idea.text,
                    self.format_number(idea.avg_monthly_searches),
                    idea.competition,
                    idea.competition_index,
                    self.format_growth_rate(idea.recent_growth_percentage),
                    self.format_growth_rate(idea.growth_percentage),
                    f"${idea.low_cpc:.2f}",
                    f"${idea.high_cpc:.2f}",
                    "点击计算"  # KGR列的初始值
                ))
                
            self.update_status(f"成功获取 {len(self.search_results)} 个关键词的相关数据")
        except GoogleAdsException as ex:
            self.update_status(f"Google Ads API 错误: {ex.error.message}")
            messagebox.showerror("API错误", ex.error.message)
        except Exception as e:
            self.update_status(f"发生错误: {str(e)}")
            messagebox.showerror("错误", str(e))

    def clear_keywords(self):
        """清空输入"""
        self.keyword_input.delete("1.0", tk.END)
        self.url_input.delete(0, tk.END)
        self.update_status("已清空搜索条件")

    def update_status(self, message):
        if hasattr(self, 'status_text'):
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, message + "\n")
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)

    def treeview_sort_column(self, tree, col, reverse):
        """
        表格列排序
        
        Args:
            tree: Treeview对象
            col: 列名
            reverse: 是否反向排序
        """
        # 获取所有项目的ID
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        
        # 根据列类型进行不同的排序处理
        if col in ['avg_monthly_searches', 'competition_index', 'low_cpc', 'high_cpc']:
            # 数值型列，需要转换为float进行排序
            l.sort(key=lambda x: float(x[0].replace('$', '').replace(',', '')), reverse=reverse)
        elif col in ['recent_growth', 'growth']:
            # 增长率列，需要特殊处理无穷大值
            def growth_sort_key(x):
                value = x[0].replace('%', '').strip()
                if value == '∞':
                    return float('inf')
                return float(value)
            l.sort(key=growth_sort_key, reverse=reverse)
        elif col == 'competition':
            # 竞争度列，使用竞争指数的值排序
            competition_index_list = [(tree.set(k, 'competition_index'), k) for k in tree.get_children('')]
            competition_index_list.sort(key=lambda x: float(x[0]), reverse=reverse)
            # 重新排序
            for index, (val, k) in enumerate(competition_index_list):
                tree.move(k, '', index)
            # 更新标题并返回，避免重复排序
            for column in tree['columns']:
                if column == col:
                    tree.heading(column, text=f"{self.get_column_title(column)} {'↓' if reverse else '↑'}")
                else:
                    tree.heading(column, text=self.get_column_title(column))
            # 重新绑定点击事件，切换排序方向
            tree.heading(col, command=lambda: self.treeview_sort_column(tree, col, not reverse))
            return
        else:
            # 文本列，直接排序
            l.sort(reverse=reverse)
            
        # 重新排序
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)
            
        # 更新所有列的标题，添加排序指示器
        for column in tree['columns']:
            if column == col:
                tree.heading(column, text=f"{self.get_column_title(column)} {'↓' if reverse else '↑'}")
            else:
                tree.heading(column, text=self.get_column_title(column))
        
        # 重新绑定点击事件，切换排序方向
        tree.heading(col, command=lambda: self.treeview_sort_column(tree, col, not reverse))

    def extract_numeric_value(self, text):
        """
        从文本中提取数值
        
        Args:
            text: 要处理的文本
            
        Returns:
            float: 提取的数值，如果无法提取则返回0
        """
        try:
            # 移除千位分隔符和货币符号
            text = text.replace(',', '').replace('$', '')
            # 移除百分号并转换为浮点数
            if '%' in text:
                return float(text.replace('%', ''))
            return float(text)
        except (ValueError, TypeError):
            return 0
    
    def get_column_title(self, column):
        """
        获取列的原始标题
        
        Args:
            column: 列名
            
        Returns:
            str: 列的原始标题
        """
        titles = {
            'keyword': '关键词',
            'avg_monthly_searches': '月均搜索量',
            'competition': '竞争度',
            'competition_index': '竞争指数',
            'recent_growth': '近三月增长率',
            'growth': '年增长率',
            'low_cpc': '首页最低CPC',
            'high_cpc': '首页最高CPC',
            'kgr': 'KGR(avg/latest)'
        }
        return titles.get(column, column)

    def format_growth_rate(self, rate):
        """格式化增长率显示"""
        if rate == float('inf'):
            return "∞"  # 显示无穷大符号
        return f"{rate:.1f}%"

    def format_number(self, number):
        """格式化数字显示，添加千位分隔符"""
        return f"{number:,}"

    def export_results(self):
        """导出搜索结果到CSV文件"""
        if not hasattr(self, 'search_results') or not self.search_results:
            messagebox.showwarning("提示", "没有可导出的搜索结果")
            return
            
        # 让用户选择保存位置
        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            title='选择保存位置'
        )
        
        if not file_path:  # 用户取消了保存
            return
            
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:  # 使用 utf-8-sig 以支持Excel正确显示中文
                writer = csv.writer(f)
                # 写入表头
                writer.writerow([
                    '关键词',
                    '月均搜索量',
                    '竞争度',
                    '竞争指数',
                    '近三月增长率',
                    '年增长率',
                    '首页最低CPC',
                    '首页最高CPC'
                ])
                
                # 写入数据
                for idea in self.search_results:
                    writer.writerow([
                        idea.text,
                        idea.avg_monthly_searches,
                        idea.competition,
                        idea.competition_index,
                        f"{idea.recent_growth_percentage:.1f}%" if idea.recent_growth_percentage != float('inf') else "∞",
                        f"{idea.growth_percentage:.1f}%" if idea.growth_percentage != float('inf') else "∞",
                        f"${idea.low_cpc:.2f}",
                        f"${idea.high_cpc:.2f}"
                    ])
                    
            self.update_status(f"搜索结果已导出到：{file_path}")
            messagebox.showinfo("成功", "搜索结果导出成功！")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def handle_cell_click(self, event):
        """处理单元格点击事件"""
        region = self.result_table.identify_region(event.x, event.y)
        if region == "cell":
            # 获取点击的列和行
            column = self.result_table.identify_column(event.x)
            item = self.result_table.identify_row(event.y)
            
            # 如果点击的是KGR列（第9列）
            if column == '#9' and item:
                self.calculate_kgr(item)

    def calculate_kgr(self, item_id):
        """计算KGR值"""
        item = self.result_table.item(item_id)
        values = item['values']
        
        # 如果已经计算过KGR，就不重复计算
        if values[8] != "点击计算":
            return
            
        # 获取关键词
        keyword = values[0]
        
        # 从搜索结果中找到对应的关键词数据
        keyword_data = None
        for idea in self.search_results:
            if idea.text == keyword:
                keyword_data = idea
                break
                
        if not keyword_data or not keyword_data.monthly_searches:
            messagebox.showerror("错误", "无法获取关键词的历史数据")
            return
            
        try:
            self.update_status(f"计算 '{keyword}' 的KGR，这个功能需要访问 Google 搜索，可能会受到 Google 的访问限制，注意控制使用频率...")

            # 获取最近一个月的搜索量
            latest_search_volume = keyword_data.monthly_searches[-1].monthly_searches
            
            # 计算KGR
            kgr_avg, kgr_latest, allintitle_count = self.kgr_calculator.calculate(keyword, latest_search_volume, keyword_data.avg_monthly_searches)
            
            # 更新表格中的KGR值
            values = list(values)
            values[8] = f"{kgr_avg:.3f} ({kgr_latest:.3f})"  # KGR值保留三位小数
            self.result_table.item(item_id, values=values)
            
            # 更新状态
            self.update_status(f"KGR计算完成 - 月均搜索量： {keyword_data.avg_monthly_searches}, 最近一个月搜索量: {latest_search_volume}, allintitle: {allintitle_count}")
            
        except (ValueError, ZeroDivisionError, IndexError) as e:
            messagebox.showerror("错误", f"计算KGR时出错：{str(e)}")
            
def main():
    root = tk.Tk()
    app = GoogleAdsKeywordTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
