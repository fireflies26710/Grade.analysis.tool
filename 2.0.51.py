import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib as mpl
import warnings
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


class EnhancedScoreAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("智能成绩分析系统 v2.0.51")

        # 初始化数据存储
        self.dataset = {}
        self.current_semester = ""
        self.grade_subjects = {
            '七年级': ['语文', '数学', '英语', '道德与法治', '历史', '地理', '生物', '体育'],
            '八年级': ['语文', '数学', '英语', '物理', '道德与法治', '历史', '地理', '生物', '体育'],
            '九年级': ['语文', '数学', '英语', '物理', '化学', '历史', '道德与法治', '体育']
        }
        self.grade_standards = {
            '七年级': {'优秀': 90, '良好': 80, '及格': 60},
            '八年级': {'优秀': 90, '良好': 80, '及格': 60},
            '九年级': {'优秀': 75, '良好': 60, '及格': 50}
        }
        self.full_marks = {}
        self.custom_subjects = {}

        # 创建界面组件
        self.create_widgets()
        self.create_semester_menu()

    # === 界面组件 ===
    def create_widgets(self):
        """创建主界面组件"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板")
        control_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")

        # 学期管理
        ttk.Button(control_frame, text="新建学期", command=self.create_semester).grid(row=0, column=0, padx=5)
        self.semester_combo = ttk.Combobox(control_frame, state="readonly")
        self.semester_combo.grid(row=0, column=1, padx=5)
        self.semester_combo.bind("<<ComboboxSelected>>", self.select_semester)

        # 年级选择
        ttk.Label(control_frame, text="当前年级：").grid(row=1, column=0)
        self.grade_combo = ttk.Combobox(control_frame,
                                        values=list(self.grade_subjects.keys()),
                                        state="readonly")
        self.grade_combo.grid(row=1, column=1, padx=5)
        self.grade_combo.bind("<<ComboboxSelected>>", self.update_grade_subjects)

        # 数据操作
        ttk.Button(control_frame, text="保存数据", command=self.save_data).grid(row=2, column=0, pady=5)
        ttk.Button(control_frame, text="加载数据", command=self.load_data).grid(row=2, column=1, pady=5)
        ttk.Button(control_frame, text="自定义学科", command=self.customize_subjects).grid(row=3, column=0, pady=5)
        ttk.Button(control_frame, text="设置满分", command=self.set_full_marks).grid(row=3, column=1, pady=5)
        ttk.Button(control_frame, text="生成报告", command=self.generate_report).grid(row=4, column=0, columnspan=2,
                                                                                      pady=5)

        # 成绩录入面板
        input_frame = ttk.LabelFrame(main_frame, text="成绩录入")
        input_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        # 学科选择
        ttk.Label(input_frame, text="选择学科：").grid(row=0, column=0)
        self.subject_combo = ttk.Combobox(input_frame, state="readonly")
        self.subject_combo.grid(row=0, column=1, padx=5)

        # 分数输入
        ttk.Label(input_frame, text="输入分数：").grid(row=1, column=0)
        self.score_entry = ttk.Entry(input_frame)
        self.score_entry.grid(row=1, column=1, padx=5)

        # 操作按钮
        ttk.Button(input_frame, text="添加成绩", command=self.add_score).grid(row=2, column=0, columnspan=2, pady=5)

        # 成绩表格
        self.tree_frame = ttk.Frame(main_frame)
        self.tree_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")

        columns = ("学科", "分数", "满分", "等级")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings")
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure('warning', foreground='red')

        # 分析面板
        analysis_frame = ttk.LabelFrame(main_frame, text="数据分析")
        analysis_frame.grid(row=0, column=1, rowspan=3, padx=10, pady=5, sticky="nsew")

        # 分析类型选择
        ttk.Label(analysis_frame, text="分析模式：").grid(row=0, column=0)
        self.analysis_mode = ttk.Combobox(analysis_frame,
                                          values=['学期分析', '趋势分析'],
                                          state="readonly")
        self.analysis_mode.grid(row=0, column=1, padx=5)
        self.analysis_mode.current(0)
        self.analysis_mode.bind("<<ComboboxSelected>>", self.toggle_analysis_mode)

        # 分析结果显示区域
        self.result_frame = ttk.Frame(analysis_frame)
        self.result_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="nsew")

        # 图表导出按钮
        ttk.Button(analysis_frame, text="导出图表", command=self.export_chart).grid(row=2, column=0, columnspan=2)

    # === 核心功能 ===
    def create_semester_menu(self):
        """初始化学期菜单"""
        self.semester_combo["values"] = list(self.dataset.keys())
        if self.semester_combo["values"]:
            self.semester_combo.current(0)
            self.select_semester()
        else:
            self.current_semester = ""

    def create_semester(self):
        """创建新学期"""
        semester_name = f"{datetime.now().year}-{datetime.now().year + 1} 第{len(self.dataset) + 1}学期"
        self.dataset[semester_name] = {
            'grade': '七年级',
            'scores': {},
            'subjects': []
        }
        self.semester_combo["values"] = list(self.dataset.keys())
        self.semester_combo.set(semester_name)
        self.current_semester = semester_name
        self.grade_combo.set('七年级')
        self.update_grade_subjects()
        self.update_data_table()

    def select_semester(self, event=None):
        """选择学期"""
        selected_semester = self.semester_combo.get()
        if selected_semester in self.dataset:
            self.current_semester = selected_semester
            current_data = self.dataset[self.current_semester]
            self.grade_combo.set(current_data['grade'])
            self.update_grade_subjects()
            self.update_data_table()

    def update_grade_subjects(self, event=None):
        """更新年级相关设置"""
        if self.current_semester:
            selected_grade = self.grade_combo.get()
            self.dataset[self.current_semester]['grade'] = selected_grade
            subjects = self.grade_subjects[selected_grade] + self.custom_subjects.get(selected_grade, [])
            self.subject_combo["values"] = subjects
            self.subject_combo.current(0) if subjects else None

    def add_score(self):
        """添加成绩到当前学期"""
        if not self.current_semester:
            messagebox.showwarning("警告", "请先创建或选择学期！")
            return

        subject = self.subject_combo.get()
        score = self.score_entry.get()

        if not subject:
            messagebox.showwarning("警告", "请选择学科！")
            return

        if not score.isdigit():
            messagebox.showwarning("警告", "请输入有效数字分数！")
            return

        score = float(score)
        full_mark = self.full_marks.get(subject, 100)

        if score > full_mark:
            messagebox.showerror("错误", f"分数不能超过该学科满分值{full_mark}")
            return

        self.dataset[self.current_semester]['scores'][subject] = score
        if subject not in self.dataset[self.current_semester]['subjects']:
            self.dataset[self.current_semester]['subjects'].append(subject)

        self.score_entry.delete(0, tk.END)
        self.update_data_table()

    # === 数据持久化 ===
    def save_data(self):
        """保存全部数据到JSON文件"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json")])
        if not filepath:
            return

        try:
            data_to_save = {
                "dataset": self.dataset,
                "full_marks": self.full_marks,
                "custom_subjects": self.custom_subjects
            }
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "数据保存成功！")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")

    def load_data(self):
        """从JSON文件加载数据"""
        filepath = filedialog.askopenfilename(
            filetypes=[("JSON文件", "*.json")])
        if not filepath:
            return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                loaded_data = json.load(f)

            # 数据验证
            if not all(key in loaded_data for key in ["dataset", "full_marks", "custom_subjects"]):
                raise ValueError("文件格式不正确")

            self.dataset = loaded_data["dataset"]
            self.full_marks = loaded_data["full_marks"]
            self.custom_subjects = loaded_data["custom_subjects"]

            # 更新界面
            self.create_semester_menu()
            if self.current_semester:
                self.update_grade_subjects()
                self.update_data_table()
            messagebox.showinfo("成功", "数据加载成功！")
        except Exception as e:
            messagebox.showerror("错误", f"加载失败：{str(e)}")

    # === 数据分析 ===
    def toggle_analysis_mode(self, event=None):
        """切换分析模式"""
        for widget in self.result_frame.winfo_children():
            widget.destroy()

        if self.analysis_mode.get() == '学期分析':
            self.show_semester_analysis()
        else:
            self.show_trend_analysis()

    def show_semester_analysis(self):
        """显示学期分析"""
        if not self.current_semester or not self.dataset[self.current_semester]['scores']:
            messagebox.showwarning("警告", "当前学期无成绩数据！")
            return

        mpl.rcParams['font.sans-serif'] = ['SimHei']
        mpl.rcParams['axes.unicode_minus'] = False

        # 创建图表
        fig = plt.figure(figsize=(10, 5))
        gs = fig.add_gridspec(1, 2)
        ax1 = fig.add_subplot(gs[0, 0])
        ax2 = fig.add_subplot(gs[0, 1])

        # 柱状图
        current_data = self.dataset[self.current_semester]
        subjects = current_data['subjects']
        scores = [current_data['scores'][sub] for sub in subjects]

        ax1.bar(subjects, scores, color='#4C72B0')
        ax1.set_title(f'{self.current_semester}成绩分析')

        # 自动调整Y轴最大值为最大满分
        max_mark = max([self.full_marks.get(sub, 100) for sub in subjects])
        ax1.set_ylim(0, max_mark * 1.1)

        # 统计信息
        avg_score = sum(scores) / len(scores)
        stats_text = f'''统计指标：
        平均分：{avg_score:.1f}
        最高分：{max(scores)}
        最低分：{min(scores)}
        学科数量：{len(subjects)}'''
        ax1.text(1.05, 0.5, stats_text, transform=ax1.transAxes, va='center')

        # 饼图
        levels = self.calculate_levels(scores, subjects)  # 传入学科列表
        ax2.pie(levels.values(), labels=levels.keys(),
                autopct='%1.1f%%', colors=['#55A868', '#4C72B0', '#C44E52', '#8172B2'])
        ax2.set_title('成绩等级分布')

        self.display_chart(fig)

    def show_semester_analysis(self):
        """显示学期分析"""
        if not self.current_semester or not self.dataset[self.current_semester]['scores']:
            messagebox.showwarning("警告", "当前学期无成绩数据！")
            return

        mpl.rcParams['font.sans-serif'] = ['SimHei']
        mpl.rcParams['axes.unicode_minus'] = False

        # 获取所有学科
        all_subjects = set()
        for sem in self.dataset.values():
            all_subjects.update(sem['subjects'])
        all_subjects = list(all_subjects)

        # 学科选择对话框
        selected_subjects = self.select_subjects_for_trend(all_subjects)
        if not selected_subjects:
            return

        # 创建图表
        fig = plt.figure(figsize=(10, 5))
        gs = fig.add_gridspec(1, 2)
        ax1 = fig.add_subplot(gs[0, 0])
        ax2 = fig.add_subplot(gs[0, 1])

        # 绘制趋势线
        semesters = sorted(self.dataset.keys())
        for subject in selected_subjects:
            scores = []
            valid_semesters = []
            for sem in semesters:
                if subject in self.dataset[sem]['scores']:
                    scores.append(self.dataset[sem]['scores'][subject])
                    valid_semesters.append(sem)
            if scores:
                ax1.plot(valid_semesters, scores, marker='o', label=subject)

        ax1.set_title('学科成绩趋势分析')
        ax1.set_ylabel('分数')
        ax1.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()

        self.display_chart(fig)

    def show_trend_analysis(self):
        """显示趋势分析"""
        if not self.current_semester or not self.dataset[self.current_semester]['scores']:
            messagebox.showwarning("警告", "当前学期无成绩数据！")
            return

        # 获取所有学科
        all_subjects = set()
        for sem in self.dataset.values():
            all_subjects.update(sem['subjects'])
        all_subjects = list(all_subjects)

        # 学科选择对话框
        selected_subjects = self.select_subjects_for_trend(all_subjects)
        if not selected_subjects:
            return

        # 创建图表
        fig = plt.figure(figsize=(10, 5))
        gs = fig.add_gridspec(1, 2)
        ax = fig.add_subplot(gs[:, :])  # 使用整个区域绘制趋势图

        # 绘制趋势线
        semesters = sorted(self.dataset.keys())
        for subject in selected_subjects:
            scores = []
            valid_semesters = []
            for sem in semesters:
                if subject in self.dataset[sem]['scores']:
                    scores.append(self.dataset[sem]['scores'][subject])
                    valid_semesters.append(sem)
            if scores:
                ax.plot(valid_semesters, scores, marker='o', label=subject)

        ax.set_title('学科成绩趋势分析')
        ax.set_ylabel('分数')
        ax.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()

        self.display_chart(fig)

    # === 辅助功能 ===
    def update_data_table(self):
        """更新成绩表格"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.current_semester and self.dataset[self.current_semester]['scores']:
            current_data = self.dataset[self.current_semester]
            for subject in current_data['subjects']:
                score = current_data['scores'][subject]
                full_mark = self.full_marks.get(subject, 100)
                level = self.get_score_level(score, subject)  # 传入学科名称
                tags = ('warning',) if level == '不及格' else ()
                self.tree.insert("", "end", values=(subject, score, full_mark, level), tags=tags)

    def get_score_level(self, score, subject):
        """获取成绩等级（基于学科满分百分比）"""
        full_mark = self.full_marks.get(subject, 100)
        excellent = full_mark * 0.9
        good = full_mark * 0.8
        passing = full_mark * 0.6

        if score >= excellent:
            return '优秀'
        elif score >= good:
            return '良好'
        elif score >= passing:
            return '及格'
        else:
            return '不及格'

    def customize_subjects(self):
        """自定义学科"""
        dialog = tk.Toplevel()
        dialog.title("自定义学科")
        dialog.geometry("300x150")

        ttk.Label(dialog, text="年级:").grid(row=0, column=0, padx=5, pady=5)
        grade_combo = ttk.Combobox(dialog, values=list(self.grade_subjects.keys()), state="readonly")
        grade_combo.grid(row=0, column=1, padx=5, pady=5)
        grade_combo.current(0)

        ttk.Label(dialog, text="新学科名称:").grid(row=1, column=0, padx=5, pady=5)
        new_subject_entry = ttk.Entry(dialog)
        new_subject_entry.grid(row=1, column=1, padx=5, pady=5)

        def add_subject():
            grade = grade_combo.get()
            new_sub = new_subject_entry.get().strip()
            if not new_sub:
                messagebox.showwarning("警告", "请输入学科名称！")
                return

            if new_sub in self.grade_subjects[grade]:
                messagebox.showwarning("警告", "该学科已存在！")
                return

            self.grade_subjects[grade].append(new_sub)
            self.custom_subjects.setdefault(grade, []).append(new_sub)
            self.update_grade_subjects()
            dialog.destroy()
            messagebox.showinfo("成功", f"已为{grade}添加新学科: {new_sub}")

        ttk.Button(dialog, text="添加", command=add_subject).grid(row=2, columnspan=2, pady=10)

    def set_full_marks(self):
        """设置学科满分"""
        dialog = tk.Toplevel()
        dialog.title("设置满分")
        dialog.geometry("300x150")

        ttk.Label(dialog, text="学科:").grid(row=0, column=0, padx=5, pady=5)
        subject_combo = ttk.Combobox(dialog, values=self.get_all_subjects(), state="readonly")
        subject_combo.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(dialog, text="满分:").grid(row=1, column=0, padx=5, pady=5)
        mark_entry = ttk.Entry(dialog)
        mark_entry.grid(row=1, column=1, padx=5, pady=5)

        def save_mark():
            subject = subject_combo.get()
            mark = mark_entry.get().strip()

            if not subject:
                messagebox.showwarning("警告", "请选择学科！")
                return

            if not mark.isdigit():
                messagebox.showwarning("警告", "请输入有效数字！")
                return

            self.full_marks[subject] = int(mark)
            dialog.destroy()
            messagebox.showinfo("成功", f"{subject}满分已设置为{mark}")
            self.update_data_table()

        ttk.Button(dialog, text="保存", command=save_mark).grid(row=2, columnspan=2, pady=10)

    def get_all_subjects(self):
        """获取所有学科（包括自定义）"""
        subjects = set()
        for grade in self.grade_subjects.values():
            subjects.update(grade)
        for custom in self.custom_subjects.values():
            subjects.update(custom)
        return sorted(subjects)

    def select_subjects_for_trend(self, subjects):
        """选择趋势分析学科"""
        dialog = tk.Toplevel()
        dialog.title("选择分析学科")
        dialog.geometry("250x300")

        selected = []
        check_vars = {}

        for idx, sub in enumerate(subjects):
            var = tk.BooleanVar()
            check = ttk.Checkbutton(dialog, text=sub, variable=var)
            check.grid(row=idx, column=0, sticky="w", padx=10, pady=2)
            check_vars[sub] = var

        def confirm():
            nonlocal selected
            selected = [sub for sub, var in check_vars.items() if var.get()]
            dialog.destroy()

        ttk.Button(dialog, text="确定", command=confirm).grid(row=len(subjects) + 1, column=0, pady=10)
        dialog.wait_window()
        return selected

    def display_chart(self, fig):
        """显示图表"""
        if hasattr(self, 'canvas'):
            self.canvas.get_tk_widget().destroy()
        self.canvas = FigureCanvasTkAgg(fig, self.result_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def export_chart(self):
        """导出图表"""
        if not hasattr(self, 'canvas'):
            messagebox.showwarning("警告", "请先生成图表！")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG图片", "*.png"), ("PDF文档", "*.pdf"), ("SVG矢量图", "*.svg")])
        if filepath:
            try:
                self.canvas.figure.savefig(filepath, dpi=300, bbox_inches='tight')
                messagebox.showinfo("成功", "图表导出成功！")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{str(e)}")

    def load_chinese_font(self):
        font_path = os.path.join(os.getcwd(), "simhei.ttf")  # 字体文件路径
        if not os.path.exists(font_path):
            raise FileNotFoundError(f"字体文件未找到：{font_path}")
        pdfmetrics.registerFont(TTFont("SimHei", font_path))

    def generate_report(self):
        self.load_chinese_font()
        """生成PDF报告"""
        if not self.current_semester:
            messagebox.showwarning("警告", "请先选择学期！")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF文档", "*.pdf")])
        if not filepath:
            return

        try:

            # 创建PDF文档
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4

            # 标题
            c.setFont('SimHei', 16)
            c.drawString(50, height - 50, f"{self.current_semester}成绩分析报告")

            # 基本信息
            c.setFont('SimHei', 12)
            y = height - 100
            c.drawString(50, y, f"年级：{self.dataset[self.current_semester]['grade']}")
            y -= 30
            c.drawString(50, y, f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            y -= 50

            # 数据表格
            data = [["学科", "分数", "满分", "等级"]]
            current_data = self.dataset[self.current_semester]
            for subj in current_data['subjects']:
                score = current_data['scores'][subj]
                full = self.full_marks.get(subj, 100)
                level = self.get_score_level(score, current_data['grade'])
                data.append([subj, str(score), str(full), level])

            table = Table(data, colWidths=[100, 60, 60, 60])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4C72B0')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'SimHei'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F3F6FA')),
                ('GRID', (0, 0), (-1, -1), 1, colors.grey)
            ]))

            table.wrapOn(c, width - 100, height)
            table.drawOn(c, 50, y - 150)

            # 保存PDF
            c.showPage()
            c.save()
            messagebox.showinfo("成功", "成绩报告已生成！")
        except Exception as e:
            messagebox.showerror("错误", f"报告生成失败：{str(e)}")

    def calculate_levels(self, scores, subjects):
        """计算等级分布（基于各科满分）"""
        levels = {'优秀': 0, '良好': 0, '及格': 0, '不及格': 0}
        for subject, score in zip(subjects, scores):
            full_mark = self.full_marks.get(subject, 100)
            excellent = full_mark * 0.9
            good = full_mark * 0.8
            passing = full_mark * 0.6

            if score >= excellent:
                levels['优秀'] += 1
            elif score >= good:
                levels['良好'] += 1
            elif score >= passing:
                levels['及格'] += 1
            else:
                levels['不及格'] += 1
        return levels


if __name__ == "__main__":
    root = tk.Tk()
    app = EnhancedScoreAnalyzer(root)
    root.mainloop()