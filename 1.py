import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class ScoreAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("成绩分析工具")

        # 初始化数据存储字典
        self.scores = {}

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        """创建界面组件"""
        # 输入框架
        input_frame = ttk.LabelFrame(self.root, text="成绩输入")
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # 学科名称输入
        ttk.Label(input_frame, text="学科名称：").grid(row=0, column=0)
        self.subject_entry = ttk.Entry(input_frame)
        self.subject_entry.grid(row=0, column=1, padx=5)

        # 分数输入
        ttk.Label(input_frame, text="分数：").grid(row=1, column=0)
        self.score_entry = ttk.Entry(input_frame)
        self.score_entry.grid(row=1, column=1, padx=5)

        # 添加按钮
        add_btn = ttk.Button(input_frame, text="添加成绩", command=self.add_score)
        add_btn.grid(row=2, column=0, columnspan=2, pady=5)

        # 分析按钮
        analyze_btn = ttk.Button(self.root, text="生成分析", command=self.analyze_scores)
        analyze_btn.grid(row=1, column=0, pady=10)

        # 结果显示框架
        self.result_frame = ttk.Frame(self.root)
        self.result_frame.grid(row=2, column=0, padx=10, pady=10)

    def add_score(self):
        """添加成绩到字典"""
        subject = self.subject_entry.get()
        score = self.score_entry.get()

        # 输入验证
        if not subject or not score:
            return

        try:
            self.scores[subject] = float(score)
            # 清空输入框
            self.subject_entry.delete(0, tk.END)
            self.score_entry.delete(0, tk.END)
        except ValueError:
            pass

    def analyze_scores(self):
        """生成成绩分析图表"""
        # 清除旧图表
        for widget in self.result_frame.winfo_children():
            widget.destroy()

        # 创建Matplotlib图形
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4))

        # 柱状图
        subjects = list(self.scores.keys())
        scores = list(self.scores.values())
        ax1.bar(subjects, scores, color='skyblue')
        ax1.set_title('各科成绩对比')
        ax1.set_ylim(0, 100)

        # 饼状图（等级分布）
        levels = {'优秀': 0, '良好': 0, '及格': 0, '不及格': 0}
        for score in scores:
            if score >= 85:
                levels['优秀'] += 1
            elif score >= 70:
                levels['良好'] += 1
            elif score >= 60:
                levels['及格'] += 1
            else:
                levels['不及格'] += 1

        ax2.pie(levels.values(), labels=levels.keys(),
                autopct='%1.1f%%', colors=['#66b3ff', '#99ff99', '#ffcc99', '#ff9999'])
        ax2.set_title('成绩等级分布')

        # 将图表嵌入Tkinter窗口
        canvas = FigureCanvasTkAgg(fig, master=self.result_frame)
        canvas.draw()
        canvas.get_tk_widget().pack()


# 创建主窗口并运行
root = tk.Tk()
app = ScoreAnalyzer(root)
root.mainloop()