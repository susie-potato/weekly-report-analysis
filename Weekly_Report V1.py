# requirements.txt 文件内容（需要单独创建）:
# pandas==1.5.3
# matplotlib==3.7.0
# streamlit==1.28.0
# openpyxl==3.1.2
# plotly==5.15.0

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
import io
import warnings
warnings.filterwarnings('ignore')

# 设置页面
st.set_page_config(
    page_title="每周数据汇报分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 设置中文字体和颜色主题
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 主色调 - 浅紫色
MAIN_COLOR = '#B19CD9'
SECONDARY_COLORS = ['#C9B6E4', '#D8C7F0', '#A291C0', '#8A7BA8']

class WeeklyReportAnalyzer:
    def __init__(self, data):
        """初始化分析器"""
        self.df = data.copy()
        self.process_data()
        
    def process_data(self):
        """处理数据"""
        # 过滤掉汇总行
        self.df = self.df[self.df['审核组员'].notna()]
        self.df = self.df[self.df['审核组员'] != '']
        
        # 转换数据类型
        numeric_columns = ['审核量', '审核时间', '周扣分', '数量加分', '图文错误量', 'video错误量', '工作日']
        for col in numeric_columns:
            self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
        
        # 提取时间信息
        self.df['周期'] = self.df['时间']
        
        # 获取所有周期
        self.periods = sorted(self.df['周期'].unique())
        
    def get_periods_by_range(self, start_date, end_date):
        """根据日期范围获取周期"""
        selected_periods = []
        for period in self.periods:
            period_start = period.split('-')[0].strip()
            if start_date <= period_start <= end_date:
                selected_periods.append(period)
        
        return selected_periods
    
    def plot_weekly_deduction_trend(self, periods=None, title_suffix=""):
        """绘制周扣分和日扣分双折线图"""
        if periods is None:
            periods = self.periods
        
        period_data = self.df[self.df['周期'].isin(periods)]
        weekly_avg = period_data.groupby('周期').agg({
            '周扣分': 'mean',
            '日均扣分': 'mean'
        }).reset_index()
        
        fig, ax1 = plt.subplots(figsize=(10, 5))
        
        # 周扣分折线
        line1 = ax1.plot(weekly_avg['周期'], weekly_avg['周扣分'], 
                        marker='o', linewidth=2, markersize=6, 
                        color=MAIN_COLOR, label='周扣分')
        ax1.set_ylabel('周扣分', color=MAIN_COLOR, fontsize=12)
        ax1.tick_params(axis='y', labelcolor=MAIN_COLOR)
        ax1.set_ylim(bottom=0)
        
        # 日扣分折线（次坐标轴）
        ax2 = ax1.twinx()
        line2 = ax2.plot(weekly_avg['周期'], weekly_avg['日均扣分'], 
                        marker='s', linewidth=2, markersize=6, 
                        color=SECONDARY_COLORS[1], label='日均扣分')
        ax2.set_ylabel('日均扣分', color=SECONDARY_COLORS[1], fontsize=12)
        ax2.tick_params(axis='y', labelcolor=SECONDARY_COLORS[1])
        ax2.set_ylim(bottom=0)
        
        # 合并图例
        lines = line1 + line2
        labels = [l.get_label() for l in lines]
        ax1.legend(lines, labels, loc='upper left')
        
        plt.title(f'周扣分与日均扣分趋势图 {title_suffix}', fontsize=14, pad=20)
        plt.xticks(rotation=45)
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        
        return fig
    
    def plot_member_deduction_comparison(self, periods=None, title_suffix=""):
        """绘制成员扣分簇状图"""
        if periods is None:
            periods = self.periods[-4:]  # 默认最近4周
        
        period_data = self.df[self.df['周期'].isin(periods)]
        
        # 计算每个成员的总扣分用于排序
        member_total_deduction = period_data.groupby('审核组员')['周扣分'].sum().sort_values()
        
        # 创建数据透视表
        pivot_data = period_data.pivot_table(
            index='审核组员', 
            columns='周期', 
            values='周扣分', 
            aggfunc='mean'
        ).fillna(0)
        
        # 按总扣分排序
        pivot_data = pivot_data.loc[member_total_deduction.index]
        
        # 绘制簇状图
        fig, ax = plt.subplots(figsize=(12, 6))
        
        x = np.arange(len(pivot_data))
        width = 0.8 / len(periods)
        
        for i, period in enumerate(periods):
            if period in pivot_data.columns:
                ax.bar(x + i * width, pivot_data[period], width, 
                      label=period, color=SECONDARY_COLORS[i % len(SECONDARY_COLORS)])
        
        ax.set_xlabel('审核组员', fontsize=12)
        ax.set_ylabel('周扣分', fontsize=12)
        ax.set_title(f'审核成员周扣分对比 {title_suffix}', fontsize=14, pad=20)
        ax.set_xticks(x + width * (len(periods) - 1) / 2)
        ax.set_xticklabels(pivot_data.index, rotation=45, ha='right')
        ax.legend()
        ax.grid(True, alpha=0.3, axis='y')
        
        plt.tight_layout()
        return fig
    
    def plot_weekly_error_analysis(self, periods=None, title_suffix=""):
        """绘制每周错误量柱状图"""
        if periods is None:
            periods = self.periods
        
        period_data = self.df[self.df['周期'].isin(periods)]
        weekly_errors = period_data.groupby('周期').agg({
            '图文错误量': 'sum',
            'video错误量': 'sum'
        }).reset_index()
        
        fig, ax = plt.subplots(figsize=(10, 5))
        
        x = np.arange(len(weekly_errors))
        width = 0.35
        
        ax.bar(x - width/2, weekly_errors['图文错误量'], width, 
               label='图文错误量', color=MAIN_COLOR, alpha=0.8)
        ax.bar(x + width/2, weekly_errors['video错误量'], width, 
               label='video错误量', color=SECONDARY_COLORS[1], alpha=0.8)
        
        ax.set_xlabel('周期', fontsize=12)
        ax.set_ylabel('错误数量', fontsize=12)
        ax.set_title(f'每周错误量分析 {title_suffix}', fontsize=14, pad=20)
        ax.set_xticks(x)
        ax.set_xticklabels(weekly_errors['周期'], rotation=45)
        ax.legend()
        ax.grid(True, alpha=0.3, axis='y')
        
        plt.tight_layout()
        return fig
    
    def plot_member_bonus_comparison(self, periods=None, title_suffix=""):
        """绘制成员数量加分簇状图"""
        if periods is None:
            periods = self.periods[-4:]  # 默认最近4周
        
        period_data = self.df[self.df['周期'].isin(periods)]
        
        # 计算每个成员的总加分用于排序
        member_total_bonus = period_data.groupby('审核组员')['数量加分'].sum().sort_values(ascending=False)
        
        # 创建数据透视表
        pivot_data = period_data.pivot_table(
            index='审核组员', 
            columns='周期', 
            values='数量加分', 
            aggfunc='mean'
        ).fillna(0)
        
        # 按总加分排序
        pivot_data = pivot_data.loc[member_total_bonus.index]
        
        # 绘制簇状图
        fig, ax = plt.subplots(figsize=(12, 6))
        
        x = np.arange(len(pivot_data))
        width = 0.8 / len(periods)
        
        for i, period in enumerate(periods):
            if period in pivot_data.columns:
                ax.bar(x + i * width, pivot_data[period], width, 
                      label=period, color=SECONDARY_COLORS[i % len(SECONDARY_COLORS)])
        
        ax.set_xlabel('审核组员', fontsize=12)
        ax.set_ylabel('数量加分', fontsize=12)
        ax.set_title(f'审核成员数量加分对比 {title_suffix}', fontsize=14, pad=20)
        ax.set_xticks(x + width * (len(periods) - 1) / 2)
        ax.set_xticklabels(pivot_data.index, rotation=45, ha='right')
        ax.legend()
        ax.grid(True, alpha=0.3, axis='y')
        
        plt.tight_layout()
        return fig
    
    def generate_weekly_template(self, target_period):
        """生成固定模板表格"""
        if target_period not in self.periods:
            return None
        
        # 获取目标周期数据
        target_data = self.df[self.df['周期'] == target_period].copy()
        
        # 获取上一周期数据用于对比
        current_idx = self.periods.index(target_period)
        prev_period = self.periods[current_idx - 1] if current_idx > 0 else None
        
        if prev_period:
            prev_data = self.df[self.df['周期'] == prev_period].set_index('审核组员')
        else:
            prev_data = None
        
        # 按周扣分排序
        target_data = target_data.sort_values('周扣分')
        
        # 创建对比数据
        result_data = []
        for _, row in target_data.iterrows():
            member = row['审核组员']
            
            # 对比箭头
            arrow_review = ""
            arrow_deduction = ""
            
            if prev_data is not None and member in prev_data.index:
                # 审核量对比
                prev_review = prev_data.loc[member, '审核量']
                curr_review = row['审核量']
                if curr_review > prev_review:
                    arrow_review = "↑"
                elif curr_review < prev_review:
                    arrow_review = "↓"
                
                # 扣分对比
                prev_deduction = prev_data.loc[member, '周扣分']
                curr_deduction = row['周扣分']
                if curr_deduction < prev_deduction:
                    arrow_deduction = "↓"
                elif curr_deduction > prev_deduction:
                    arrow_deduction = "↑"
            
            result_data.append({
                '审核组员': member,
                '审核量': f"{row['审核量']}{arrow_review}",
                '审核时间': f"{row['审核时间']:.1f}",
                '日均审核量': f"{row['日均审核量']:.1f}",
                '周扣分': f"{row['周扣分']:.1f}{arrow_deduction}",
                '日均扣分': f"{row['日均扣分']:.3f}",
                '数量加分': row['数量加分'],
                '表现评级': row['表现评级'],
                '备注': row['备注'] if pd.notna(row['备注']) else ""
            })
        
        # 计算平均值
        avg_daily_review = target_data['日均审核量'].mean()
        avg_weekly_deduction = target_data['周扣分'].mean()
        avg_daily_deduction = target_data['日均扣分'].mean()
        
        return pd.DataFrame(result_data), avg_daily_review, avg_weekly_deduction, avg_daily_deduction

def main():
    # 标题和说明
    st.title("📊 每周数据汇报分析系统")
    st.markdown("---")
    
    # 文件上传
    uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx'], help="请上传'每周汇报.xlsx'文件")
    
    if uploaded_file is not None:
        try:
            # 读取数据
            df = pd.read_excel(uploaded_file, sheet_name='Sheet1')
            
            # 创建分析器
            analyzer = WeeklyReportAnalyzer(df)
            
            st.success(f"数据加载成功！共找到 {len(analyzer.periods)} 个周期")
            st.write(f"**可用周期**: {', '.join(analyzer.periods)}")
            
            # 侧边栏导航
            st.sidebar.title("导航")
            analysis_type = st.sidebar.selectbox(
                "选择分析类型",
                ["周扣分趋势", "成员扣分对比", "错误量分析", "成员加分对比", "周报模板"]
            )
            
            st.markdown("---")
            
            if analysis_type == "周扣分趋势":
                st.header("📈 周扣分趋势分析")
                
                col1, col2 = st.columns([1, 3])
                
                with col1:
                    period_option = st.radio(
                        "选择周期范围",
                        ["全部周期", "最近4周", "指定周期"]
                    )
                    
                    if period_option == "指定周期":
                        start_date = st.text_input("起始日期", "9.13")
                        end_date = st.text_input("结束日期", "10.25")
                
                with col2:
                    if period_option == "全部周期":
                        periods = analyzer.periods
                        title_suffix = "(全部周期)"
                    elif period_option == "最近4周":
                        periods = analyzer.periods[-4:]
                        title_suffix = "(最近4周)"
                    else:
                        periods = analyzer.get_periods_by_range(start_date, end_date)
                        title_suffix = f"({start_date}-{end_date})"
                    
                    if periods:
                        fig = analyzer.plot_weekly_deduction_trend(periods, title_suffix)
                        st.pyplot(fig)
                    else:
                        st.warning("未找到匹配的周期")
            
            elif analysis_type == "成员扣分对比":
                st.header("👥 审核成员扣分对比")
                
                col1, col2 = st.columns([1, 3])
                
                with col1:
                    period_option = st.radio(
                        "选择周期范围",
                        ["最近8周", "最近4周", "全部周期", "指定周期"]
                    )
                    
                    if period_option == "指定周期":
                        start_date = st.text_input("起始日期", "9.13", key="deduction_start")
                        end_date = st.text_input("结束日期", "10.25", key="deduction_end")
                
                with col2:
                    if period_option == "最近8周":
                        periods = analyzer.periods[-8:] if len(analyzer.periods) >= 8 else analyzer.periods
                        title_suffix = "(最近8周)"
                    elif period_option == "最近4周":
                        periods = analyzer.periods[-4:]
                        title_suffix = "(最近4周)"
                    elif period_option == "全部周期":
                        periods = analyzer.periods
                        title_suffix = "(全部周期)"
                    else:
                        periods = analyzer.get_periods_by_range(start_date, end_date)
                        title_suffix = f"({start_date}-{end_date})"
                    
                    if periods:
                        fig = analyzer.plot_member_deduction_comparison(periods, title_suffix)
                        st.pyplot(fig)
                    else:
                        st.warning("未找到匹配的周期")
            
            elif analysis_type == "错误量分析":
                st.header("❌ 每周错误量分析")
                
                col1, col2 = st.columns([1, 3])
                
                with col1:
                    period_option = st.radio(
                        "选择周期范围",
                        ["最近8周", "最近4周", "全部周期", "指定周期"],
                        key="error_option"
                    )
                    
                    if period_option == "指定周期":
                        start_date = st.text_input("起始日期", "9.13", key="error_start")
                        end_date = st.text_input("结束日期", "10.25", key="error_end")
                
                with col2:
                    if period_option == "最近8周":
                        periods = analyzer.periods[-8:] if len(analyzer.periods) >= 8 else analyzer.periods
                        title_suffix = "(最近8周)"
                    elif period_option == "最近4周":
                        periods = analyzer.periods[-4:]
                        title_suffix = "(最近4周)"
                    elif period_option == "全部周期":
                        periods = analyzer.periods
                        title_suffix = "(全部周期)"
                    else:
                        periods = analyzer.get_periods_by_range(start_date, end_date)
                        title_suffix = f"({start_date}-{end_date})"
                    
                    if periods:
                        fig = analyzer.plot_weekly_error_analysis(periods, title_suffix)
                        st.pyplot(fig)
                    else:
                        st.warning("未找到匹配的周期")
            
            elif analysis_type == "成员加分对比":
                st.header("⭐ 审核成员加分对比")
                
                col1, col2 = st.columns([1, 3])
                
                with col1:
                    period_option = st.radio(
                        "选择周期范围",
                        ["最近8周", "最近4周", "全部周期", "指定周期"],
                        key="bonus_option"
                    )
                    
                    if period_option == "指定周期":
                        start_date = st.text_input("起始日期", "9.13", key="bonus_start")
                        end_date = st.text_input("结束日期", "10.25", key="bonus_end")
                
                with col2:
                    if period_option == "最近8周":
                        periods = analyzer.periods[-8:] if len(analyzer.periods) >= 8 else analyzer.periods
                        title_suffix = "(最近8周)"
                    elif period_option == "最近4周":
                        periods = analyzer.periods[-4:]
                        title_suffix = "(最近4周)"
                    elif period_option == "全部周期":
                        periods = analyzer.periods
                        title_suffix = "(全部周期)"
                    else:
                        periods = analyzer.get_periods_by_range(start_date, end_date)
                        title_suffix = f"({start_date}-{end_date})"
                    
                    if periods:
                        fig = analyzer.plot_member_bonus_comparison(periods, title_suffix)
                        st.pyplot(fig)
                    else:
                        st.warning("未找到匹配的周期")
            
            elif analysis_type == "周报模板":
                st.header("📋 周报模板生成")
                
                col1, col2 = st.columns([1, 3])
                
                with col1:
                    target_period = st.selectbox("选择目标周期", analyzer.periods)
                    generate_btn = st.button("生成周报模板")
                
                with col2:
                    if generate_btn:
                        result = analyzer.generate_weekly_template(target_period)
                        if result:
                            template_df, avg_review, avg_weekly_ded, avg_daily_ded = result
                            
                            st.subheader(f"周报模板 - {target_period}")
                            
                            # 显示统计信息
                            st.metric("日均审核量", f"{avg_review:.1f}")
                            st.metric("周均扣分", f"{avg_weekly_ded:.2f}")
                            st.metric("日均扣分", f"{avg_daily_ded:.3f}")
                            
                            # 显示数据表格
                            st.dataframe(template_df, use_container_width=True)
                            
                            st.info("↑ 表示相比上一周期上升，↓ 表示相比上一周期下降")
                        else:
                            st.error("生成周报模板失败")
        
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
    
    else:
        st.info("👆 请上传Excel文件开始分析")
        
        # 显示使用说明
        with st.expander("使用说明"):
            st.markdown("""
            ### 📖 使用说明
            
            1. **上传文件**: 点击"上传Excel文件"按钮，选择您的`每周汇报.xlsx`文件
            2. **选择分析类型**: 在左侧边栏选择您想要的分析类型
            3. **设置参数**: 根据需要调整周期范围等参数
            4. **查看结果**: 系统会自动生成相应的图表和报告
            
            ### 📊 分析功能
            
            - **周扣分趋势**: 查看周扣分和日均扣分的变化趋势
            - **成员扣分对比**: 比较不同成员的扣分情况
            - **错误量分析**: 分析图文和视频错误的数量
            - **成员加分对比**: 比较不同成员的加分情况
            - **周报模板**: 生成指定周期的详细报告
            
            ### 🎨 颜色说明
            
            - 主色调: 浅紫色
            - 对比箭头: ↑表示上升，↓表示下降
            """)

if __name__ == "__main__":
    main()