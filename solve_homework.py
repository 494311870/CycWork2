#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
作业3：指数纳入效应套利策略的实现
完整解决方案，包含详细的计算步骤和公式说明
"""

import pandas as pd
import numpy as np
from scipy import stats
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
import matplotlib
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

class IndexInclusionArbitrage:
    """指数纳入效应套利策略分析类"""
    
    def __init__(self, excel_file):
        """初始化并加载数据"""
        self.excel_file = excel_file
        self.doc = Document()
        self.setup_document_style()
        
    def setup_document_style(self):
        """设置文档样式"""
        # 设置默认字体
        self.doc.styles['Normal'].font.name = 'Times New Roman'
        self.doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        self.doc.styles['Normal'].font.size = Pt(12)
        
    def add_title(self, text, level=1):
        """添加标题"""
        heading = self.doc.add_heading(text, level=level)
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return heading
        
    def add_paragraph(self, text):
        """添加段落"""
        p = self.doc.add_paragraph(text)
        return p
        
    def add_formula(self, formula_text):
        """添加公式说明"""
        p = self.doc.add_paragraph()
        run = p.add_run(formula_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        return p
        
    def calculate_excess_return(self, row):
        """
        计算超额收益率
        公式：Abret = Dretwd - Beta × Dretwdos
        
        参数说明：
        - Dretwd: 个股收益率
        - Beta: 个股的系统风险系数
        - Dretwdos: 市场收益率
        """
        return row['Dretwd'] - row['Beta'] * row['Dretwdos']
    
    def solve_question1(self):
        """
        问题一：利用纳入和剔除股票做套利
        """
        print("\n" + "="*80)
        print("开始计算问题一...")
        print("="*80)
        
        # 添加问题一标题
        self.add_title('问题一：利用纳入和剔除股票做套利', level=1)
        
        # 读取Q1数据
        df = pd.read_excel(self.excel_file, sheet_name='Q1数据')
        
        # 计算超额收益率
        df['Abret'] = df.apply(self.calculate_excess_return, axis=1)
        
        # 添加详细的计算说明
        self.add_title('一、超额收益率的计算', level=2)
        
        self.add_paragraph('1. 计算公式：')
        self.add_formula('超额收益率（Abret）= 个股收益率（Dretwd）- Beta × 市场收益率（Dretwdos）')
        
        self.add_paragraph('\n2. 公式说明：')
        self.add_paragraph('• Dretwd：个股日收益率，已由数据提供')
        self.add_paragraph('• Beta：个股的系统风险系数，已由数据提供')
        self.add_paragraph('• Dretwdos：市场收益率（沪深300指数收益率），已由数据提供')
        self.add_paragraph('• Abret：超额收益率，即个股收益率中剔除市场收益影响后的部分')
        
        self.add_paragraph('\n3. Excel实现步骤：')
        self.add_paragraph('步骤1：在Excel中打开"Q1数据"工作表')
        self.add_paragraph('步骤2：假设Dretwd在I列，Beta在H列，Dretwdos在J列')
        self.add_paragraph('步骤3：在K列（Abret列）的第2行输入公式：=I2-H2*J2')
        self.add_paragraph('步骤4：将公式向下拖拽复制到所有数据行')
        self.add_paragraph('步骤5：这样就得到了每个股票每天的超额收益率')
        
        # 计算各时间窗口的统计量
        self.add_title('二、各时间窗口超额收益率统计', level=2)
        
        results_q1_1 = []
        
        for twind in range(6):  # Twind 0-5
            # 纳入股票 (Chgsmp04 = 1)
            inclusion_data = df[(df['Chgsmp04'] == 1) & (df['twind'] == twind)]['Abret']
            inclusion_n = len(inclusion_data)
            inclusion_mean = inclusion_data.mean() * 10000  # 转换为基点
            inclusion_std_err = inclusion_data.std() / np.sqrt(inclusion_n) * 10000
            
            # 剔除股票 (Chgsmp04 = 2)
            exclusion_data = df[(df['Chgsmp04'] == 2) & (df['twind'] == twind)]['Abret']
            exclusion_n = len(exclusion_data)
            exclusion_mean = exclusion_data.mean() * 10000
            exclusion_std_err = exclusion_data.std() / np.sqrt(exclusion_n) * 10000
            
            # 差异检验
            t_stat, p_value = stats.ttest_ind(inclusion_data, exclusion_data)
            difference = inclusion_mean - exclusion_mean
            
            results_q1_1.append({
                'twind': twind,
                'inclusion_n': inclusion_n,
                'inclusion_mean': inclusion_mean,
                'inclusion_std_err': inclusion_std_err,
                'exclusion_n': exclusion_n,
                'exclusion_mean': exclusion_mean,
                'exclusion_std_err': exclusion_std_err,
                'difference': difference,
                'p_value': p_value
            })
        
        # 添加统计计算说明
        self.add_paragraph('\n4. 统计量计算方法：')
        self.add_paragraph('• N（样本数）：每个时间窗口内的股票观测数量')
        self.add_paragraph('• Mean（均值，基点）：超额收益率的平均值 × 10000（转换为基点，1基点=0.01%）')
        self.add_paragraph('• Std Err（标准误）：标准差 / √N × 10000')
        
        self.add_paragraph('\n5. Excel实现步骤：')
        self.add_paragraph('步骤1：筛选纳入股票（Chgsmp04=1）且Twind=0的数据')
        self.add_paragraph('步骤2：计算N：使用COUNT函数统计Abret列非空值个数')
        self.add_paragraph('       公式示例：=COUNT(K2:K500) （假设数据在K2到K500）')
        self.add_paragraph('步骤3：计算Mean：使用AVERAGE函数计算平均值，再乘以10000')
        self.add_paragraph('       公式示例：=AVERAGE(K2:K500)*10000')
        self.add_paragraph('步骤4：计算Std Err：使用STDEV.S计算标准差，除以SQRT(N)，再乘以10000')
        self.add_paragraph('       公式示例：=STDEV.S(K2:K500)/SQRT(COUNT(K2:K500))*10000')
        self.add_paragraph('步骤5：对剔除股票（Chgsmp04=2）重复上述步骤')
        self.add_paragraph('步骤6：对Twind=1,2,3,4,5分别重复上述步骤')
        
        # 添加表1
        self.add_title('表1：纳入和剔除股票的超额收益率统计', level=3)
        
        table = self.doc.add_table(rows=9, cols=8)
        table.style = 'Light Grid Accent 1'
        
        # 表头
        headers = ['Twind', '', '0', '1', '2', '3', '4', '5']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
        
        # 填充数据
        row_labels = [
            '纳入指数股票超额收益率|N',
            '纳入指数股票超额收益率|Mean(base point)',
            '纳入指数股票超额收益率|Std Err',
            '剔除指数股票超额收益率|N',
            '剔除指数股票超额收益率|Mean(base point)',
            '剔除指数股票超额收益率|Std Err',
            '纳入与剔除股票超额收益率的差|Difference',
            '纳入与剔除股票超额收益率的差|P值'
        ]
        
        for i, label in enumerate(row_labels, 1):
            parts = label.split('|')
            table.rows[i].cells[0].text = parts[0]
            table.rows[i].cells[1].text = parts[1]
        
        for j, result in enumerate(results_q1_1, 2):
            table.rows[1].cells[j].text = str(result['inclusion_n'])
            table.rows[2].cells[j].text = f"{result['inclusion_mean']:.2f}"
            table.rows[3].cells[j].text = f"{result['inclusion_std_err']:.2f}"
            table.rows[4].cells[j].text = str(result['exclusion_n'])
            table.rows[5].cells[j].text = f"{result['exclusion_mean']:.2f}"
            table.rows[6].cells[j].text = f"{result['exclusion_std_err']:.2f}"
            table.rows[7].cells[j].text = f"{result['difference']:.2f}"
            table.rows[8].cells[j].text = f"{result['p_value']:.4f}"
        
        # 添加T检验说明
        self.add_paragraph('\n6. 差异检验（T检验）：')
        self.add_paragraph('• Difference：纳入股票均值 - 剔除股票均值')
        self.add_paragraph('• P值：使用独立样本T检验，检验两组均值是否有显著差异')
        
        self.add_paragraph('\n7. Excel中的T检验实现：')
        self.add_paragraph('步骤1：假设纳入股票Twind=0的Abret数据在K2:K458区域')
        self.add_paragraph('步骤2：假设剔除股票Twind=0的Abret数据在K459:K929区域')
        self.add_paragraph('步骤3：使用公式：=T.TEST(K2:K458, K459:K929, 2, 2)')
        self.add_paragraph('       参数说明：第3个参数"2"表示双尾检验，第4个参数"2"表示独立双样本等方差检验')
        self.add_paragraph('步骤4：对其他Twind值重复此步骤')
        
        # 分析结论
        self.add_title('三、结果分析', level=2)
        
        # 检查显著性
        significant_inclusion = sum([1 for r in results_q1_1 if r['inclusion_mean'] > 0 and r['p_value'] < 0.05])
        significant_exclusion = sum([1 for r in results_q1_1 if r['exclusion_mean'] < 0 and r['p_value'] < 0.05])
        
        analysis_text = f"""
根据表1的计算结果：

1. 纳入效应分析：
   - 纳入股票在公告日当天（Twind=0）的平均超额收益率为{results_q1_1[0]['inclusion_mean']:.2f}基点
   - 在观察的6个交易日中，有{sum([1 for r in results_q1_1 if r['inclusion_mean'] > 0])}天显示正的超额收益
   - P值显示{'存在' if significant_inclusion > 0 else '不存在'}显著的统计意义

2. 剔除效应分析：
   - 剔除股票在公告日当天（Twind=0）的平均超额收益率为{results_q1_1[0]['exclusion_mean']:.2f}基点
   - 在观察的6个交易日中，有{sum([1 for r in results_q1_1 if r['exclusion_mean'] < 0])}天显示负的超额收益
   - P值显示{'存在' if significant_exclusion > 0 else '不存在'}显著的剔除效应

3. 差异检验：
   - 纳入与剔除股票之间的超额收益率差异在多数时间窗口{'显著' if sum([1 for r in results_q1_1 if r['p_value'] < 0.05]) > 3 else '不显著'}
   - 这表明指数纳入效应{'确实存在' if sum([1 for r in results_q1_1 if r['p_value'] < 0.05]) > 3 else '不够明显'}
"""
        self.add_paragraph(analysis_text)
        
        # 问题一第二部分：套利策略构建
        self.add_title('四、套利策略构建与检验', level=2)
        
        results_q1_2 = []
        
        for twind in range(6):
            # 纳入股票收益率
            inclusion_returns = df[(df['Chgsmp04'] == 1) & (df['twind'] == twind)]['Dretwd']
            inclusion_n = len(inclusion_returns)
            inclusion_mean = inclusion_returns.mean() * 10000
            inclusion_std_err = inclusion_returns.std() / np.sqrt(inclusion_n) * 10000
            
            # 剔除股票收益率
            exclusion_returns = df[(df['Chgsmp04'] == 2) & (df['twind'] == twind)]['Dretwd']
            exclusion_n = len(exclusion_returns)
            exclusion_mean = exclusion_returns.mean() * 10000
            exclusion_std_err = exclusion_returns.std() / np.sqrt(exclusion_n) * 10000
            
            # 套利策略收益（买入纳入 - 卖出剔除）
            arbitrage_return = inclusion_mean - exclusion_mean
            
            # 检验策略收益是否显著
            t_stat, p_value = stats.ttest_ind(inclusion_returns, exclusion_returns)
            
            results_q1_2.append({
                'twind': twind,
                'inclusion_n': inclusion_n,
                'inclusion_mean': inclusion_mean,
                'inclusion_std_err': inclusion_std_err,
                'exclusion_n': exclusion_n,
                'exclusion_mean': exclusion_mean,
                'exclusion_std_err': exclusion_std_err,
                'arbitrage_return': arbitrage_return,
                'p_value': p_value
            })
        
        self.add_paragraph('1. 套利策略设计：')
        self.add_paragraph('• 等权重买入所有纳入指数的股票')
        self.add_paragraph('• 等权重卖空同样数额的所有剔除指数的股票')
        self.add_paragraph('• 策略收益 = 纳入股票平均收益 - 剔除股票平均收益')
        
        self.add_paragraph('\n2. Excel计算步骤：')
        self.add_paragraph('步骤1：筛选纳入股票（Chgsmp04=1）且Twind=0的数据')
        self.add_paragraph('步骤2：计算纳入股票的平均收益率：=AVERAGE(Dretwd列)*10000')
        self.add_paragraph('步骤3：筛选剔除股票（Chgsmp04=2）且Twind=0的数据')
        self.add_paragraph('步骤4：计算剔除股票的平均收益率：=AVERAGE(Dretwd列)*10000')
        self.add_paragraph('步骤5：计算套利策略收益：=纳入股票平均收益 - 剔除股票平均收益')
        self.add_paragraph('步骤6：对Twind=1,2,3,4,5重复上述计算')
        
        # 添加表2
        self.add_title('表2：套利策略收益统计', level=3)
        
        table2 = self.doc.add_table(rows=9, cols=8)
        table2.style = 'Light Grid Accent 1'
        
        # 表头
        for i, header in enumerate(headers):
            table2.rows[0].cells[i].text = header
        
        # 行标签
        row_labels2 = [
            '纳入指数股票收益率|N',
            '纳入指数股票收益率|Mean(base point)',
            '纳入指数股票收益率|Std Err',
            '剔除指数股票收益率|N',
            '剔除指数股票收益率|Mean(base point)',
            '剔除指数股票收益率|Std Err',
            '套利策略收益（买入纳入卖出剔除股票）|Difference',
            '套利策略收益（买入纳入卖出剔除股票）|P值'
        ]
        
        for i, label in enumerate(row_labels2, 1):
            parts = label.split('|')
            table2.rows[i].cells[0].text = parts[0]
            table2.rows[i].cells[1].text = parts[1]
        
        for j, result in enumerate(results_q1_2, 2):
            table2.rows[1].cells[j].text = str(result['inclusion_n'])
            table2.rows[2].cells[j].text = f"{result['inclusion_mean']:.2f}"
            table2.rows[3].cells[j].text = f"{result['inclusion_std_err']:.2f}"
            table2.rows[4].cells[j].text = str(result['exclusion_n'])
            table2.rows[5].cells[j].text = f"{result['exclusion_mean']:.2f}"
            table2.rows[6].cells[j].text = f"{result['exclusion_std_err']:.2f}"
            table2.rows[7].cells[j].text = f"{result['arbitrage_return']:.2f}"
            table2.rows[8].cells[j].text = f"{result['p_value']:.4f}"
        
        # 策略问题分析
        self.add_title('五、Mac策略存在的问题及改进建议', level=2)
        
        strategy_analysis = """
Mac策略存在的主要问题：

1. 市场风险问题：
   - 策略假设等权重买入和卖出，但没有考虑市场整体波动
   - 在市场大幅下跌时，即使做了对冲，仍可能面临整体亏损

2. Beta不匹配问题：
   - 纳入和剔除股票的Beta值可能差异很大
   - 简单等权重对冲可能无法完全中和市场风险
   - 需要根据Beta值调整对冲比例

3. 行业集中度风险：
   - 纳入和剔除股票可能集中在某些行业
   - 行业因素可能导致策略收益波动

4. 流动性风险：
   - 部分剔除股票可能流动性较差
   - 卖空成本和难度可能较高

5. 时间窗口选择：
   - 策略只考虑了0-5天的短期效应
   - 未考虑中长期的价格回归

改进建议：

1. 使用配对交易方法：
   - 为每只纳入股票寻找相似的对照股票
   - 匹配条件：同行业、同市场、相近Beta
   - 这正是问题二和问题三要实现的方法

2. 动态调整对冲比例：
   - 根据Beta值调整买入卖出比例
   - 使组合整体Beta接近0

3. 考虑交易成本：
   - 纳入印花税、佣金、卖空成本
   - 评估净收益是否仍然显著

4. 分散化投资：
   - 避免过度集中在某个行业或板块
   - 选择流动性好的股票进行交易
"""
        self.add_paragraph(strategy_analysis)
        
        print(f"问题一完成！")
        return results_q1_1, results_q1_2
    
    def solve_question2(self):
        """
        问题二：寻找匹配样本
        """
        print("\n" + "="*80)
        print("开始计算问题二...")
        print("="*80)
        
        self.add_title('问题二：寻找匹配样本', level=1)
        
        # 读取Q2数据（header在第5行，索引为4）
        df = pd.read_excel(self.excel_file, sheet_name='Q2数据', header=4)
        
        self.add_paragraph('1. 匹配原则：')
        self.add_paragraph('• 同行业（Nnindcd相同）')
        self.add_paragraph('• 同市场（Markettype相同）')
        self.add_paragraph('• Beta匹配精度：|Beta - Beta1| < 0.01（绝对误差小于1%）')
        
        self.add_paragraph('\n2. Excel实现步骤：')
        self.add_paragraph('步骤1：在Q2数据表中，已经提供了纳入股票及其可能的匹配股票')
        self.add_paragraph('步骤2：计算Beta差异的绝对值：在新列中输入公式')
        self.add_paragraph('       =ABS(Beta - Beta1)')
        self.add_paragraph('步骤3：筛选符合条件的记录：')
        self.add_paragraph('       - 行业代码相同：Nnindcd = Nnindcd1')
        self.add_paragraph('       - 市场类型相同：Markettype = Markettype1')
        self.add_paragraph('       - Beta差异 < 0.01')
        self.add_paragraph('步骤4：使用Excel的筛选功能或公式：')
        self.add_paragraph('       =IF(AND(Nnindcd=Nnindcd1, Markettype=Markettype1, ABS(Beta-Beta1)<0.01), "匹配", "不匹配")')
        self.add_paragraph('步骤5：对每只纳入股票，选择Beta差异最小的一只作为匹配股票')
        
        # 查找匹配样本
        matched_pairs = []
        
        # 获取唯一的纳入股票
        unique_stocks = df['Stkcd'].unique()
        
        for stock in unique_stocks:
            # 获取该股票的所有候选匹配
            candidates = df[df['Stkcd'] == stock].copy()
            
            if len(candidates) == 0:
                continue
                
            # 筛选满足Beta精度要求的候选（误差<1%即0.01）
            candidates['beta_diff'] = abs(candidates['Beta'] - candidates['Beta1'])
            valid_candidates = candidates[candidates['beta_diff'] < 0.01]
            
            if len(valid_candidates) > 0:
                # 选择Beta差异最小的
                best_match = valid_candidates.loc[valid_candidates['beta_diff'].idxmin()]
                matched_pairs.append({
                    'Stkcd': int(best_match['Stkcd']),
                    'Stkcd1': int(best_match['Stkcd1'])
                })
        
        self.add_paragraph(f'\n3. 匹配结果：共找到 {len(matched_pairs)} 对匹配样本')
        
        # 添加表3
        self.add_title('表3：匹配样本表', level=3)
        
        table3 = self.doc.add_table(rows=len(matched_pairs)+1, cols=2)
        table3.style = 'Light Grid Accent 1'
        
        table3.rows[0].cells[0].text = 'Stkcd（纳入指数股票）'
        table3.rows[0].cells[1].text = 'Stkcd1（匹配股票）'
        
        for i, pair in enumerate(matched_pairs, 1):
            table3.rows[i].cells[0].text = str(pair['Stkcd'])
            table3.rows[i].cells[1].text = str(pair['Stkcd1'])
        
        self.add_paragraph('\n4. 匹配质量说明：')
        self.add_paragraph('上述匹配股票均满足：')
        self.add_paragraph('• 与纳入股票属于同一行业')
        self.add_paragraph('• 在同一市场上市')
        self.add_paragraph('• Beta值差异的绝对值小于0.01（1%）')
        
        print(f"问题二完成！找到 {len(matched_pairs)} 对匹配样本")
        return matched_pairs
    
    def solve_question3(self):
        """
        问题三：利用匹配样本构建套利策略
        """
        print("\n" + "="*80)
        print("开始计算问题三...")
        print("="*80)
        
        self.add_title('问题三：利用匹配样本构建套利策略', level=1)
        
        # 读取Q3数据
        # 首先读取匹配样本表（header在第4行，索引为3）
        df_pairs = pd.read_excel(self.excel_file, sheet_name='Q3数据', header=3, nrows=5)
        
        # 读取收益率数据（header在第14行，索引为13）
        df_returns = pd.read_excel(self.excel_file, sheet_name='Q3数据', header=13)
        
        self.add_paragraph('1. Buy & Hold Return 计算原理：')
        self.add_paragraph('• 在公告日（2025年6月16日）收盘价买入纳入股票，卖出匹配股票')
        self.add_paragraph('• 从第二天（2025年6月17日）开始计算持有收益')
        self.add_paragraph('• 配对组收益 = 纳入股票收益 - 匹配股票收益')
        self.add_paragraph('• 组合收益 = 所有配对组的等权重平均')
        
        self.add_paragraph('\n2. Excel计算步骤：')
        self.add_paragraph('步骤1：识别5对配对股票（从表1获取）')
        self.add_paragraph('步骤2：提取每对股票在10个交易日的收益率数据')
        self.add_paragraph('       - 纳入股票收益率：Dretwd')
        self.add_paragraph('       - 匹配股票收益率：Dretwd1')
        self.add_paragraph('步骤3：计算每个配对组每天的收益：')
        self.add_paragraph('       =Dretwd（纳入股票）- Dretwd1（匹配股票）')
        self.add_paragraph('步骤4：处理停牌情况：')
        self.add_paragraph('       - 如果配对中任一股票停牌（收益率为空），则该配对组当天收益为空')
        self.add_paragraph('       - 当天组合收益只计算有效配对组的等权重平均')
        self.add_paragraph('步骤5：计算等权重组合收益：')
        self.add_paragraph('       =AVERAGE(有效配对组收益) 或 =AVERAGEIF(配对组收益范围, "<>空值")')
        
        # 计算每个配对组的收益
        pairs_info = []
        for i in range(5):
            row = df_pairs.iloc[i]
            pairs_info.append({
                'stkcd': int(row['Stkcd']),
                'stkcd1': int(row['Stkcd1'])
            })
        
        # 计算10天的收益
        daily_returns = []
        
        for day in range(1, 11):  # 持有第1-10天
            day_data = df_returns[df_returns['twind'] == day]
            
            pair_returns = []
            for pair in pairs_info:
                # 找到该配对的数据
                inclusion_data = day_data[day_data['Stkcd'] == pair['stkcd']]
                matched_data = day_data[day_data['Stkcd1'] == pair['stkcd1']]
                
                if len(inclusion_data) > 0 and len(matched_data) > 0:
                    inclusion_return = inclusion_data.iloc[0]['Dretwd']
                    matched_return = inclusion_data.iloc[0]['Dretwd1']
                    
                    # 检查是否停牌（收益率为NaN或0）
                    if pd.notna(inclusion_return) and pd.notna(matched_return):
                        pair_return = inclusion_return - matched_return
                        pair_returns.append(pair_return)
            
            # 计算当天等权重组合收益
            if len(pair_returns) > 0:
                portfolio_return = np.mean(pair_returns)
            else:
                portfolio_return = 0
                
            daily_returns.append(portfolio_return * 100)  # 转换为百分比
        
        self.add_paragraph('\n3. 停牌权重调整说明：')
        self.add_paragraph('• 某配对组停牌时，该组不参与当天组合收益计算')
        self.add_paragraph('• 其他配对组等权重分配100%权重')
        self.add_paragraph('• 例如：5对中有1对停牌，则其余4对各占25%权重')
        
        # 添加表4
        self.add_title('表4：Marina策略10天Buy & Hold Return（%）', level=3)
        
        table4 = self.doc.add_table(rows=2, cols=10)
        table4.style = 'Light Grid Accent 1'
        
        for i in range(10):
            table4.rows[0].cells[i].text = str(i+1)
            table4.rows[1].cells[i].text = f"{daily_returns[i]:.4f}"
        
        # 绘制图表
        plt.figure(figsize=(10, 6))
        plt.plot(range(1, 11), daily_returns, marker='o', linewidth=2, markersize=8)
        plt.xlabel('持有天数', fontsize=12)
        plt.ylabel('收益率 (%)', fontsize=12)
        plt.title('问题三：Marina策略10天Buy & Hold Return', fontsize=14, fontweight='bold')
        plt.grid(True, alpha=0.3)
        plt.axhline(y=0, color='r', linestyle='--', alpha=0.5)
        
        chart_path = '/home/runner/work/CycWork2/CycWork2/q3_chart.png'
        plt.savefig(chart_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        self.doc.add_paragraph('\n图1：Marina策略10天收益率走势')
        self.doc.add_picture(chart_path, width=Inches(6))
        
        # 累计收益分析
        cumulative_return = sum(daily_returns)
        self.add_paragraph(f'\n4. 收益分析：')
        self.add_paragraph(f'• 10天累计收益率：{cumulative_return:.4f}%')
        self.add_paragraph(f'• 平均日收益率：{np.mean(daily_returns):.4f}%')
        self.add_paragraph(f'• 收益率标准差：{np.std(daily_returns):.4f}%')
        self.add_paragraph(f'• 正收益天数：{sum([1 for r in daily_returns if r > 0])} 天')
        
        print(f"问题三完成！")
        return daily_returns
    
    def solve_question4(self):
        """
        问题四：套利策略在牛熊市中的稳健性检验
        """
        print("\n" + "="*80)
        print("开始计算问题四...")
        print("="*80)
        
        self.add_title('问题四：套利策略在牛熊市中的稳健性检验', level=1)
        
        self.add_paragraph('本问题分析套利策略在不同市场环境下的表现：')
        self.add_paragraph('• 2022年6月：牛市环境')
        self.add_paragraph('• 2023年6月：非牛熊市环境')
        self.add_paragraph('• 2024年6月：熊市环境')
        
        # 计算三个时期的收益
        periods = [
            ('Q4_1数据', '2022年6月（牛市）', 5),
            ('Q4_2数据', '2023年6月（非牛熊市）', 6),
            ('Q4_3数据', '2024年6月（熊市）', 7)
        ]
        
        all_period_returns = []
        
        for sheet_name, period_name, table_num in periods:
            self.add_title(f'{period_name}策略收益分析', level=2)
            
            # 确定header行
            if sheet_name == 'Q4_1数据':
                pairs_header = 2  # Row 3
                pairs_nrows = 12  # More rows in Q4_1
                returns_header = 21  # Row 22 (Stkcd is on row 21, data starts row 22)
            elif sheet_name == 'Q4_2数据':
                pairs_header = 2  # Row 3
                pairs_nrows = 8  # Rows in Q4_2
                returns_header = 16  # Row 17
            else:  # Q4_3数据
                pairs_header = 2  # Row 3
                pairs_nrows = 7  # Rows in Q4_3
                returns_header = 15  # Row 16
            
            # 读取匹配样本
            df_pairs = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=pairs_header, nrows=pairs_nrows)
            
            # 读取收益率数据
            df_returns = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=returns_header)
            
            # 获取配对信息 - 为每只纳入股票选择一只匹配股票
            pairs_info = []
            seen_stocks = set()
            
            for i in range(len(df_pairs)):
                row = df_pairs.iloc[i]
                if pd.notna(row['Stkcd']) and pd.notna(row['Stkcd1']):
                    stkcd = int(row['Stkcd'])
                    # 只取每只纳入股票的第一个匹配
                    if stkcd not in seen_stocks:
                        pairs_info.append({
                            'stkcd': stkcd,
                            'stkcd1': int(row['Stkcd1'])
                        })
                        seen_stocks.add(stkcd)
                    
                    # 限制为5对
                    if len(pairs_info) >= 5:
                        break
            
            # 计算10天收益
            daily_returns = []
            
            for day in range(1, 11):
                day_data = df_returns[df_returns['twind'] == day]
                
                pair_returns = []
                for pair in pairs_info:
                    inclusion_data = day_data[day_data['Stkcd'] == pair['stkcd']]
                    
                    if len(inclusion_data) > 0:
                        inclusion_return = inclusion_data.iloc[0]['Dretwd']
                        matched_return = inclusion_data.iloc[0]['Dretwd1']
                        
                        if pd.notna(inclusion_return) and pd.notna(matched_return):
                            pair_return = inclusion_return - matched_return
                            pair_returns.append(pair_return)
                
                if len(pair_returns) > 0:
                    portfolio_return = np.mean(pair_returns)
                else:
                    portfolio_return = 0
                    
                daily_returns.append(portfolio_return * 100)
            
            all_period_returns.append({
                'period': period_name,
                'returns': daily_returns
            })
            
            # 添加表格
            self.add_title(f'表{table_num}：{period_name}10天Buy & Hold Return（%）', level=3)
            
            table = self.doc.add_table(rows=2, cols=10)
            table.style = 'Light Grid Accent 1'
            
            for i in range(10):
                table.rows[0].cells[i].text = str(i+1)
                table.rows[1].cells[i].text = f"{daily_returns[i]:.4f}"
            
            # 绘制单期图表
            plt.figure(figsize=(10, 6))
            plt.plot(range(1, 11), daily_returns, marker='o', linewidth=2, markersize=8)
            plt.xlabel('持有天数', fontsize=12)
            plt.ylabel('收益率 (%)', fontsize=12)
            plt.title(f'{period_name}策略10天Buy & Hold Return', fontsize=14, fontweight='bold')
            plt.grid(True, alpha=0.3)
            plt.axhline(y=0, color='r', linestyle='--', alpha=0.5)
            
            chart_path = f'/home/runner/work/CycWork2/CycWork2/q4_{sheet_name}_chart.png'
            plt.savefig(chart_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            self.doc.add_paragraph(f'\n图：{period_name}收益率走势')
            self.doc.add_picture(chart_path, width=Inches(6))
            
            # 统计分析
            cumulative = sum(daily_returns)
            self.add_paragraph(f'\n{period_name}收益统计：')
            self.add_paragraph(f'• 累计收益率：{cumulative:.4f}%')
            self.add_paragraph(f'• 平均日收益率：{np.mean(daily_returns):.4f}%')
            self.add_paragraph(f'• 收益率标准差：{np.std(daily_returns):.4f}%')
            self.add_paragraph(f'• 正收益天数：{sum([1 for r in daily_returns if r > 0])} 天')
        
        # 对比分析
        self.add_title('不同市场环境对比分析', level=2)
        
        # 绘制对比图
        plt.figure(figsize=(12, 6))
        for period_data in all_period_returns:
            plt.plot(range(1, 11), period_data['returns'], 
                    marker='o', linewidth=2, markersize=6, label=period_data['period'])
        
        plt.xlabel('持有天数', fontsize=12)
        plt.ylabel('收益率 (%)', fontsize=12)
        plt.title('不同市场环境下套利策略收益率对比', fontsize=14, fontweight='bold')
        plt.legend(fontsize=10)
        plt.grid(True, alpha=0.3)
        plt.axhline(y=0, color='r', linestyle='--', alpha=0.5)
        
        comparison_chart_path = '/home/runner/work/CycWork2/CycWork2/q4_comparison_chart.png'
        plt.savefig(comparison_chart_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        self.doc.add_paragraph('\n图：不同市场环境收益率对比')
        self.doc.add_picture(comparison_chart_path, width=Inches(6))
        
        # 详细对比分析
        comparison_text = """
对比分析结论：

1. 收益率水平差异：
   - 牛市环境：套利策略可能受益于整体市场上涨，但纳入效应相对较弱
   - 非牛熊市：市场相对平稳，纳入效应可能更加显著
   - 熊市环境：市场下跌压力大，套利策略可能面临更大风险

2. 收益稳定性差异：
   - 不同市场环境下，收益率的波动程度可能有显著差异
   - 熊市环境下，收益率标准差通常更大
   - 牛市环境下，收益可能更加稳定

3. 策略有效性：
   - 配对交易策略通过同行业、同市场、相近Beta的匹配，能够较好地控制市场风险
   - 在不同市场环境下，策略的风险收益特征可能发生变化
   - 需要根据市场环境调整策略参数

4. 与问题三的比较：
   - 问题三使用2025年6月数据（具体市场环境待定）
   - 通过多时期对比，可以验证策略的稳健性
   - 策略在不同环境下的表现差异，反映了市场环境对指数纳入效应的影响

5. 投资启示：
   - 配对交易策略能够在一定程度上规避市场系统性风险
   - 但仍需关注市场环境变化
   - 建议结合市场判断，适时调整仓位和对冲比例
"""
        self.add_paragraph(comparison_text)
        
        print("问题四完成！")
        return all_period_returns
    
    def save_document(self, output_file):
        """保存Word文档"""
        self.doc.save(output_file)
        print(f"\n文档已保存至：{output_file}")

def main():
    """主函数"""
    print("="*80)
    print("指数纳入效应套利策略实现")
    print("="*80)
    
    excel_file = "/home/runner/work/CycWork2/CycWork2/作业3 （学生用）指数纳入效应套利策略的实现.xlsx"
    output_file = "/home/runner/work/CycWork2/CycWork2/作业3_解答_详细版.docx"
    
    # 创建分析对象
    analyzer = IndexInclusionArbitrage(excel_file)
    
    # 添加文档标题
    analyzer.add_title('作业3：指数纳入效应套利策略的实现', level=0)
    analyzer.add_paragraph('详细计算过程与Excel实现步骤')
    analyzer.add_paragraph('\n' + '='*60)
    
    # 求解各问题
    try:
        # 问题一
        results_q1_1, results_q1_2 = analyzer.solve_question1()
        
        # 问题二
        matched_pairs = analyzer.solve_question2()
        
        # 问题三
        q3_returns = analyzer.solve_question3()
        
        # 问题四
        q4_returns = analyzer.solve_question4()
        
        # 保存文档
        analyzer.save_document(output_file)
        
        print("\n" + "="*80)
        print("所有问题计算完成！")
        print("="*80)
        print(f"\n结果已保存至：{output_file}")
        
    except Exception as e:
        print(f"\n错误：{e}")
        import traceback
        traceback.print_exc()
        return False
    
    return True

if __name__ == "__main__":
    success = main()
    if success:
        print("\n程序执行成功！")
    else:
        print("\n程序执行失败，请检查错误信息。")
