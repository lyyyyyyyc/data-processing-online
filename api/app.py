from flask import Flask, request, jsonify, render_template_string, send_file
import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from scipy import stats
import os
import io
import uuid
from datetime import datetime
import traceback
import tempfile
import base64

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

class DataProcessor:
    def __init__(self):
        self.df = None
        self.code_history = []
    
    def load_data_from_base64(self, file_content, filename):
        """从base64数据加载Excel"""
        try:
            # 解码base64数据
            file_data = base64.b64decode(file_content.split(',')[1])
            
            # 使用BytesIO创建文件对象
            file_obj = io.BytesIO(file_data)
            
            if filename.endswith('.xlsx') or filename.endswith('.xls'):
                self.df = pd.read_excel(file_obj)
                code = f"""
# 数据加载代码
import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from scipy import stats

# 加载数据
df = pd.read_excel('{filename}')
print(f"数据形状: {self.df.shape}")
print(f"列名: {list(self.df.columns)}")
"""
                self.code_history.append(("数据加载", code))
                return True, f"成功加载数据，形状: {self.df.shape}"
            else:
                return False, "不支持的文件格式，请上传Excel文件"
        except Exception as e:
            return False, f"文件加载失败: {str(e)}"
    
    def handle_missing_values(self, method='drop', fill_value=None):
        """处理缺失值"""
        if self.df is None:
            return False, "请先上传数据文件", ""
        
        try:
            original_shape = self.df.shape
            
            if method == 'drop':
                self.df = self.df.dropna()
                code = f"""
# 缺失值处理 - 删除含有缺失值的行
df_cleaned = df.dropna()
print(f"原始数据形状: {original_shape}")
print(f"处理后数据形状: {self.df.shape}")
"""
            elif method == 'fill_mean':
                numeric_cols = self.df.select_dtypes(include=[np.number]).columns
                self.df[numeric_cols] = self.df[numeric_cols].fillna(self.df[numeric_cols].mean())
                code = f"""
# 缺失值处理 - 用均值填充数值列
numeric_cols = df.select_dtypes(include=[np.number]).columns
df_cleaned = df.copy()
df_cleaned[numeric_cols] = df_cleaned[numeric_cols].fillna(df_cleaned[numeric_cols].mean())
print(f"数值列: {list(numeric_cols)}")
print("用均值填充缺失值")
"""
            elif method == 'fill_median':
                numeric_cols = self.df.select_dtypes(include=[np.number]).columns
                self.df[numeric_cols] = self.df[numeric_cols].fillna(self.df[numeric_cols].median())
                code = f"""
# 缺失值处理 - 用中位数填充数值列
numeric_cols = df.select_dtypes(include=[np.number]).columns
df_cleaned = df.copy()
df_cleaned[numeric_cols] = df_cleaned[numeric_cols].fillna(df_cleaned[numeric_cols].median())
print(f"数值列: {list(numeric_cols)}")
print("用中位数填充缺失值")
"""
            elif method == 'fill_value' and fill_value is not None:
                self.df = self.df.fillna(fill_value)
                code = f"""
# 缺失值处理 - 用指定值填充
df_cleaned = df.fillna({fill_value})
print(f"用值 {fill_value} 填充所有缺失值")
"""
            
            self.code_history.append(("缺失值处理", code))
            return True, f"缺失值处理完成。原始形状: {original_shape}, 处理后形状: {self.df.shape}", code
            
        except Exception as e:
            return False, f"缺失值处理失败: {str(e)}", ""
    
    def handle_outliers(self, method='iqr', threshold=3):
        """处理异常值"""
        if self.df is None:
            return False, "请先上传数据文件", ""
        
        try:
            numeric_cols = self.df.select_dtypes(include=[np.number]).columns
            original_shape = self.df.shape
            
            if method == 'iqr':
                Q1 = self.df[numeric_cols].quantile(0.25)
                Q3 = self.df[numeric_cols].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                
                mask = True
                for col in numeric_cols:
                    mask = mask & (self.df[col] >= lower_bound[col]) & (self.df[col] <= upper_bound[col])
                
                self.df = self.df[mask]
                
                code = f"""
# 异常值处理 - IQR方法
numeric_cols = df.select_dtypes(include=[np.number]).columns
Q1 = df[numeric_cols].quantile(0.25)
Q3 = df[numeric_cols].quantile(0.75)
IQR = Q3 - Q1
lower_bound = Q1 - 1.5 * IQR
upper_bound = Q3 + 1.5 * IQR

mask = True
for col in numeric_cols:
    mask = mask & (df[col] >= lower_bound[col]) & (df[col] <= upper_bound[col])

df_no_outliers = df[mask]
print(f"原始数据形状: {original_shape}")
print(f"处理后数据形状: {self.df.shape}")
"""
            
            elif method == 'zscore':
                z_scores = np.abs(stats.zscore(self.df[numeric_cols]))
                mask = (z_scores < threshold).all(axis=1)
                self.df = self.df[mask]
                
                code = f"""
# 异常值处理 - Z-score方法 (阈值: {threshold})
from scipy import stats
import numpy as np

numeric_cols = df.select_dtypes(include=[np.number]).columns
z_scores = np.abs(stats.zscore(df[numeric_cols]))
mask = (z_scores < {threshold}).all(axis=1)
df_no_outliers = df[mask]
print(f"原始数据形状: {original_shape}")
print(f"处理后数据形状: {self.df.shape}")
"""
            
            self.code_history.append(("异常值处理", code))
            return True, f"异常值处理完成。原始形状: {original_shape}, 处理后形状: {self.df.shape}", code
            
        except Exception as e:
            return False, f"异常值处理失败: {str(e)}", ""
    
    def handle_duplicates(self):
        """处理重复值"""
        if self.df is None:
            return False, "请先上传数据文件", ""
        
        try:
            original_shape = self.df.shape
            duplicates_count = self.df.duplicated().sum()
            self.df = self.df.drop_duplicates()
            
            code = f"""
# 重复值处理
print(f"发现重复行数: {duplicates_count}")
df_no_duplicates = df.drop_duplicates()
print(f"原始数据形状: {original_shape}")
print(f"处理后数据形状: {self.df.shape}")
"""
            
            self.code_history.append(("重复值处理", code))
            return True, f"重复值处理完成。删除了 {duplicates_count} 行重复数据。原始形状: {original_shape}, 处理后形状: {self.df.shape}", code
            
        except Exception as e:
            return False, f"重复值处理失败: {str(e)}", ""
    
    def standardize_data(self, method='zscore'):
        """数据标准化"""
        if self.df is None:
            return False, "请先上传数据文件", ""
        
        try:
            numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()
            
            if not numeric_cols:
                return False, "没有找到可标准化的数值列", ""
            
            if method == 'zscore':
                scaler = StandardScaler()
                self.df[numeric_cols] = scaler.fit_transform(self.df[numeric_cols])
                
                code = f"""
# Z-score标准化
from sklearn.preprocessing import StandardScaler

numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
scaler = StandardScaler()
df_standardized = df.copy()
df_standardized[numeric_cols] = scaler.fit_transform(df_standardized[numeric_cols])

print(f"标准化的列: {numeric_cols}")
print("使用Z-score标准化: (x - μ) / σ")
"""
            
            elif method == 'minmax':
                scaler = MinMaxScaler()
                self.df[numeric_cols] = scaler.fit_transform(self.df[numeric_cols])
                
                code = f"""
# Min-Max标准化
from sklearn.preprocessing import MinMaxScaler

numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
scaler = MinMaxScaler()
df_standardized = df.copy()
df_standardized[numeric_cols] = scaler.fit_transform(df_standardized[numeric_cols])

print(f"标准化的列: {numeric_cols}")
print("使用Min-Max标准化: (x - min) / (max - min)")
"""
            
            self.code_history.append(("数据标准化", code))
            return True, f"数据标准化完成。标准化列: {numeric_cols}", code
            
        except Exception as e:
            return False, f"数据标准化失败: {str(e)}", ""
    
    def correlation_analysis(self):
        """相关性分析"""
        if self.df is None:
            return False, "请先上传数据文件", "", False
        
        try:
            numeric_cols = self.df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) < 2:
                return False, "需要至少2个数值列进行相关性分析", "", False
            
            code = f"""
# 相关性分析
numeric_cols = df.select_dtypes(include=[np.number]).columns
correlation_matrix = df[numeric_cols].corr()

print("相关性矩阵:")
print(correlation_matrix)

high_corr_pairs = []
for i in range(len(correlation_matrix.columns)):
    for j in range(i+1, len(correlation_matrix.columns)):
        corr_value = correlation_matrix.iloc[i, j]
        if abs(corr_value) > 0.7:
            high_corr_pairs.append((correlation_matrix.columns[i], 
                                  correlation_matrix.columns[j], 
                                  corr_value))

print("\\n高相关性特征对 (|r| > 0.7):")
for pair in high_corr_pairs:
    print(f"{{pair[0]}} - {{pair[1]}}: {{pair[2]:.3f}}")
"""
            
            self.code_history.append(("相关性分析", code))
            return True, f"相关性分析完成。分析了 {len(numeric_cols)} 个数值列", code, False
            
        except Exception as e:
            return False, f"相关性分析失败: {str(e)}", "", False
    
    def t_test(self, column1, column2=None, value=None):
        """t检验"""
        if self.df is None:
            return False, "请先上传数据文件", "", False
        
        try:
            if column1 not in self.df.columns:
                return False, f"列 '{column1}' 不存在", "", False
            
            if column2 is not None:
                # 双样本t检验
                if column2 not in self.df.columns:
                    return False, f"列 '{column2}' 不存在", "", False
                
                data1 = self.df[column1].dropna()
                data2 = self.df[column2].dropna()
                
                t_stat, p_value = stats.ttest_ind(data1, data2)
                
                code = f"""
# 双样本t检验
from scipy import stats

data1 = df['{column1}'].dropna()
data2 = df['{column2}'].dropna()

t_statistic, p_value = stats.ttest_ind(data1, data2)

print(f"双样本t检验结果:")
print(f"列1: {column1}, 样本数: {{len(data1)}}, 均值: {{data1.mean():.4f}}")
print(f"列2: {column2}, 样本数: {{len(data2)}}, 均值: {{data2.mean():.4f}}")
print(f"t统计量: {{t_statistic:.4f}}")
print(f"p值: {{p_value:.4f}}")
print(f"显著性水平0.05下{'拒绝' if p_value < 0.05 else '接受'}原假设")
"""
                
                result_text = f"双样本t检验: t统计量={t_stat:.4f}, p值={p_value:.4f}"
                
            elif value is not None:
                # 单样本t检验
                data = self.df[column1].dropna()
                t_stat, p_value = stats.ttest_1samp(data, value)
                
                code = f"""
# 单样本t检验
from scipy import stats

data = df['{column1}'].dropna()
test_value = {value}

t_statistic, p_value = stats.ttest_1samp(data, test_value)

print(f"单样本t检验结果:")
print(f"列: {column1}, 样本数: {{len(data)}}, 样本均值: {{data.mean():.4f}}")
print(f"检验值: {value}")
print(f"t统计量: {{t_statistic:.4f}}")
print(f"p值: {{p_value:.4f}}")
print(f"显著性水平0.05下{'拒绝' if p_value < 0.05 else '接受'}原假设")
"""
                
                result_text = f"单样本t检验: t统计量={t_stat:.4f}, p值={p_value:.4f}"
            
            else:
                return False, "请指定第二列或检验值", "", False
            
            self.code_history.append(("t检验", code))
            return True, result_text, code, False
            
        except Exception as e:
            return False, f"t检验失败: {str(e)}", "", False
    
    def chi_square_test(self, column1, column2):
        """卡方检验"""
        if self.df is None:
            return False, "请先上传数据文件", "", False
        
        try:
            if column1 not in self.df.columns or column2 not in self.df.columns:
                return False, "指定的列不存在", "", False
            
            # 创建列联表
            contingency_table = pd.crosstab(self.df[column1], self.df[column2])
            
            # 卡方检验
            chi2, p_value, dof, expected = stats.chi2_contingency(contingency_table)
            
            code = f"""
# 卡方检验
from scipy import stats
import pandas as pd

# 创建列联表
contingency_table = pd.crosstab(df['{column1}'], df['{column2}'])
print("列联表:")
print(contingency_table)

# 执行卡方检验
chi2_statistic, p_value, degrees_of_freedom, expected_frequencies = stats.chi2_contingency(contingency_table)

print(f"\\n卡方检验结果:")
print(f"卡方统计量: {{chi2_statistic:.4f}}")
print(f"p值: {{p_value:.4f}}")
print(f"自由度: {{degrees_of_freedom}}")
print(f"显著性水平0.05下{'拒绝' if p_value < 0.05 else '接受'}原假设(变量独立)")

print(f"\\n期望频率:")
print(expected_frequencies)
"""
            
            result_text = f"卡方检验: χ²={chi2:.4f}, p值={p_value:.4f}, 自由度={dof}"
            
            self.code_history.append(("卡方检验", code))
            return True, result_text, code, False
            
        except Exception as e:
            return False, f"卡方检验失败: {str(e)}", "", False
    
    def get_result_excel(self):
        """获取处理结果的Excel数据"""
        if self.df is None:
            return None
        
        try:
            output = io.BytesIO()
            self.df.to_excel(output, index=False)
            output.seek(0)
            return output.getvalue()
        except Exception as e:
            return None
    
    def get_complete_code(self):
        """获取完整的Python代码"""
        if not self.code_history:
            return "# 没有执行任何操作"
        
        complete_code = "# 完整的数据预处理Python代码\n"
        complete_code += "# 生成时间: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n\n"
        
        for operation, code in self.code_history:
            complete_code += f"# {operation}\n"
            complete_code += code + "\n\n"
        
        return complete_code

# 全局数据处理器
processor = DataProcessor()

# HTML模板
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>数据预处理在线工具</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.1.3/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .main-container {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1);
            margin: 20px auto;
            max-width: 1200px;
            backdrop-filter: blur(10px);
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            text-align: center;
            padding: 30px;
            border-radius: 20px 20px 0 0;
        }
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 15px;
            padding: 50px;
            text-align: center;
            margin: 30px;
            transition: all 0.3s ease;
            background: #f8f9fa;
            cursor: pointer;
        }
        .upload-area:hover {
            border-color: #764ba2;
            background: #e9ecef;
        }
        .btn-custom {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            border-radius: 25px;
            padding: 10px 30px;
            color: white;
            transition: all 0.3s ease;
        }
        .function-card {
            background: white;
            border-radius: 15px;
            padding: 20px;
            margin: 15px 0;
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.1);
        }
        .hidden { display: none; }
        .code-display {
            background: #2d3748;
            color: #e2e8f0;
            border-radius: 10px;
            padding: 20px;
            font-family: 'Monaco', monospace;
            font-size: 14px;
            overflow-x: auto;
            white-space: pre-wrap;
        }
        .data-info {
            background: #f8f9fa;
            border-radius: 15px;
            padding: 20px;
            margin: 30px;
        }
        .nav-pills .nav-link {
            border-radius: 25px;
            margin: 0 5px;
        }
        .nav-pills .nav-link.active {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="main-container">
            <div class="header">
                <h1><i class="fas fa-chart-line"></i> 数据预处理在线工具</h1>
                <p>轻松处理您的Excel数据 - 缺失值、异常值、标准化、统计分析一站式解决</p>
            </div>

            <div id="upload-section">
                <div class="upload-area" onclick="document.getElementById('fileInput').click()">
                    <div style="font-size: 4rem; color: #667eea; margin-bottom: 20px;">
                        <i class="fas fa-cloud-upload-alt"></i>
                    </div>
                    <h4>点击选择Excel文件</h4>
                    <p class="text-muted">支持 .xlsx 和 .xls 格式，最大 16MB</p>
                    <input type="file" id="fileInput" accept=".xlsx,.xls" style="display: none;">
                    <button class="btn btn-custom mt-3" type="button">
                        <i class="fas fa-folder-open"></i> 选择文件
                    </button>
                </div>
            </div>

            <div id="data-info-section" class="hidden">
                <div class="data-info">
                    <h4><i class="fas fa-info-circle"></i> 数据信息</h4>
                    <div id="dataInfo"></div>
                </div>
            </div>

            <div id="functions-section" class="hidden">
                <div class="container">
                    <h4 class="text-center mb-4"><i class="fas fa-cogs"></i> 选择数据处理功能</h4>
                    
                    <!-- 功能标签页 -->
                    <ul class="nav nav-pills justify-content-center mb-4" role="tablist">
                        <li class="nav-item" role="presentation">
                            <button class="nav-link active" id="preprocessing-tab" data-bs-toggle="pill" data-bs-target="#preprocessing" type="button">
                                <i class="fas fa-broom"></i> 数据清洗
                            </button>
                        </li>
                        <li class="nav-item" role="presentation">
                            <button class="nav-link" id="standardization-tab" data-bs-toggle="pill" data-bs-target="#standardization" type="button">
                                <i class="fas fa-balance-scale"></i> 数据标准化
                            </button>
                        </li>
                        <li class="nav-item" role="presentation">
                            <button class="nav-link" id="analysis-tab" data-bs-toggle="pill" data-bs-target="#analysis" type="button">
                                <i class="fas fa-chart-bar"></i> 统计分析
                            </button>
                        </li>
                    </ul>

                    <!-- 功能内容 -->
                    <div class="tab-content">
                        <!-- 数据清洗 -->
                        <div class="tab-pane fade show active" id="preprocessing">
                            <div class="row">
                                <div class="col-md-4">
                                    <div class="function-card">
                                        <h5><i class="fas fa-exclamation-triangle"></i> 缺失值处理</h5>
                                        <select class="form-select mb-3" id="missingMethod">
                                            <option value="drop">删除含缺失值的行</option>
                                            <option value="fill_mean">用均值填充</option>
                                            <option value="fill_median">用中位数填充</option>
                                            <option value="fill_value">用指定值填充</option>
                                        </select>
                                        <div class="mb-3" id="fillValueDiv" style="display: none;">
                                            <input type="number" class="form-control" id="fillValue" placeholder="填充值" step="any">
                                        </div>
                                        <button class="btn btn-custom w-100" onclick="processMissingValues()">执行处理</button>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="function-card">
                                        <h5><i class="fas fa-search"></i> 异常值处理</h5>
                                        <select class="form-select mb-3" id="outlierMethod">
                                            <option value="iqr">IQR方法</option>
                                            <option value="zscore">Z-score方法</option>
                                        </select>
                                        <div class="mb-3" id="thresholdDiv" style="display: none;">
                                            <input type="number" class="form-control" id="zThreshold" placeholder="Z-score阈值" value="3" step="0.1">
                                        </div>
                                        <button class="btn btn-custom w-100" onclick="processOutliers()">执行处理</button>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="function-card">
                                        <h5><i class="fas fa-copy"></i> 重复值处理</h5>
                                        <p class="text-muted">自动检测并删除重复的行</p>
                                        <button class="btn btn-custom w-100" onclick="processDuplicates()">执行处理</button>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- 数据标准化 -->
                        <div class="tab-pane fade" id="standardization">
                            <div class="row justify-content-center">
                                <div class="col-md-6">
                                    <div class="function-card">
                                        <h5><i class="fas fa-ruler"></i> 数据标准化</h5>
                                        <select class="form-select mb-3" id="standardizationMethod">
                                            <option value="zscore">Z-score标准化</option>
                                            <option value="minmax">Min-Max标准化</option>
                                        </select>
                                        <div class="mb-3">
                                            <small class="text-muted">
                                                Z-score: (x - μ) / σ<br>
                                                Min-Max: (x - min) / (max - min)
                                            </small>
                                        </div>
                                        <button class="btn btn-custom w-100" onclick="processStandardization()">执行标准化</button>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <!-- 统计分析 -->
                        <div class="tab-pane fade" id="analysis">
                            <div class="row">
                                <div class="col-md-4">
                                    <div class="function-card">
                                        <h5><i class="fas fa-project-diagram"></i> 相关性分析</h5>
                                        <p class="text-muted">分析数值列之间的相关性</p>
                                        <button class="btn btn-custom w-100" onclick="processCorrelation()">执行分析</button>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="function-card">
                                        <h5><i class="fas fa-calculator"></i> t检验</h5>
                                        <div class="mb-3">
                                            <select class="form-select" id="tTestColumn1">
                                                <option value="">选择第一列</option>
                                            </select>
                                        </div>
                                        <div class="mb-3">
                                            <select class="form-select" id="tTestType">
                                                <option value="two_sample">双样本检验</option>
                                                <option value="one_sample">单样本检验</option>
                                            </select>
                                        </div>
                                        <div class="mb-3" id="tTestColumn2Div">
                                            <select class="form-select" id="tTestColumn2">
                                                <option value="">选择第二列</option>
                                            </select>
                                        </div>
                                        <div class="mb-3" id="tTestValueDiv" style="display: none;">
                                            <input type="number" class="form-control" id="tTestValue" placeholder="检验值" step="any">
                                        </div>
                                        <button class="btn btn-custom w-100" onclick="processTTest()">执行检验</button>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <div class="function-card">
                                        <h5><i class="fas fa-th"></i> 卡方检验</h5>
                                        <div class="mb-3">
                                            <select class="form-select" id="chiSquareColumn1">
                                                <option value="">选择第一列</option>
                                            </select>
                                        </div>
                                        <div class="mb-3">
                                            <select class="form-select" id="chiSquareColumn2">
                                                <option value="">选择第二列</option>
                                            </select>
                                        </div>
                                        <button class="btn btn-custom w-100" onclick="processChiSquare()">执行检验</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div id="loading" class="hidden text-center p-5">
                <div class="spinner-border text-primary" role="status"></div>
                <p class="mt-3">正在处理数据，请稍候...</p>
            </div>

            <div id="result-section" class="hidden">
                <div class="container">
                    <div id="resultMessage"></div>
                    <div id="downloadSection" class="text-center mb-4">
                        <button class="btn btn-success btn-lg" onclick="downloadResult()">
                            <i class="fas fa-download"></i> 下载处理后的Excel文件
                        </button>
                    </div>
                    <div class="mb-4">
                        <h5><i class="fas fa-code"></i> 完整Python代码</h5>
                        <div class="d-flex justify-content-end mb-2">
                            <button class="btn btn-outline-secondary btn-sm" onclick="copyCode()">
                                <i class="fas fa-copy"></i> 复制代码
                            </button>
                        </div>
                        <div class="code-display" id="pythonCode"></div>
                    </div>
                    <div class="text-center">
                        <button class="btn btn-custom" onclick="resetProcessor()">
                            <i class="fas fa-refresh"></i> 处理新文件
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.1.3/js/bootstrap.bundle.min.js"></script>
    <script>
        let currentDataInfo = null;

        document.addEventListener('DOMContentLoaded', function() {
            const fileInput = document.getElementById('fileInput');
            const missingMethod = document.getElementById('missingMethod');
            const outlierMethod = document.getElementById('outlierMethod');
            const tTestType = document.getElementById('tTestType');
            
            fileInput.addEventListener('change', handleFileSelect);
            missingMethod.addEventListener('change', toggleFillValue);
            outlierMethod.addEventListener('change', toggleThreshold);
            tTestType.addEventListener('change', toggleTTestInputs);
        });

        function handleFileSelect(e) {
            const file = e.target.files[0];
            if (file) {
                handleFile(file);
            }
        }

        function handleFile(file) {
            if (!file.name.match(/\\.(xlsx|xls)$/)) {
                alert('请选择Excel文件');
                return;
            }
            
            if (file.size > 16 * 1024 * 1024) {
                alert('文件大小不能超过16MB');
                return;
            }
            
            showLoading(true);
            
            const reader = new FileReader();
            reader.onload = function(e) {
                uploadFile(e.target.result, file.name);
            };
            reader.readAsDataURL(file);
        }

        function uploadFile(fileContent, filename) {
            fetch('/upload', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    file_content: fileContent,
                    filename: filename
                })
            })
            .then(response => response.json())
            .then(data => {
                showLoading(false);
                if (data.success) {
                    currentDataInfo = data.data_info;
                    displayDataInfo(data.data_info);
                    populateColumnSelects(data.data_info.columns);
                    showSection('data-info-section');
                    showSection('functions-section');
                    hideSection('upload-section');
                    showAlert(data.message, 'success');
                } else {
                    showAlert(data.message, 'danger');
                }
            })
            .catch(error => {
                showLoading(false);
                showAlert('上传失败：' + error.message, 'danger');
            });
        }

        function displayDataInfo(dataInfo) {
            let missingHtml = '';
            let hasMissing = false;
            
            for (const [col, count] of Object.entries(dataInfo.missing_values)) {
                if (count > 0) {
                    missingHtml += `<p><strong>${col}：</strong> ${count} 个缺失值</p>`;
                    hasMissing = true;
                }
            }
            
            if (!hasMissing) {
                missingHtml = '<p class="text-success">无缺失值</p>';
            }

            const html = `
                <div class="row">
                    <div class="col-md-6">
                        <h6><i class="fas fa-table"></i> 基本信息</h6>
                        <p><strong>数据形状：</strong> ${dataInfo.shape[0]} 行 × ${dataInfo.shape[1]} 列</p>
                        <p><strong>列名：</strong> ${dataInfo.columns.join(', ')}</p>
                    </div>
                    <div class="col-md-6">
                        <h6><i class="fas fa-exclamation-circle"></i> 缺失值统计</h6>
                        ${missingHtml}
                    </div>
                </div>
            `;
            document.getElementById('dataInfo').innerHTML = html;
        }

        function populateColumnSelects(columns) {
            const selects = ['tTestColumn1', 'tTestColumn2', 'chiSquareColumn1', 'chiSquareColumn2'];
            
            selects.forEach(selectId => {
                const select = document.getElementById(selectId);
                if (select) {
                    select.innerHTML = '<option value="">选择列</option>';
                    columns.forEach(col => {
                        const option = document.createElement('option');
                        option.value = col;
                        option.textContent = col;
                        select.appendChild(option);
                    });
                }
            });
        }

        function toggleFillValue() {
            const method = document.getElementById('missingMethod').value;
            const fillValueDiv = document.getElementById('fillValueDiv');
            fillValueDiv.style.display = method === 'fill_value' ? 'block' : 'none';
        }

        function toggleThreshold() {
            const method = document.getElementById('outlierMethod').value;
            const thresholdDiv = document.getElementById('thresholdDiv');
            thresholdDiv.style.display = method === 'zscore' ? 'block' : 'none';
        }

        function toggleTTestInputs() {
            const testType = document.getElementById('tTestType').value;
            const column2Div = document.getElementById('tTestColumn2Div');
            const valueDiv = document.getElementById('tTestValueDiv');
            
            if (testType === 'two_sample') {
                column2Div.style.display = 'block';
                valueDiv.style.display = 'none';
            } else {
                column2Div.style.display = 'none';
                valueDiv.style.display = 'block';
            }
        }

        function processMissingValues() {
            const method = document.getElementById('missingMethod').value;
            const fillValue = document.getElementById('fillValue').value;
            
            const params = { method };
            if (method === 'fill_value' && fillValue !== '') {
                params.fill_value = parseFloat(fillValue);
            }
            
            processData('missing_values', params);
        }

        function processOutliers() {
            const method = document.getElementById('outlierMethod').value;
            const threshold = document.getElementById('zThreshold').value;
            
            const params = { method };
            if (method === 'zscore') {
                params.threshold = parseFloat(threshold);
            }
            
            processData('outliers', params);
        }

        function processDuplicates() {
            processData('duplicates', {});
        }

        function processStandardization() {
            const method = document.getElementById('standardizationMethod').value;
            processData('standardization', { method });
        }

        function processCorrelation() {
            processData('correlation', {});
        }

        function processTTest() {
            const column1 = document.getElementById('tTestColumn1').value;
            const testType = document.getElementById('tTestType').value;
            
            if (!column1) {
                showAlert('请选择第一列', 'warning');
                return;
            }
            
            const params = { column1 };
            
            if (testType === 'two_sample') {
                const column2 = document.getElementById('tTestColumn2').value;
                if (!column2) {
                    showAlert('请选择第二列', 'warning');
                    return;
                }
                params.column2 = column2;
            } else {
                const value = document.getElementById('tTestValue').value;
                if (value === '') {
                    showAlert('请输入检验值', 'warning');
                    return;
                }
                params.value = parseFloat(value);
            }
            
            processData('t_test', params);
        }

        function processChiSquare() {
            const column1 = document.getElementById('chiSquareColumn1').value;
            const column2 = document.getElementById('chiSquareColumn2').value;
            
            if (!column1 || !column2) {
                showAlert('请选择两列进行卡方检验', 'warning');
                return;
            }
            
            processData('chi_square', { column1, column2 });
        }

        function processData(operation, parameters) {
            showLoading(true);
            hideSection('result-section');
            
            fetch('/process', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    operation: operation,
                    parameters: parameters
                })
            })
            .then(response => response.json())
            .then(data => {
                showLoading(false);
                displayResult(data);
            })
            .catch(error => {
                showLoading(false);
                showAlert('处理失败：' + error.message, 'danger');
            });
        }

        function displayResult(data) {
            const resultMessage = document.getElementById('resultMessage');
            const pythonCode = document.getElementById('pythonCode');
            const downloadSection = document.getElementById('downloadSection');
            
            const messageClass = data.success ? 'alert-success' : 'alert-danger';
            resultMessage.innerHTML = `<div class="alert ${messageClass}">${data.message}</div>`;
            
            pythonCode.textContent = data.complete_code;
            
            downloadSection.style.display = data.can_download ? 'block' : 'none';
            
            showSection('result-section');
            
            document.getElementById('result-section').scrollIntoView({ behavior: 'smooth' });
        }

        function downloadResult() {
            window.location.href = '/download';
        }

        function copyCode() {
            const codeElement = document.getElementById('pythonCode');
            const textArea = document.createElement('textarea');
            textArea.value = codeElement.textContent;
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
            
            showAlert('代码已复制到剪贴板', 'success');
        }

        function resetProcessor() {
            fetch('/reset', { method: 'POST' })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    hideSection('data-info-section');
                    hideSection('functions-section');
                    hideSection('result-section');
                    showSection('upload-section');
                    document.getElementById('fileInput').value = '';
                    showAlert(data.message, 'success');
                }
            });
        }

        function showSection(sectionId) {
            document.getElementById(sectionId).classList.remove('hidden');
        }

        function hideSection(sectionId) {
            document.getElementById(sectionId).classList.add('hidden');
        }

        function showLoading(show) {
            const loading = document.getElementById('loading');
            loading.style.display = show ? 'block' : 'none';
        }

        function showAlert(message, type) {
            const alertDiv = document.createElement('div');
            alertDiv.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
            alertDiv.style.cssText = 'top: 20px; right: 20px; z-index: 9999; max-width: 400px;';
            alertDiv.innerHTML = `
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            `;
            
            document.body.appendChild(alertDiv);
            
            setTimeout(() => {
                if (alertDiv.parentNode) {
                    alertDiv.parentNode.removeChild(alertDiv);
                }
            }, 3000);
        }
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        data = request.json
        file_content = data.get('file_content')
        filename = data.get('filename')
        
        success, message = processor.load_data_from_base64(file_content, filename)
        
        if success:
            info = {
                'shape': list(processor.df.shape),
                'columns': list(processor.df.columns),
                'missing_values': {k: int(v) for k, v in processor.df.isnull().sum().to_dict().items()}
            }
            return jsonify({'success': True, 'message': message, 'data_info': info})
        else:
            return jsonify({'success': False, 'message': message})
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'上传失败: {str(e)}'})

@app.route('/process', methods=['POST'])
def process_data():
    try:
        data = request.json
        operation = data.get('operation')
        params = data.get('parameters', {})
        
        success = False
        message = ""
        code = ""
        can_download = True
        
        if operation == 'missing_values':
            method = params.get('method', 'drop')
            fill_value = params.get('fill_value')
            success, message, code = processor.handle_missing_values(method, fill_value)
        
        elif operation == 'outliers':
            method = params.get('method', 'iqr')
            threshold = params.get('threshold', 3)
            success, message, code = processor.handle_outliers(method, threshold)
        
        elif operation == 'duplicates':
            success, message, code = processor.handle_duplicates()
        
        elif operation == 'standardization':
            method = params.get('method', 'zscore')
            success, message, code = processor.standardize_data(method)
        
        elif operation == 'correlation':
            success, message, code, can_download = processor.correlation_analysis()
        
        elif operation == 't_test':
            column1 = params.get('column1')
            column2 = params.get('column2')
            value = params.get('value')
            success, message, code, can_download = processor.t_test(column1, column2, value)
        
        elif operation == 'chi_square':
            column1 = params.get('column1')
            column2 = params.get('column2')
            success, message, code, can_download = processor.chi_square_test(column1, column2)
        
        if not success and can_download:
            can_download = False
            message = "此种功能无法给出表格"
        
        complete_code = processor.get_complete_code()
        
        return jsonify({
            'success': success,
            'message': message,
            'code': code,
            'complete_code': complete_code,
            'can_download': can_download and success
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'处理失败: {str(e)}',
            'code': "",
            'complete_code': processor.get_complete_code(),
            'can_download': False
        })

@app.route('/download')
def download_result():
    try:
        excel_data = processor.get_result_excel()
        if excel_data:
            filename = f'processed_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            
            response = app.response_class(
                excel_data,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={"Content-disposition": f"attachment; filename={filename}"}
            )
            return response
        else:
            return jsonify({'success': False, 'message': '没有可下载的数据'})
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'下载失败: {str(e)}'})

@app.route('/reset', methods=['POST'])
def reset_processor():
    global processor
    processor = DataProcessor()
    return jsonify({'success': True, 'message': '已重置，可以上传新文件'})

# Vercel需要这个入口点
app_instance = app

if __name__ == '__main__':
    app.run(debug=True)
