from flask import Flask, request, jsonify, render_template_string, send_file
import pandas as pd
import os
import io
import uuid
from datetime import datetime
import traceback
import tempfile
import base64
import math

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
            
            # 根据文件扩展名选择读取方法
            if filename.endswith('.xlsx'):
                self.df = pd.read_excel(file_obj, engine='openpyxl')
            elif filename.endswith('.xls'):
                self.df = pd.read_excel(file_obj)
            else:
                return False, "不支持的文件格式"
            
            self.code_history = [
                "import pandas as pd",
                f"# 读取Excel文件",
                f"df = pd.read_excel('{filename}')"
            ]
            
            return True, f"成功加载数据，共{self.df.shape[0]}行{self.df.shape[1]}列"
        
        except Exception as e:
            return False, f"文件加载失败: {str(e)}"
    
    def handle_missing_values(self, method='drop', fill_value=None):
        """处理缺失值"""
        try:
            original_shape = self.df.shape
            
            if method == 'drop':
                self.df = self.df.dropna()
                code = "df = df.dropna()"
            elif method == 'mean':
                numeric_cols = self.df.select_dtypes(include=['number']).columns
                self.df[numeric_cols] = self.df[numeric_cols].fillna(self.df[numeric_cols].mean())
                code = "numeric_cols = df.select_dtypes(include=['number']).columns\ndf[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())"
            elif method == 'median':
                numeric_cols = self.df.select_dtypes(include=['number']).columns
                self.df[numeric_cols] = self.df[numeric_cols].fillna(self.df[numeric_cols].median())
                code = "numeric_cols = df.select_dtypes(include=['number']).columns\ndf[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].median())"
            elif method == 'value' and fill_value is not None:
                self.df = self.df.fillna(fill_value)
                code = f"df = df.fillna({repr(fill_value)})"
            else:
                return False, "无效的缺失值处理方法", ""
            
            self.code_history.append(f"# 处理缺失值 - {method}")
            self.code_history.append(code)
            
            new_shape = self.df.shape
            message = f"缺失值处理完成。原始数据: {original_shape[0]}行{original_shape[1]}列，处理后: {new_shape[0]}行{new_shape[1]}列"
            
            return True, message, code
        
        except Exception as e:
            return False, f"缺失值处理失败: {str(e)}", ""
    
    def handle_outliers(self, method='iqr', threshold=3):
        """处理异常值"""
        try:
            original_shape = self.df.shape
            numeric_cols = self.df.select_dtypes(include=['number']).columns
            
            if len(numeric_cols) == 0:
                return False, "没有数值列可以处理异常值", ""
            
            if method == 'iqr':
                for col in numeric_cols:
                    Q1 = self.df[col].quantile(0.25)
                    Q3 = self.df[col].quantile(0.75)
                    IQR = Q3 - Q1
                    lower_bound = Q1 - 1.5 * IQR
                    upper_bound = Q3 + 1.5 * IQR
                    self.df = self.df[(self.df[col] >= lower_bound) & (self.df[col] <= upper_bound)]
                
                code = """# IQR方法处理异常值
numeric_cols = df.select_dtypes(include=['number']).columns
for col in numeric_cols:
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    df = df[(df[col] >= lower_bound) & (df[col] <= upper_bound)]"""
            
            elif method == 'zscore':
                for col in numeric_cols:
                    mean = self.df[col].mean()
                    std = self.df[col].std()
                    if std > 0:
                        z_scores = abs((self.df[col] - mean) / std)
                        self.df = self.df[z_scores < threshold]
                
                code = f"""# Z-score方法处理异常值
numeric_cols = df.select_dtypes(include=['number']).columns
for col in numeric_cols:
    mean = df[col].mean()
    std = df[col].std()
    if std > 0:
        z_scores = abs((df[col] - mean) / std)
        df = df[z_scores < {threshold}]"""
            
            self.code_history.append(f"# 处理异常值 - {method}")
            self.code_history.append(code)
            
            new_shape = self.df.shape
            message = f"异常值处理完成。原始数据: {original_shape[0]}行，处理后: {new_shape[0]}行"
            
            return True, message, code
        
        except Exception as e:
            return False, f"异常值处理失败: {str(e)}", ""
    
    def handle_duplicates(self):
        """处理重复值"""
        try:
            original_shape = self.df.shape
            self.df = self.df.drop_duplicates()
            new_shape = self.df.shape
            
            code = "df = df.drop_duplicates()"
            self.code_history.append("# 删除重复行")
            self.code_history.append(code)
            
            removed_count = original_shape[0] - new_shape[0]
            message = f"重复值处理完成。删除了 {removed_count} 行重复数据，剩余 {new_shape[0]} 行"
            
            return True, message, code
        
        except Exception as e:
            return False, f"重复值处理失败: {str(e)}", ""
    
    def standardize_data(self, method='zscore'):
        """数据标准化"""
        try:
            numeric_cols = self.df.select_dtypes(include=['number']).columns
            
            if len(numeric_cols) == 0:
                return False, "没有数值列可以标准化", ""
            
            if method == 'zscore':
                for col in numeric_cols:
                    mean = self.df[col].mean()
                    std = self.df[col].std()
                    if std > 0:
                        self.df[col] = (self.df[col] - mean) / std
                
                code = """# Z-score标准化
numeric_cols = df.select_dtypes(include=['number']).columns
for col in numeric_cols:
    mean = df[col].mean()
    std = df[col].std()
    if std > 0:
        df[col] = (df[col] - mean) / std"""
            
            elif method == 'minmax':
                for col in numeric_cols:
                    min_val = self.df[col].min()
                    max_val = self.df[col].max()
                    if max_val > min_val:
                        self.df[col] = (self.df[col] - min_val) / (max_val - min_val)
                
                code = """# Min-Max标准化
numeric_cols = df.select_dtypes(include=['number']).columns
for col in numeric_cols:
    min_val = df[col].min()
    max_val = df[col].max()
    if max_val > min_val:
        df[col] = (df[col] - min_val) / (max_val - min_val)"""
            
            self.code_history.append(f"# 数据标准化 - {method}")
            self.code_history.append(code)
            
            message = f"数据标准化完成，使用{method}方法处理了{len(numeric_cols)}个数值列"
            
            return True, message, code
        
        except Exception as e:
            return False, f"数据标准化失败: {str(e)}", ""
    
    def correlation_analysis(self):
        """相关性分析"""
        try:
            numeric_cols = self.df.select_dtypes(include=['number']).columns
            
            if len(numeric_cols) < 2:
                return False, "需要至少2个数值列进行相关性分析", "", False
            
            corr_matrix = self.df[numeric_cols].corr()
            
            code = """# 相关性分析
import pandas as pd
numeric_cols = df.select_dtypes(include=['number']).columns
correlation_matrix = df[numeric_cols].corr()
print(correlation_matrix)"""
            
            self.code_history.append("# 相关性分析")
            self.code_history.append(code)
            
            # 创建相关性结果DataFrame
            self.df = corr_matrix.round(4)
            
            message = f"相关性分析完成，分析了{len(numeric_cols)}个数值列之间的相关性"
            
            return True, message, code, True
        
        except Exception as e:
            return False, f"相关性分析失败: {str(e)}", "", False
    
    def t_test(self, column1, column2=None, value=None):
        """t检验（简化版本，不使用scipy）"""
        try:
            if column1 not in self.df.columns:
                return False, f"列 '{column1}' 不存在", "", False
            
            if not pd.api.types.is_numeric_dtype(self.df[column1]):
                return False, f"列 '{column1}' 不是数值类型", "", False
            
            if column2:  # 双样本t检验
                if column2 not in self.df.columns:
                    return False, f"列 '{column2}' 不存在", "", False
                
                if not pd.api.types.is_numeric_dtype(self.df[column2]):
                    return False, f"列 '{column2}' 不是数值类型", "", False
                
                sample1 = self.df[column1].dropna()
                sample2 = self.df[column2].dropna()
                
                mean1, mean2 = sample1.mean(), sample2.mean()
                var1, var2 = sample1.var(), sample2.var()
                n1, n2 = len(sample1), len(sample2)
                
                # 简化的t统计量计算
                pooled_var = ((n1-1)*var1 + (n2-1)*var2) / (n1+n2-2)
                t_stat = (mean1 - mean2) / math.sqrt(pooled_var * (1/n1 + 1/n2))
                
                code = f"""# 双样本t检验 (简化版本)
import math
sample1 = df['{column1}'].dropna()
sample2 = df['{column2}'].dropna()
mean1, mean2 = sample1.mean(), sample2.mean()
var1, var2 = sample1.var(), sample2.var()
n1, n2 = len(sample1), len(sample2)
pooled_var = ((n1-1)*var1 + (n2-1)*var2) / (n1+n2-2)
t_statistic = (mean1 - mean2) / math.sqrt(pooled_var * (1/n1 + 1/n2))
print(f'T统计量: {{t_statistic:.4f}}')"""
                
                message = f"双样本t检验完成。T统计量: {t_stat:.4f}, 样本1均值: {mean1:.4f}, 样本2均值: {mean2:.4f}"
                
            else:  # 单样本t检验
                if value is None:
                    return False, "单样本t检验需要指定检验值", "", False
                
                sample = self.df[column1].dropna()
                mean = sample.mean()
                std = sample.std()
                n = len(sample)
                
                t_stat = (mean - value) / (std / math.sqrt(n))
                
                code = f"""# 单样本t检验 (简化版本)
import math
sample = df['{column1}'].dropna()
mean = sample.mean()
std = sample.std()
n = len(sample)
test_value = {value}
t_statistic = (mean - test_value) / (std / math.sqrt(n))
print(f'T统计量: {{t_statistic:.4f}}')"""
                
                message = f"单样本t检验完成。T统计量: {t_stat:.4f}, 样本均值: {mean:.4f}, 检验值: {value}"
            
            self.code_history.append("# t检验")
            self.code_history.append(code)
            
            return True, message, code, False
        
        except Exception as e:
            return False, f"t检验失败: {str(e)}", "", False
    
    def chi_square_test(self, column1, column2):
        """卡方检验（简化版本，不使用scipy）"""
        try:
            if column1 not in self.df.columns or column2 not in self.df.columns:
                return False, "指定的列不存在", "", False
            
            # 创建交叉表
            contingency_table = pd.crosstab(self.df[column1], self.df[column2])
            
            # 简化的卡方统计量计算
            row_totals = contingency_table.sum(axis=1)
            col_totals = contingency_table.sum(axis=0)
            total = contingency_table.sum().sum()
            
            chi_square = 0
            for i in range(len(row_totals)):
                for j in range(len(col_totals)):
                    observed = contingency_table.iloc[i, j]
                    expected = (row_totals.iloc[i] * col_totals.iloc[j]) / total
                    if expected > 0:
                        chi_square += (observed - expected) ** 2 / expected
            
            code = f"""# 卡方检验 (简化版本)
contingency_table = pd.crosstab(df['{column1}'], df['{column2}'])
row_totals = contingency_table.sum(axis=1)
col_totals = contingency_table.sum(axis=0)
total = contingency_table.sum().sum()

chi_square = 0
for i in range(len(row_totals)):
    for j in range(len(col_totals)):
        observed = contingency_table.iloc[i, j]
        expected = (row_totals.iloc[i] * col_totals.iloc[j]) / total
        if expected > 0:
            chi_square += (observed - expected) ** 2 / expected

print(f'卡方统计量: {{chi_square:.4f}}')
print(contingency_table)"""
            
            self.code_history.append("# 卡方检验")
            self.code_history.append(code)
            
            # 将交叉表作为结果
            self.df = contingency_table
            
            message = f"卡方检验完成。卡方统计量: {chi_square:.4f}"
            
            return True, message, code, True
        
        except Exception as e:
            return False, f"卡方检验失败: {str(e)}", "", False
    
    def get_complete_code(self):
        """获取完整的Python代码"""
        return "\n".join(self.code_history)
    
    def to_excel(self):
        """导出为Excel"""
        try:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                self.df.to_excel(writer, index=False, sheet_name='processed_data')
            output.seek(0)
            return output
        except Exception as e:
            raise Exception(f"导出Excel失败: {str(e)}")

# 创建全局处理器实例
processor = DataProcessor()

# HTML模板（内联）
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
        .drag-area {
            border: 2px dashed #007bff;
            border-radius: 10px;
            background: #f8f9fa;
            padding: 40px;
            text-align: center;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        .drag-area:hover {
            background: #e9ecef;
            border-color: #0056b3;
        }
        .drag-area.drag-over {
            background: #d4edda;
            border-color: #28a745;
        }
        .btn-custom {
            background: linear-gradient(45deg, #007bff, #6610f2);
            border: none;
            color: white;
            border-radius: 25px;
            padding: 10px 25px;
            transition: all 0.3s ease;
        }
        .btn-custom:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,123,255,0.3);
            color: white;
        }
        .card {
            border-radius: 15px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            border: none;
        }
        .nav-pills .nav-link {
            border-radius: 20px;
            margin: 0 5px;
        }
        .nav-pills .nav-link.active {
            background: linear-gradient(45deg, #007bff, #6610f2);
        }
        .code-display {
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 15px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            white-space: pre-wrap;
            max-height: 400px;
            overflow-y: auto;
        }
        .form-control, .form-select {
            border-radius: 10px;
        }
        .hide-element {
            display: none !important;
        }
    </style>
</head>
<body class="bg-light">
    <div class="container-fluid py-4">
        <div class="row justify-content-center">
            <div class="col-lg-10">
                <div class="text-center mb-4">
                    <h1 class="display-4 text-primary"><i class="fas fa-chart-line"></i> 数据预处理在线工具</h1>
                    <p class="lead text-muted">专业的Excel数据处理平台，支持缺失值处理、异常值检测、数据标准化等功能</p>
                </div>

                <!-- 文件上传区域 -->
                <div class="card mb-4">
                    <div class="card-body">
                        <h5 class="card-title"><i class="fas fa-upload"></i> 上传Excel文件</h5>
                        <div class="drag-area" id="dragArea">
                            <i class="fas fa-cloud-upload-alt fa-3x text-primary mb-3"></i>
                            <h6>拖拽Excel文件到此处，或点击选择文件</h6>
                            <p class="text-muted">支持 .xlsx 和 .xls 格式，最大16MB</p>
                            <input type="file" id="fileInput" accept=".xlsx,.xls" style="display: none;">
                        </div>
                    </div>
                </div>

                <!-- 数据信息展示 -->
                <div class="card mb-4 hide-element" id="dataInfoCard">
                    <div class="card-body">
                        <h5 class="card-title"><i class="fas fa-info-circle"></i> 数据信息</h5>
                        <div id="dataInfo"></div>
                    </div>
                </div>

                <!-- 功能选择区域 -->
                <div class="card mb-4 hide-element" id="functionsCard">
                    <div class="card-body">
                        <h5 class="card-title"><i class="fas fa-cogs"></i> 选择处理功能</h5>
                        
                        <ul class="nav nav-pills mb-4" id="functionTabs" role="tablist">
                            <li class="nav-item" role="presentation">
                                <button class="nav-link active" id="missing-tab" data-bs-toggle="pill" data-bs-target="#missing" type="button">
                                    <i class="fas fa-exclamation-triangle"></i> 缺失值处理
                                </button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" id="outlier-tab" data-bs-toggle="pill" data-bs-target="#outlier" type="button">
                                    <i class="fas fa-search"></i> 异常值处理
                                </button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" id="duplicate-tab" data-bs-toggle="pill" data-bs-target="#duplicate" type="button">
                                    <i class="fas fa-copy"></i> 重复值处理
                                </button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" id="standardization-tab" data-bs-toggle="pill" data-bs-target="#standardization" type="button">
                                    <i class="fas fa-balance-scale"></i> 数据标准化
                                </button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" id="correlation-tab" data-bs-toggle="pill" data-bs-target="#correlation" type="button">
                                    <i class="fas fa-project-diagram"></i> 相关性分析
                                </button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" id="ttest-tab" data-bs-toggle="pill" data-bs-target="#ttest" type="button">
                                    <i class="fas fa-calculator"></i> t检验
                                </button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" id="chisquare-tab" data-bs-toggle="pill" data-bs-target="#chisquare" type="button">
                                    <i class="fas fa-table"></i> 卡方检验
                                </button>
                            </li>
                        </ul>

                        <div class="tab-content" id="functionTabContent">
                            <!-- 缺失值处理 -->
                            <div class="tab-pane fade show active" id="missing" role="tabpanel">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label class="form-label">处理方法：</label>
                                        <select class="form-select" id="missingMethod">
                                            <option value="drop">删除含缺失值的行</option>
                                            <option value="mean">用均值填充</option>
                                            <option value="median">用中位数填充</option>
                                            <option value="value">用指定值填充</option>
                                        </select>
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">填充值：</label>
                                        <input type="text" class="form-control" id="fillValue" placeholder="仅在选择指定值填充时使用" disabled>
                                    </div>
                                </div>
                                <div class="mt-3">
                                    <button class="btn btn-custom" onclick="processMissingValues()">
                                        <i class="fas fa-play"></i> 处理缺失值
                                    </button>
                                </div>
                            </div>

                            <!-- 异常值处理 -->
                            <div class="tab-pane fade" id="outlier" role="tabpanel">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label class="form-label">检测方法：</label>
                                        <select class="form-select" id="outlierMethod">
                                            <option value="iqr">IQR方法</option>
                                            <option value="zscore">Z-score方法</option>
                                        </select>
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">Z-score阈值：</label>
                                        <input type="number" class="form-control" id="zscoreThreshold" value="3" step="0.1" disabled>
                                    </div>
                                </div>
                                <div class="mt-3">
                                    <button class="btn btn-custom" onclick="processOutliers()">
                                        <i class="fas fa-play"></i> 处理异常值
                                    </button>
                                </div>
                            </div>

                            <!-- 重复值处理 -->
                            <div class="tab-pane fade" id="duplicate" role="tabpanel">
                                <p class="text-muted">自动检测并删除完全重复的数据行</p>
                                <button class="btn btn-custom" onclick="processDuplicates()">
                                    <i class="fas fa-play"></i> 删除重复值
                                </button>
                            </div>

                            <!-- 数据标准化 -->
                            <div class="tab-pane fade" id="standardization" role="tabpanel">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label class="form-label">标准化方法：</label>
                                        <select class="form-select" id="standardMethod">
                                            <option value="zscore">Z-score标准化</option>
                                            <option value="minmax">Min-Max标准化</option>
                                        </select>
                                    </div>
                                </div>
                                <div class="mt-3">
                                    <button class="btn btn-custom" onclick="processStandardization()">
                                        <i class="fas fa-play"></i> 数据标准化
                                    </button>
                                </div>
                            </div>

                            <!-- 相关性分析 -->
                            <div class="tab-pane fade" id="correlation" role="tabpanel">
                                <p class="text-muted">计算数值列之间的皮尔逊相关系数矩阵</p>
                                <button class="btn btn-custom" onclick="processCorrelation()">
                                    <i class="fas fa-play"></i> 相关性分析
                                </button>
                            </div>

                            <!-- t检验 -->
                            <div class="tab-pane fade" id="ttest" role="tabpanel">
                                <div class="row">
                                    <div class="col-md-4">
                                        <label class="form-label">检验类型：</label>
                                        <select class="form-select" id="tTestType">
                                            <option value="one_sample">单样本t检验</option>
                                            <option value="two_sample">双样本t检验</option>
                                        </select>
                                    </div>
                                    <div class="col-md-4">
                                        <label class="form-label">列1：</label>
                                        <select class="form-select" id="tTestCol1">
                                            <option value="">请选择列</option>
                                        </select>
                                    </div>
                                    <div class="col-md-4">
                                        <label class="form-label">列2：</label>
                                        <select class="form-select" id="tTestCol2" disabled>
                                            <option value="">请选择列</option>
                                        </select>
                                    </div>
                                </div>
                                <div class="row mt-3">
                                    <div class="col-md-4">
                                        <label class="form-label">检验值：</label>
                                        <input type="number" class="form-control" id="tTestValue" placeholder="单样本检验的检验值" step="any">
                                    </div>
                                </div>
                                <div class="mt-3">
                                    <button class="btn btn-custom" onclick="processTTest()">
                                        <i class="fas fa-play"></i> 执行t检验
                                    </button>
                                </div>
                            </div>

                            <!-- 卡方检验 -->
                            <div class="tab-pane fade" id="chisquare" role="tabpanel">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label class="form-label">列1：</label>
                                        <select class="form-select" id="chiCol1">
                                            <option value="">请选择列</option>
                                        </select>
                                    </div>
                                    <div class="col-md-6">
                                        <label class="form-label">列2：</label>
                                        <select class="form-select" id="chiCol2">
                                            <option value="">请选择列</option>
                                        </select>
                                    </div>
                                </div>
                                <div class="mt-3">
                                    <button class="btn btn-custom" onclick="processChiSquare()">
                                        <i class="fas fa-play"></i> 执行卡方检验
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 处理结果 -->
                <div class="card hide-element" id="resultCard">
                    <div class="card-body">
                        <div class="d-flex justify-content-between align-items-center mb-3">
                            <h5 class="card-title mb-0"><i class="fas fa-check-circle text-success"></i> 处理结果</h5>
                            <button class="btn btn-success" id="downloadBtn" onclick="downloadResult()">
                                <i class="fas fa-download"></i> 下载处理后的文件
                            </button>
                        </div>
                        <div class="alert alert-success" id="resultMessage"></div>
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

            const reader = new FileReader();
            reader.onload = function(e) {
                const fileContent = e.target.result;
                uploadFile(fileContent, file.name);
            };
            reader.readAsDataURL(file);
        }

        // 拖拽功能
        const dragArea = document.getElementById('dragArea');
        
        dragArea.addEventListener('click', () => {
            document.getElementById('fileInput').click();
        });

        dragArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            dragArea.classList.add('drag-over');
        });

        dragArea.addEventListener('dragleave', () => {
            dragArea.classList.remove('drag-over');
        });

        dragArea.addEventListener('drop', (e) => {
            e.preventDefault();
            dragArea.classList.remove('drag-over');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFile(files[0]);
            }
        });

        async function uploadFile(fileContent, filename) {
            try {
                const response = await fetch('/upload', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        file_content: fileContent,
                        filename: filename
                    })
                });

                const data = await response.json();
                
                if (data.success) {
                    currentDataInfo = data.data_info;
                    showDataInfo(data.data_info);
                    showFunctionCards();
                    populateColumnSelects();
                } else {
                    alert('上传失败: ' + data.message);
                }
            } catch (error) {
                alert('上传过程中发生错误: ' + error.message);
            }
        }

        function showDataInfo(info) {
            const html = `
                <div class="row">
                    <div class="col-md-4">
                        <h6>数据维度</h6>
                        <p class="text-primary">${info.shape[0]} 行 × ${info.shape[1]} 列</p>
                    </div>
                    <div class="col-md-4">
                        <h6>列名</h6>
                        <p class="text-secondary">${info.columns.join(', ')}</p>
                    </div>
                    <div class="col-md-4">
                        <h6>缺失值统计</h6>
                        <p class="text-warning">${Object.entries(info.missing_values).map(([col, count]) => count > 0 ? `${col}: ${count}` : '').filter(Boolean).join(', ') || '无缺失值'}</p>
                    </div>
                </div>
            `;
            
            document.getElementById('dataInfo').innerHTML = html;
            document.getElementById('dataInfoCard').classList.remove('hide-element');
        }

        function showFunctionCards() {
            document.getElementById('functionsCard').classList.remove('hide-element');
        }

        function populateColumnSelects() {
            if (!currentDataInfo) return;
            
            const columns = currentDataInfo.columns;
            const selects = ['tTestCol1', 'tTestCol2', 'chiCol1', 'chiCol2'];
            
            selects.forEach(selectId => {
                const select = document.getElementById(selectId);
                select.innerHTML = '<option value="">请选择列</option>';
                columns.forEach(col => {
                    select.innerHTML += `<option value="${col}">${col}</option>`;
                });
            });
        }

        function toggleFillValue() {
            const method = document.getElementById('missingMethod').value;
            const fillValueInput = document.getElementById('fillValue');
            fillValueInput.disabled = method !== 'value';
        }

        function toggleThreshold() {
            const method = document.getElementById('outlierMethod').value;
            const thresholdInput = document.getElementById('zscoreThreshold');
            thresholdInput.disabled = method !== 'zscore';
        }

        function toggleTTestInputs() {
            const testType = document.getElementById('tTestType').value;
            const col2Select = document.getElementById('tTestCol2');
            const valueInput = document.getElementById('tTestValue');
            
            if (testType === 'two_sample') {
                col2Select.disabled = false;
                valueInput.disabled = true;
            } else {
                col2Select.disabled = true;
                valueInput.disabled = false;
            }
        }

        async function processData(operation, parameters) {
            try {
                const response = await fetch('/process', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        operation: operation,
                        parameters: parameters
                    })
                });

                const data = await response.json();
                showResult(data);
            } catch (error) {
                alert('处理过程中发生错误: ' + error.message);
            }
        }

        function processMissingValues() {
            const method = document.getElementById('missingMethod').value;
            const fillValue = document.getElementById('fillValue').value;
            
            const params = { method: method };
            if (method === 'value' && fillValue) {
                params.fill_value = isNaN(fillValue) ? fillValue : parseFloat(fillValue);
            }
            
            processData('missing_values', params);
        }

        function processOutliers() {
            const method = document.getElementById('outlierMethod').value;
            const threshold = parseFloat(document.getElementById('zscoreThreshold').value);
            
            processData('outliers', { method: method, threshold: threshold });
        }

        function processDuplicates() {
            processData('duplicates', {});
        }

        function processStandardization() {
            const method = document.getElementById('standardMethod').value;
            processData('standardization', { method: method });
        }

        function processCorrelation() {
            processData('correlation', {});
        }

        function processTTest() {
            const testType = document.getElementById('tTestType').value;
            const col1 = document.getElementById('tTestCol1').value;
            const col2 = document.getElementById('tTestCol2').value;
            const value = document.getElementById('tTestValue').value;
            
            if (!col1) {
                alert('请选择列1');
                return;
            }
            
            const params = { column1: col1 };
            
            if (testType === 'two_sample') {
                if (!col2) {
                    alert('请选择列2');
                    return;
                }
                params.column2 = col2;
            } else {
                if (!value) {
                    alert('请输入检验值');
                    return;
                }
                params.value = parseFloat(value);
            }
            
            processData('t_test', params);
        }

        function processChiSquare() {
            const col1 = document.getElementById('chiCol1').value;
            const col2 = document.getElementById('chiCol2').value;
            
            if (!col1 || !col2) {
                alert('请选择两个列');
                return;
            }
            
            processData('chi_square', { column1: col1, column2: col2 });
        }

        function showResult(data) {
            if (data.success) {
                document.getElementById('resultMessage').textContent = data.message;
                document.getElementById('pythonCode').textContent = data.code;
                
                const downloadBtn = document.getElementById('downloadBtn');
                if (data.can_download) {
                    downloadBtn.style.display = 'block';
                } else {
                    downloadBtn.style.display = 'none';
                }
                
                document.getElementById('resultCard').classList.remove('hide-element');
                document.getElementById('resultCard').scrollIntoView({ behavior: 'smooth' });
            } else {
                alert('处理失败: ' + data.message);
            }
        }

        async function downloadResult() {
            try {
                const response = await fetch('/download');
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = 'processed_data.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                } else {
                    alert('下载失败');
                }
            } catch (error) {
                alert('下载过程中发生错误: ' + error.message);
            }
        }

        function copyCode() {
            const codeText = document.getElementById('pythonCode').textContent;
            navigator.clipboard.writeText(codeText).then(() => {
                alert('代码已复制到剪贴板');
            }).catch(() => {
                alert('复制失败，请手动选择代码');
            });
        }

        async function resetProcessor() {
            try {
                await fetch('/reset', { method: 'POST' });
                location.reload();
            } catch (error) {
                location.reload();
            }
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
            'code': complete_code,
            'can_download': can_download
        })
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'处理失败: {str(e)}'})

@app.route('/download')
def download_file():
    try:
        excel_file = processor.to_excel()
        return send_file(
            excel_file,
            as_attachment=True,
            download_name='processed_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': f'下载失败: {str(e)}'}), 500

@app.route('/reset', methods=['POST'])
def reset_processor():
    global processor
    processor = DataProcessor()
    return jsonify({'success': True, 'message': '已重置，可以上传新文件'})

# Vercel需要这个入口点
app_instance = app

if __name__ == '__main__':
    app.run(debug=True)
