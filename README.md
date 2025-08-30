# 数据预处理在线工具

一个功能强大的数据预处理Web应用，支持Excel文件的在线处理和分析。

## ✨ 主要功能

### 📊 数据清洗
- **缺失值处理**：删除、均值填充、中位数填充、指定值填充
- **异常值处理**：IQR方法、Z-score方法
- **重复值处理**：自动检测并删除重复行

### 📏 数据标准化
- **Z-score标准化**：(x - μ) / σ
- **Min-Max标准化**：(x - min) / (max - min)

### 📈 统计分析
- **相关性分析**：计算数值列之间的相关系数矩阵
- **t检验**：单样本t检验、双样本t检验
- **卡方检验**：测试两个分类变量之间的独立性

## 🚀 快速开始

### 部署到Vercel

1. 创建新的GitHub仓库
2. 上传所有文件到仓库
3. 在Vercel中连接GitHub仓库
4. 自动部署完成

### 本地运行

```bash
pip install -r requirements.txt
python api/app.py
```

## 💻 使用方法

1. 上传Excel文件（.xlsx或.xls格式）
2. 查看数据基本信息和预览
3. 选择数据处理功能
4. 配置相关参数
5. 执行处理并查看结果
6. 下载处理后的Excel文件
7. 复制完整的Python代码

## 🛠️ 技术栈

- **后端**：Flask, pandas, numpy, scikit-learn, scipy
- **前端**：Bootstrap 5, JavaScript
- **部署**：Vercel

## 📁 项目结构

```
vercel_ready/
├── api/
│   └── app.py              # Flask应用主文件
├── vercel.json             # Vercel配置
├── requirements.txt        # Python依赖
└── README.md              # 项目说明
```

## 🌐 特性

- 🎨 现代化的渐变色UI设计
- 📱 响应式界面，支持移动端
- 🚀 一键部署到Vercel
- 💾 支持Excel文件上传和下载
- 🔍 实时数据预览和统计信息
- 📊 完整的Python代码生成
- ⚡ 快速的数据处理性能

## 📝 支持的文件格式

- Excel文件：.xlsx, .xls
- 文件大小限制：16MB
- 数据要求：第一行应为列标题

## 🎯 适用场景

- 数据科学项目的预处理阶段
- 统计分析前的数据清洗
- 教学和学习数据处理方法
- 快速的数据质量检查
- 生成可复现的数据处理代码

## ⚠️ 注意事项

- 上传的文件会在服务器端临时处理，不会永久存储
- 免费版Vercel有运行时间和带宽限制
- 建议使用现代浏览器以获得最佳体验

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 📄 许可证

MIT License
