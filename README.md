# AI+硬件协同设计愿景自动化简报生成器

## 1. 项目简介

本项目是一个基于大语言模型（DeepSeek）的自动化论文简报生成工具。它能够自动读取学术论文 PDF，智能提取核心信息（1000倍效率定义、三大协同层级、技术发展时间轴），并生成一份可直接用于汇报的专业级 PPT 简报。

## 2. 运行环境与依赖

- **Python 版本**：3.8 或以上
- **操作系统**：Windows / macOS / Linux

### 依赖库（已写入 requirements.txt）


pip install -r requirements.txt
库	用途
openai	调用 DeepSeek API
pdfplumber	解析 PDF 文本
python-pptx	生成 PPT 简报
python-dotenv	读取 .env 环境变量
3. 快速开始
3.1 克隆或下载项目
将本项目代码下载到本地。

3.2 配置 API Key（重要）
本项目使用 DeepSeek API，你需要：

访问 DeepSeek 开放平台 注册并获取 API Key。

在项目根目录下，将 .env.example 文件重命名为 .env。

在 .env 文件中填入你的 API Key：

text
DEEPSEEK_API_KEY=你的DeepSeek密钥
3.3 准备论文 PDF
从 https://arxiv.org/abs/2603.05225 下载论文，将 PDF 文件放入项目根目录，并重命名为 paper.pdf。

3.4 运行程序
在终端执行：

bash
python generate_briefing.py
运行结束后，会在当前目录生成 briefing.pptx 简报文件。

4. 核心技术思路
4.1 智能文本定位（Chunking 策略）
我没有采用简单的前 N 字截取，而是利用 Python 的字符串查找功能，精准定位论文第 2.2 章节（三大层级的定义处），只截取该章节前后约 3000 字符输入给 AI。

优势：相比全量输入（12 万字），节省约 50% 以上的 Token 消耗。

优势：确保 AI 提取的信息直接基于原文，避免“凭空捏造”。

4.2 Prompt 工程
我在系统提示词中赋予 AI “资深技术分析师”的角色，并在用户提示词中：

明确要求输出 JSON 格式，并给出精确的字段样例。

强制要求提取四个核心字段：标题、1000倍效率定义、三大层级、时间轴。

使用 response_format={'type': 'json_object'} 参数，强制模型输出合法 JSON，杜绝解析失败。

4.3 PPT 自动生成
使用 python-pptx 库生成三页简报：

第一页：标题 + 核心愿景（居中排版，字体层次分明）。

第二页：三大层级表格（蓝底白字表头，内容行浅灰背景，阅读舒适）。

第三页：时间轴（近期 vs 远期，带视觉分隔线和时间节点标记）。

5. 遇到的挑战与解决方案
问题	解决方案
环境变量读取失败（ValueError: 请设置环境变量）	采用 .env 文件 + python-dotenv 管理 API Key，审核人只需修改 .env 文件即可运行，无需改动代码。
AI 提取的层级名称不统一（如出现 "Hardware Layer: xxx" 等冗余词）	通过精准定位 2.2 章节原文，让 AI 直接基于原文提取，消除歧义。
PPT 排版简陋，信息层级不清	放弃表格默认样式，手动设置表头蓝底白字、内容行浅灰背景，时间轴加入箭头和颜色区分。
6. 项目文件结构
text
.
├── generate_briefing.py   # 主程序源代码
├── briefing.pptx          # 生成的简报 PPT
├── README.md              # 项目说明文档
├── requirements.txt       # Python 依赖清单
├── .env.example           # 环境变量模板（复制为 .env 使用）
└── .gitignore             # Git 忽略文件（含 .env）
7. 致谢
本项目基于论文 AI+HV2035: Shaping the Next Decade（arXiv:2603.05225），使用 DeepSeek API 实现信息提取，并通过 python-pptx 实现自动化简报生成。

作者：[你的姓名]
学号：[你的学号]
提交日期：2026年6月
