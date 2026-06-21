import os
import json
import re
import pdfplumber
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from openai import OpenAI


# ==================== 1. 初始化 DeepSeek ====================
api_key = "你的 DeepSeek API Key"   # 请替换成你自己的密钥，不要提交真实密钥

client = OpenAI(
    api_key=api_key,
    base_url="https://api.deepseek.com"
)


# ==================== 2. AI 调用 ====================
def call_ai(user_content, system_content=None):
    messages = []
    if system_content:
        messages.append({"role": "system", "content": system_content})
    messages.append({"role": "user", "content": user_content})

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=messages,
        response_format={'type': 'json_object'},  # 强制 JSON 输出，省去清洗麻烦
        stream=False
    )
    return response.choices[0].message.content


# ==================== 3. PDF 读取 ====================
def read_pdf(pdf_path):
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text


# ==================== 4. 文本分块（按字符硬切） ====================
def chunk_text(text, max_chars=3000):
    chunks = []
    for i in range(0, len(text), max_chars):
        chunks.append(text[i:i + max_chars])
    return chunks

def create_ppt(data):
    """将提取的 JSON 数据生成 PPT"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    # ===== 第一页：标题 + 核心愿景 =====
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide1.shapes.add_textbox(Inches(1), Inches(0.8), Inches(8), Inches(1.2))
    title_box.text_frame.text = "AI+硬件协同设计愿景简报"
    title_box.text_frame.paragraphs[0].font.size = Pt(36)
    title_box.text_frame.paragraphs[0].font.bold = True
    title_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    goal_box = slide1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(2))
    goal_box.text_frame.word_wrap = True
    goal_box.text_frame.text = f"核心目标：\n{data['core_goal']}"
    goal_box.text_frame.paragraphs[0].font.size = Pt(20)

    footer = slide1.shapes.add_textbox(Inches(1), Inches(5), Inches(8), Inches(0.5))
    footer.text_frame.text = "基于论文：AI+HV2035: Shaping the Next Decade (arXiv:2603.05225)"
    footer.text_frame.paragraphs[0].font.size = Pt(12)

    # ===== 第二页：三大层级表格 =====
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    title2 = slide2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    title2.text_frame.text = "三大协同设计层级"
    title2.text_frame.paragraphs[0].font.size = Pt(28)
    title2.text_frame.paragraphs[0].font.bold = True

    table = slide2.shapes.add_table(4, 2, Inches(1.5), Inches(1.8), Inches(7), Inches(2.5)).table
    table.columns[0].width = Inches(2.5)
    table.columns[1].width = Inches(4.5)

    table.cell(0, 0).text = "抽象层级"
    table.cell(0, 1).text = "关键特征"
    for i, level_info in enumerate(data["three_levels"], start=1):
        level_name = level_info["level"].split(":")[0].strip()
        table.cell(i, 0).text = level_name
        table.cell(i, 1).text = level_info["features"]

    # ===== 第三页：时间轴 =====
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    title3 = slide3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    title3.text_frame.text = "技术发展时间轴"
    title3.text_frame.paragraphs[0].font.size = Pt(28)
    title3.text_frame.paragraphs[0].font.bold = True

    near_box = slide3.shapes.add_textbox(Inches(1), Inches(1.8), Inches(4), Inches(2.5))
    near_text = "近期（2-5年）\n" + "\n".join([f"• {tech}" for tech in data["timeline"]["near_term_2_5_years"]])
    near_box.text_frame.text = near_text
    near_box.text_frame.paragraphs[0].font.size = Pt(18)

    far_box = slide3.shapes.add_textbox(Inches(5.5), Inches(1.8), Inches(4), Inches(2.5))
    far_text = "远期（6-10年）\n" + "\n".join([f"• {tech}" for tech in data["timeline"]["far_term_6_10_years"]])
    far_box.text_frame.text = far_text
    far_box.text_frame.paragraphs[0].font.size = Pt(18)

    prs.save("briefing.pptx")
    print("✅ PPT 已生成：briefing.pptx")
# ==================== 5. 主程序（Day 2） ====================
if __name__ == "__main__":
    print("=== Day 2: 提取完整结构化信息 ===")

    # 1. 读 PDF
    print("正在读论文...")
    full_text = read_pdf("paper.pdf")
    print(f"读到了 {len(full_text)} 字符")

    # 2. 分块并合并前两块（确保覆盖 2.2 节）
    chunks = chunk_text(full_text, max_chars=3000)
    # 取前两块合并，约 6000 字符（足够包含摘要 + 2.2 节开头）
    first_part = chunks[0] + "\n" + (chunks[1] if len(chunks) > 1 else "")
    print(f"喂给 AI 的文本长度：{len(first_part)} 字符")

    # 3. 核心提示词（完全按照作业要求设计）
    prompt = f"""
你是一位资深技术分析师。请从以下论文内容中提取三类关键信息，并严格按 JSON 格式输出。

论文内容（开头部分）：
{first_part}

请提取以下内容：
1. "core_goal"：用一句中文描述论文提出的 "1000倍效率提升" 的具体定义（Intelligence per Joule）。
2. "three_levels"：提取论文第 2.2 节提到的三个抽象层级（Hardware, Algorithm, Application）及其关键特征（每个特征用 20 字以内概括）。
3. "timeline"：提取 "近期（2-5年）" 和 "远期（6-10年）" 的关键技术趋势（各列出 2-3 项，如 3D Integration, Photonics 等）。

输出格式必须严格符合以下 JSON 结构，不要带任何额外文字：
{{
    "core_goal": "字符串",
    "three_levels": [
        {{"level": "Hardware Layer: Hardware Technologies. ", "features": "特征描述"}},
        {{"level": "Algorithm Layer: Algorithms and Paradigms.", "features": "特征描述"}},
        {{"level": "Application Layer: Applications and Societal Impact.", "features": "特征描述"}}
    ],
    "timeline": {{
        "near_term_2_5_years": ["技术1", "技术2"],
        "far_term_6_10_years": ["技术1", "技术2"]
    }}
}}
"""

    # 4. 调用 AI
    print("正在呼叫 DeepSeek 提取完整信息...")
    try:
        reply = call_ai(prompt)
        print("AI 原始返回：", reply)

        # 5. 解析 JSON（由于强制了 json_object，此处直接 loads 即可）
        data = json.loads(reply)

        # 打印漂亮的结果
        print("\n" + "=" * 50)
        print("✅ 提取成功！结果如下：")
        print(json.dumps(data, ensure_ascii=False, indent=2))
        create_ppt(data)
        print("=" * 50)

    except Exception as e:
        print(f"❌ 出错：{e}")
        print("原始返回内容：", reply if 'reply' in locals() else "无返回")