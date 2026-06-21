import json
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import pdfplumber  # 记得添加这个导入

def read_pdf(pdf_path):
    """读取PDF文件并返回文本内容"""
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text

def chunk_text(text, max_chars=3000):
    """将长文本按最大字符数分块"""
    chunks = []
    for i in range(0, len(text), max_chars):
        chunks.append(text[i:i + max_chars])
    return chunks

# ==================== 1. API 配置 ====================
api_key = "xxx"  # 请替换成你自己的密钥，不要提交真实密钥
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
        response_format={'type': 'json_object'},
        stream=False
    )
    return response.choices[0].message.content


# ==================== 3. PPT 生成 ====================
def create_ppt(data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # ===== 幻灯片1：标题页（使用标题布局）=====
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide1.shapes.title
    title.text = data.get("title", "AI+硬件协同设计愿景简报")
    # 如果有副标题占位符，也可以填充
    if len(slide1.placeholders) > 1:
        subtitle = slide1.placeholders[1]
        subtitle.text = "基于论文：AI+HV2035: Shaping the Next Decade"

    # ===== 幻灯片2：三大层级（卡片式布局）=====
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局

    # 页面标题
    title2 = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
    title2.text_frame.text = "三大协同设计层级"
    title2.text_frame.paragraphs[0].font.size = Pt(32)
    title2.text_frame.paragraphs[0].font.bold = True

    # 检查数据中有没有 three_levels 字段
    if "three_levels" in data and len(data["three_levels"]) >= 3:
        levels = data["three_levels"]
        # 定义三张卡片的位置和颜色
        cards = [
            {"x": 0.5, "color": RGBColor(0, 112, 192), "title": "Hardware"},
            {"x": 4.8, "color": RGBColor(0, 176, 80), "title": "Algorithm"},
            {"x": 9.1, "color": RGBColor(237, 125, 49), "title": "Application"}
        ]
        for i, card in enumerate(cards):
            # 画卡片背景（圆角矩形）
            shape = slide2.shapes.add_shape(
                1,  # 矩形
                Inches(card["x"]), Inches(1.5),
                Inches(3.8), Inches(3.5)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = card["color"]
            shape.line.color.rgb = card["color"]

            # 层级名称（白色大号居中）
            name_box = slide2.shapes.add_textbox(
                Inches(card["x"] + 0.2), Inches(1.7),
                Inches(3.4), Inches(0.7)
            )
            name_box.text_frame.text = card["title"]
            name_box.text_frame.paragraphs[0].font.size = Pt(24)
            name_box.text_frame.paragraphs[0].font.bold = True
            name_box.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            name_box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

            # 特征描述（白色，项目符号）
            feature_text = levels[i].get("features", "")
            # 尝试把中文逗号变成换行
            feature_items = [item.strip() for item in feature_text.replace("，", ",").split(",")]

            feature_box = slide2.shapes.add_textbox(
                Inches(card["x"] + 0.3), Inches(2.6),
                Inches(3.2), Inches(2.2)
            )
            tf = feature_box.text_frame
            tf.word_wrap = True
            tf.clear()
            for idx, item in enumerate(feature_items[:4]):  # 最多显示4条
                p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                p.text = f"• {item}"
                p.font.size = Pt(14)
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.space_after = Pt(6)
    else:
        # 如果没有 three_levels，显示提示
        fallback_box = slide2.shapes.add_textbox(Inches(1), Inches(2), Inches(10), Inches(2))
        fallback_box.text_frame.text = "未提取到三大层级数据，请检查论文内容或提示词。"
        fallback_box.text_frame.paragraphs[0].font.size = Pt(20)

    # ===== 幻灯片3：时间轴 =====
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    title3 = slide3.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
    title3.text_frame.text = "技术发展时间轴"
    title3.text_frame.paragraphs[0].font.size = Pt(32)
    title3.text_frame.paragraphs[0].font.bold = True

    # 近期（左）
    near_box = slide3.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5.5), Inches(4))
    tf = near_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = "近期（2-5年）"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 112, 192)
    for item in data.get("timeline", {}).get("near_term_2_5_years", []):
        p = tf.add_paragraph()
        p.text = f"• {item}"
        p.font.size = Pt(18)
        p.level = 1
        p.space_after = Pt(8)

    # 中间箭头（指向未来）
    arrow = slide3.shapes.add_shape(
        13,  # 右箭头
        Inches(6.2), Inches(3.0),
        Inches(0.8), Inches(0.4)
    )
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = RGBColor(192, 0, 0)
    arrow.line.color.rgb = RGBColor(192, 0, 0)

    # 远期（右）
    far_box = slide3.shapes.add_textbox(Inches(7.5), Inches(1.5), Inches(5.5), Inches(4))
    tf = far_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = "远期（6-10年）"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(237, 125, 49)
    for item in data.get("timeline", {}).get("far_term_6_10_years", []):
        p = tf.add_paragraph()
        p.text = f"• {item}"
        p.font.size = Pt(18)
        p.level = 1
        p.space_after = Pt(8)

    prs.save("briefing.pptx")
    print("✅ PPT 已生成：briefing.pptx")

# ==================== 4. 主程序 ====================
if __name__ == "__main__":
    print("=== Day 2: 提取完整结构化信息 ===")

    # 1. 读PDF
    print("正在读论文...")
    full_text = read_pdf("paper.pdf")
    print(f"读到了 {len(full_text)} 字符")

    # 2. 分块并合并前两块（确保覆盖2.2节）
    chunks = chunk_text(full_text, max_chars=3000)
    # 取前两块合并，约6000字符（足够包含摘要 + 2.2节开头）
    first_part = chunks[0] + "\n" + (chunks[1] if len(chunks) > 1 else "")
    print(f"喂给AI的文本长度：{len(first_part)} 字符")

    # 3. 构造prompt
    prompt = f"""
    你是一个资深技术分析师。请从以下论文内容中提取结构化信息，并以JSON格式返回。

    论文内容：
    {first_part}

    请提取以下内容：
    1. "title": 论文标题（英文原样）
    2. "core_goal": 用一句中文描述"1000倍效率提升"的定义
    3. "three_levels": 论文2.2节的三个抽象层级（Hardware / Algorithm / Application）及其特征
    4. "timeline": 近期（2-5年）和远期（6-10年）的技术趋势

    请严格按照以下JSON格式返回：
    {{
        "title": "论文标题",
        "core_goal": "1000倍效率提升的定义",
        "three_levels": [
            {{"level": "Hardware", "features": "特征描述"}},
            {{"level": "Algorithm", "features": "特征描述"}},
            {{"level": "Application", "features": "特征描述"}}
        ],
        "timeline": {{
            "near_term_2_5_years": ["技术1", "技术2"],
            "far_term_6_10_years": ["技术1", "技术2"]
        }}
    }}
    """

    # 4. 调用AI
    print("正在呼叫DeepSeek提取完整信息...")
    try:
        reply = call_ai(prompt)
        print("AI原始返回：", reply)

        # 5. 解析JSON
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