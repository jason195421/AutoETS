# pip install python-docx
import os
import json
import re
import random
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_LINE_SPACING
from datetime import datetime

# 路径配置
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
appdata_path = os.getenv('APPDATA')
ets_path = os.path.join(appdata_path, 'ETS')
program_files_path = os.getenv('ProgramFiles(x86)')
ets_program_path = os.path.join(program_files_path, 'ETS') if program_files_path else None

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def check_ets_status():
    appdata_exists = os.path.exists(ets_path)
    program_exists = ets_program_path and os.path.exists(ets_program_path)
    return "ETS安装 ✔" if (appdata_exists or program_exists) else "ETS安装 ✘"

def show_menu():
    clear_screen()
    ets_status = check_ets_status()
    print("\n" + "="*50)
    print("         ETS NEXT  ")
    print("="*50)
    print(f"状态: {ets_status}")
    print("="*50)
    print("1. 开始解析")
    print("2. 关于")
    print("3. 使用方法")
    print("0. 退出")
    print("="*50)
    
    # 随机Phi风格Tips
    tips_list = [
        "咕咕咕～解析前记得给设备留口呼吸哦～",
        "解析成功=AP判定！失败=Miss，不过咕咕们会原谅你的～",
        "选对试卷=FC，选错也别怕，重新选就好～",
        "解析时不要关闭程序，否则会像断触一样前功尽弃～",
        "乱改content_xx命名=解析判定Miss！听鸽游一句劝～",
        "解析完成后记得查看桌面的文档，别让它像隐藏谱面一样难找～"
    ]
    print(f"Tips: {random.choice(tips_list)}")

def show_about():
    clear_screen()
    print("\n" + "="*50)
    print("                 关于")
    print("="*50)
    print("Ets  NEXT ")
    print("版本: 1.2 Beta")
    print("功能: 自动解析E听说试卷并生成Word文档")
    print("作者: jason ff ")
    print("Tips: all right reserved     ")
    print("="*50)
    input("按回车键返回主菜单...")

def show_usage():
    clear_screen()
    print("\n" + "="*50)
    print("               使用方法")
    print("="*50)
    print("选择试卷：程序启动后，会列出所有下载的试卷，按照")
    print("          下载时间从新到旧排序。用户可以输入试卷")
    print("          对应的编号来选择要解析的试卷。")
    print("")
    print("生成解析文档：程序会自动解析选定试卷的内容，并生成")
    print("            解析文档。解析文档会保存到用户的桌面，")
    print("            文档名称为 E听说_解析.docx。如果文件")
    print("            已经存在，程序会在文件名后添加序号，")
    print("            避免覆盖已有文档。")
    print("")
    print("文档格式：")
    print("  - 大标题（如 'Section A'、'朗读句子' 等）为绿色")
    print("    (RGB 值为 #00B050)，中文使用等线字体，英文")
    print("    使用 Times New Roman。")
    print("  - Section A 和 Section B 的答案部分以蓝色标注，")
    print("    且加粗'答案'二字。")
    print("  - 图片描述中若存在图片，程序会自动将图片插入到")
    print("    文档中。")
    print("="*50)
    input("按回车键返回主菜单...")

def get_sorted_content_folders(folder_path):
    content_folders = [f for f in os.listdir(folder_path) if f.startswith('content_')]
    content_folders.sort(key=lambda x: int(x.split('_')[1]))
    return content_folders

def clean_html(raw_html):
    raw_html = raw_html.replace('<p>', '[NEWLINE]').replace('</p>', '[NEWLINE]')
    raw_html = raw_html.replace('<br>', '[NEWLINE]').replace('<br', '[NEWLINE]').replace('</br>', '[NEWLINE]')
    clean_text = re.sub('<.*?>', '', raw_html)
    clean_text = clean_text.replace('[NEWLINE]', '\n')
    clean_text = re.sub(r'\n+', '\n', clean_text).strip()
    return clean_text

def parse_section_from_folder(folder_path, content_folder_name, start_number):
    content_folder = os.path.join(folder_path, content_folder_name)
    json_file_path = os.path.join(content_folder, 'content2.json')

    if not os.path.exists(json_file_path):
        print(f"文件 {json_file_path} 不存在，跳过此文件夹。")
        return "", start_number

    with open(json_file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    questions = data['info']['xtlist']
    parsed_questions = []

    for question in questions:
        question_text = clean_html(question['xt_value'])
        if not question_text.strip().startswith(str(start_number) + '.'):
            question_text = f"{start_number}. {question_text}"

        options = question['xxlist']
        answers = "\n".join([f"{opt['xx_mc']}. {clean_html(opt['xx_nr'])}" for opt in options])
        correct_answer = question['answer']
        parsed_questions.append(f"{question_text}\n{answers}\n答案：{correct_answer}\n\n")
        start_number += 1

    return "\n".join(parsed_questions), start_number

def parse_section_a(folder_path):
    content_folders = get_sorted_content_folders(folder_path)
    start_number = 1
    section_a_content = ""
    # 移除内容完整性检测，直接遍历可用文件夹
    for i in [0, 1]:
        if i < len(content_folders):
            part_content, start_number = parse_section_from_folder(folder_path, content_folders[i], start_number)
            section_a_content += part_content
    return section_a_content

def parse_section_b_with_reading(folder_path):
    content_folders = get_sorted_content_folders(folder_path)
    start_number = 1
    section_b_content = ""
    # 移除内容完整性检测，直接遍历可用文件夹
    for i in range(2, len(content_folders)):
        content_folder = content_folders[i]
        part_content, start_number = parse_section_from_folder(folder_path, content_folder, start_number)
        section_b_content += part_content
    return section_b_content

def parse_image_descriptions(folder_path):
    content_folders = get_sorted_content_folders(folder_path)
    image_content = ""
    
    for folder in content_folders:
        folder_path_full = os.path.join(folder_path, folder)
        json_path = os.path.join(folder_path_full, 'content2.json')
        
        if not os.path.exists(json_path):
            continue
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        questions = [q for q in data['info']['xtlist'] if q.get('xt_type') == 6]
        if not questions:
            continue
        
        image_content += "### 图片描述题\n\n"
        for idx, q in enumerate(questions, 1):
            q_text = clean_html(q['xt_value'])
            image_content += f"{idx}. {q_text}\n"
            
            image_files = [f for f in os.listdir(folder_path_full) if f.endswith(('.jpg', '.png', '.jpeg'))]
            if image_files:
                image_content += f"[图片位置：{os.path.join(folder_path_full, image_files[0])}]\n"
            
            correct_answer = q.get('answer', '无标准答案')
            image_content += f"答案：{correct_answer}\n\n"
    
    return image_content

def generate_word_document(paper_name, section_a, section_b, image_desc):
    doc = Document()
    
    # 文档基础样式
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(12)
    doc.styles['Normal'].paragraph_format.line_spacing = WD_LINE_SPACING.SINGLE
    doc.styles['Normal'].font.color.rgb = RGBColor(0, 0, 0)
    for style in doc.styles:
        if style.font.name:
            style.font.element.rPr.rFonts.set(qn('w:eastAsia'), '等线')
    
    # 标题与时间
    title = doc.add_heading(f'E听说试卷解析 - {paper_name}', 0)
    title.alignment = 1
    title.font.color.rgb = RGBColor(0, 176, 80)
    title.font.name = '等线'
    title.font.element.rPr.rFonts.set(qn('w:eastAsia'), '等线')
    
    time_paragraph = doc.add_paragraph(f'解析时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    time_paragraph.alignment = 2
    
    doc.add_paragraph('='*50)
    
    # Section A
    a_heading = doc.add_heading('Section A', level=1)
    a_heading.font.color.rgb = RGBColor(0, 176, 80)
    a_heading.font.name = 'Times New Roman'
    a_heading.font.element.rPr.rFonts.set(qn('w:eastAsia'), '等线')
    
    a_paragraph = doc.add_paragraph()
    for line in section_a.split('\n'):
        if line.startswith('答案：'):
            run = a_paragraph.add_run(line)
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.bold = True
        else:
            a_paragraph.add_run(line)
        a_paragraph.add_run('\n')
    
    doc.add_page_break()
    
    # Section B
    b_heading = doc.add_heading('Section B', level=1)
    b_heading.font.color.rgb = RGBColor(0, 176, 80)
    b_heading.font.name = 'Times New Roman'
    b_heading.font.element.rPr.rFonts.set(qn('w:eastAsia'), '等线')
    
    b_paragraph = doc.add_paragraph()
    for line in section_b.split('\n'):
        if line.startswith('答案：'):
            run = b_paragraph.add_run(line)
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.bold = True
        else:
            b_paragraph.add_run(line)
        b_paragraph.add_run('\n')
    
    doc.add_page_break()
    
    # 图片描述题
    if image_desc:
        img_heading = doc.add_heading('图片描述题', level=1)
        img_heading.font.color.rgb = RGBColor(0, 176, 80)
        img_heading.font.name = 'Times New Roman'
        img_heading.font.element.rPr.rFonts.set(qn('w:eastAsia'), '等线')
        
        img_paragraph = doc.add_paragraph()
        for line in image_desc.split('\n'):
            if line.startswith('答案：'):
                run = img_paragraph.add_run(line)
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.bold = True
            elif '[图片位置：' in line:
                img_path = re.findall(r'\[图片位置：(.*?)\]', line)[0]
                if os.path.exists(img_path):
                    doc.add_picture(img_path, width=Inches(5))
                    img_paragraph.add_run('\n（图片已插入）\n')
                else:
                    img_paragraph.add_run('（图片文件不存在，无法插入）\n')
            else:
                img_paragraph.add_run(line)
            img_paragraph.add_run('\n')
    
    # 保存文档（避免覆盖）
    doc_name = 'E听说_解析.docx'
    doc_path = os.path.join(desktop_path, doc_name)
    counter = 1
    while os.path.exists(doc_path):
        doc_path = os.path.join(desktop_path, f'E听说_解析_{counter}.docx')
        counter += 1
    
    doc.save(doc_path)
    return doc_path

def select_paper():
    papers_path = os.path.join(ets_path, 'papers')
    if not os.path.exists(papers_path):
        print(f"未找到试卷目录：{papers_path}")
        print("请确保E听说已下载试卷到默认路径")
        input("按回车键返回主菜单...")
        return None
    
    paper_folders = [f for f in os.listdir(papers_path) if os.path.isdir(os.path.join(papers_path, f))]
    if not paper_folders:
        print("未找到已下载的试卷")
        input("按回车键返回主菜单...")
        return None
    
    # 按修改时间排序（新到旧）
    paper_folders.sort(key=lambda x: os.path.getmtime(os.path.join(papers_path, x)), reverse=True)
    
    clear_screen()
    print("\n" + "="*50)
    print("               已下载试卷列表")
    print("="*50)
    for idx, folder in enumerate(paper_folders, 1):
        modify_time = datetime.fromtimestamp(os.path.getmtime(os.path.join(papers_path, folder))).strftime("%Y-%m-%d")
        print(f"{idx}. {folder} （下载时间：{modify_time}）")
    print("0. 返回主菜单")
    print("="*50)
    
    while True:
        try:
            choice = input("请输入试卷编号：")
            if choice == '0':
                return None
            choice_idx = int(choice) - 1
            if 0 <= choice_idx < len(paper_folders):
                selected_folder = paper_folders[choice_idx]
                return os.path.join(papers_path, selected_folder), selected_folder
            else:
                print("输入错误，请输入有效的编号")
        except ValueError:
            print("输入错误，请输入数字")

def start_parse():
    clear_screen()
    print("\n" + "="*50)
    print("               开始解析试卷")
    print("="*50)
    
    paper_path, paper_name = select_paper() or (None, None)
    if not paper_path:
        return
    
    print(f"\n正在解析试卷：{paper_name}...")
    section_a = parse_section_a(paper_path)
    print("Section A 解析完成")
    section_b = parse_section_b_with_reading(paper_path)
    print("Section B 解析完成")
    image_desc = parse_image_descriptions(paper_path)
    print("图片描述题解析完成")
    
    print("正在生成Word文档...")
    doc_path = generate_word_document(paper_name, section_a, section_b, image_desc)
    
    print(f"\n解析完成！文档保存路径：{doc_path}")
    input("按回车键返回主菜单...")

def main():
    while True:
        show_menu()
        choice = input("请输入功能编号：")
        if choice == '1':
            start_parse()
        elif choice == '2':
            show_about()
        elif choice == '3':
            show_usage()
        elif choice == '0':
            clear_screen()
            print("感谢使用，再见！")
            break
        else:
            input("输入错误，请输入有效的功能编号，按回车键继续...")

if __name__ == "__main__":
    main()
