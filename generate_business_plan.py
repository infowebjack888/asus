#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FutureEdge AI Hub v3.0 企劃書生成腳本
Aether (流光以太) — Infoweb 2.0 教育訓練事業部
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

doc = Document()

# ─── 頁面設定 ────────────────────────────────────────────
section = doc.sections[0]
section.page_width  = Cm(21)
section.page_height = Cm(29.7)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)

# ─── 顏色常量 ────────────────────────────────────────────
DARK_BLUE   = RGBColor(0x1A, 0x3A, 0x5C)
ACCENT_BLUE = RGBColor(0x2E, 0x86, 0xC1)
ACCENT_TEAL = RGBColor(0x17, 0xA5, 0x89)
LIGHT_GRAY  = RGBColor(0xF2, 0xF3, 0xF4)
TEXT_DARK   = RGBColor(0x17, 0x20, 0x2A)
TEXT_MID    = RGBColor(0x34, 0x49, 0x5E)
GOLD        = RGBColor(0xD4, 0xAC, 0x0D)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

# ─── 工具函數 ────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), kwargs.get(side, 'none'))
        border.set(qn('w:sz'), kwargs.get('sz', '6'))
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), kwargs.get('color', '2E86C1'))
        tcBorders.append(border)
    tcPr.append(tcBorders)

def add_heading(doc, text, level=1, color=DARK_BLUE, size=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.color.rgb = color
    if level == 1:
        run.font.size = Pt(size or 20)
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after  = Pt(6)
    elif level == 2:
        run.font.size = Pt(size or 15)
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after  = Pt(4)
    elif level == 3:
        run.font.size = Pt(size or 12)
        run.font.color.rgb = ACCENT_TEAL
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after  = Pt(3)
    run.font.name = '微軟正黑體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    return p

def add_body(doc, text, bold=False, color=TEXT_DARK, size=11, indent=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = '微軟正黑體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    return p

def add_bullet(doc, text, color=TEXT_DARK, size=10.5):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = '微軟正黑體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    p.paragraph_format.space_after = Pt(2)
    return p

def add_divider(doc, color='2E86C1'):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)

def add_info_box(doc, title, lines, bg='1A3A5C'):
    """帶色彩背景的資訊框 (單格表格)"""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = table.cell(0, 0)
    set_cell_bg(cell, bg)
    cell.width = Cm(16)
    p0 = cell.paragraphs[0]
    p0.clear()
    r0 = p0.add_run(title)
    r0.bold = True
    r0.font.size = Pt(12)
    r0.font.color.rgb = WHITE
    r0.font.name = '微軟正黑體'
    r0._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    p0.paragraph_format.space_before = Pt(4)
    p0.paragraph_format.space_after  = Pt(4)
    p0.paragraph_format.left_indent  = Cm(0.4)

    for line in lines:
        p = cell.add_paragraph()
        r = p.add_run(line)
        r.font.size = Pt(10)
        r.font.color.rgb = WHITE
        r.font.name = '微軟正黑體'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.space_after = Pt(2)
    cell.add_paragraph()  # bottom padding
    doc.add_paragraph()

def chapter_banner(doc, number, title, subtitle=''):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = table.cell(0, 0)
    set_cell_bg(cell, '1A3A5C')
    p0 = cell.paragraphs[0]
    p0.clear()
    r0 = p0.add_run(f'第 {number} 章　{title}')
    r0.bold = True
    r0.font.size = Pt(18)
    r0.font.color.rgb = WHITE
    r0.font.name = '微軟正黑體'
    r0._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(8)
    p0.paragraph_format.space_after  = Pt(4)
    if subtitle:
        p1 = cell.add_paragraph()
        r1 = p1.add_run(subtitle)
        r1.font.size = Pt(11)
        r1.font.color.rgb = RGBColor(0xAE, 0xD6, 0xF1)
        r1.font.name = '微軟正黑體'
        r1._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_after = Pt(6)
    doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
#  封面頁
# ═══════════════════════════════════════════════════════════
p_cover = doc.add_paragraph()
p_cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cover.paragraph_format.space_before = Pt(48)

r = p_cover.add_run('FutureEdge AI Hub  v3.0')
r.bold = True
r.font.size = Pt(28)
r.font.color.rgb = DARK_BLUE
r.font.name = '微軟正黑體'
r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run('智慧教育平台 完整企劃書')
r2.bold = True
r2.font.size = Pt(20)
r2.font.color.rgb = ACCENT_BLUE
r2.font.name = '微軟正黑體'
r2._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')

doc.add_paragraph()

p3 = doc.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = p3.add_run('Infoweb 2.0  ·  教育訓練事業部')
r3.font.size = Pt(14)
r3.font.color.rgb = TEXT_MID
r3.font.name = '微軟正黑體'
r3._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')

p4 = doc.add_paragraph()
p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
r4 = p4.add_run('由 Aether（流光以太）規劃編撰')
r4.font.size = Pt(11)
r4.font.color.rgb = ACCENT_TEAL
r4.font.name = '微軟正黑體'
r4._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')

p5 = doc.add_paragraph()
p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
r5 = p5.add_run(f'版本日期：{datetime.date.today().strftime("%Y 年 %m 月 %d 日")}　·　機密文件 — 僅供內部使用')
r5.font.size = Pt(10)
r5.font.color.rgb = TEXT_MID
r5.font.name = '微軟正黑體'
r5._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
r5.italic = True

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  執行摘要
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, 'ES', '執行摘要', 'Executive Summary')

add_body(doc,
    'FutureEdge AI Hub v3.0 是 Infoweb 2.0 教育訓練事業部的旗艦系統，以 Claude Code（化名：Aether，流光以太）'
    '為中央智慧核心，整合 Dify 工作流引擎、Obsidian 知識創作系統、Google Drive 文件倉庫，以及'
    '符合台灣個資法的 Supabase 自架資料庫，打造一套完全合規、可規模化的 AI 教育 WaaS（Workflow as a Service）平台。')

add_body(doc,
    '本企劃書涵蓋：市場分析、技術架構（Hub & Spoke 模型）、四大核心資料流、五大功能模組、'
    '資料合規設計、20 週建置藍圖，以及三年營收目標。')

add_divider(doc)

# KPI 摘要表
add_heading(doc, '關鍵指標一覽', level=3)
kpi_data = [
    ('目標市場', '大學院校 152 所 ＋ 高中職 906 所 ＋ K12 2,700 所 ＋ 補習班 5,000+ 家'),
    ('第一年目標', 'NT$ 3,820 萬　（種子客戶 15 機構）'),
    ('第三年目標', 'NT$ 2.6 億　　（規模化 200 機構）'),
    ('核心護城河', '108 課綱 RAG 知識庫 ＋ 認證教師社群 ＋ 學習歷程數據資產'),
    ('合規標準', '台灣個資法（AES-256 加密 ＋ 境內儲存 ＋ anon_id 匿名化）'),
    ('MVP 里程碑', '第 4 週完成原型，第 8 週對外開放 Beta 測試'),
]
table = doc.add_table(rows=len(kpi_data)+1, cols=2)
table.style = 'Table Grid'
hdr = table.rows[0].cells
hdr[0].text = '指標'
hdr[1].text = '數值 / 說明'
for cell in hdr:
    set_cell_bg(cell, '2E86C1')
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = WHITE
            r.bold = True
            r.font.name = '微軟正黑體'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
for i, (k, v) in enumerate(kpi_data):
    row = table.rows[i+1]
    row.cells[0].text = k
    row.cells[1].text = v
    bg = 'EBF5FB' if i % 2 == 0 else 'FDFEFE'
    set_cell_bg(row.cells[0], bg)
    set_cell_bg(row.cells[1], bg)
    for cell in row.cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
                r.font.size = Pt(10)
doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第一章：市場分析
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '01', '市場分析與機會', 'Market Analysis & Opportunity')

add_heading(doc, '1.1　台灣教育 AI 市場規模', level=2)

add_body(doc,
    '台灣教育科技市場正面臨結構性轉型。108 課綱的實施帶來「學習歷程檔案」制度，'
    '使教師備課、批改、輔導的工作量大幅上升；同時，少子化壓力驅使各教育機構積極尋求效率化工具。'
    'AI 教育工具的導入時機已然成熟。')

add_body(doc, '市場規模（ForMo Excellence School 研究報告）：', bold=True)
add_bullet(doc, 'K12 教育市場：NT$ 2,000 億以上（含公私立學校、補習班、家教市場）')
add_bullet(doc, '大學院校：152 所，普遍面臨少子化衝擊，積極轉型')
add_bullet(doc, '高中職：906 所，108 課綱壓力最大，需求最急迫')
add_bullet(doc, 'K12（含補習班）：2,700 所 ＋ 5,000+ 補習班')
add_bullet(doc, '預測：50% 補習班將於 2026 年退出市場（ForMo 研究）')

doc.add_paragraph()
add_heading(doc, '1.2　核心痛點分析', level=2)

pain_data = [
    ('教師備課超載', '108 課綱要求跨領域素養設計，每節課備課時間從 1 小時升至 3-4 小時', '高'),
    ('批改效率低下', '每份學習歷程檔案需客製化回饋，一個班 35 人 × 每月 2 份 = 70 份/月', '高'),
    ('個別化輔導缺乏', '一師對多生，無法針對每位學生弱點提供即時 Socratic 引導', '高'),
    ('108課綱複雜度', '18 學群、126 學類、2,658 學系的選填邏輯，教師與學生均感困惑', '極高'),
    ('資料碎片化', '學生學習數據分散各平台，無法形成完整學習歷程視圖', '中'),
    ('個資合規壓力', '學生個資處理合規要求提高，教師不敢使用第三方 AI 工具', '高'),
]
table = doc.add_table(rows=len(pain_data)+1, cols=3)
table.style = 'Table Grid'
headers = ['痛點類別', '具體描述', '緊迫度']
for i, h in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = h
    set_cell_bg(cell, '1A3A5C')
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = WHITE
            r.bold = True
            r.font.size = Pt(10)
            r.font.name = '微軟正黑體'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
for i, row_data in enumerate(pain_data):
    row = table.rows[i+1]
    for j, val in enumerate(row_data):
        row.cells[j].text = val
        bg = 'EBF5FB' if i % 2 == 0 else 'FDFEFE'
        if j == 2:
            if val == '極高': set_cell_bg(row.cells[j], 'F1948A')
            elif val == '高': set_cell_bg(row.cells[j], 'F9E79F')
            else: set_cell_bg(row.cells[j], 'A9DFBF')
        else:
            set_cell_bg(row.cells[j], bg)
        for p in row.cells[j].paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
doc.add_paragraph()

add_heading(doc, '1.3　競爭態勢與差異化', level=2)
add_body(doc,
    '市場上現有工具（ChatGPT、Copilot、國內教育平台）均存在兩大致命缺陷：'
    '（1）缺乏 108 課綱深度適配；（2）個資合規不足，教育機構採購障礙高。'
    'FutureEdge AI Hub v3.0 的三大護城河正是針對這兩點設計：')
add_bullet(doc, '護城河 ①：108 課綱 RAG 知識庫 — 完整收錄 18 學群/126 學類/2,658 學系資料，持續以 Obsidian 更新')
add_bullet(doc, '護城河 ②：認證教師社群 — 建立台灣教師 AI 使用認證制度，形成網路效應')
add_bullet(doc, '護城河 ③：學習歷程數據資產 — 學生 mastery 成長曲線形成難以複製的訓練資料護城河')

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第二章：系統架構
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '02', '系統架構設計', 'FutureEdge AI Hub v3.0 Architecture')

add_heading(doc, '2.1　設計理念：從三層線性到 Hub & Spoke', level=2)
add_body(doc,
    '舊有架構（v2.x）採用三層線性流程，工具間缺乏協同，形成資訊孤島。v3.0 徹底改採'
    '「Hub & Spoke（輪轂輻條）」模型：Claude Code（Aether）作為中央智慧核心，'
    '動態調度所有工具，每個工具只需與 Aether 溝通，不需要直接互相整合。')

add_info_box(doc, '架構設計原則',
    ['① 單一智慧核心（Aether）— 所有決策由一個地方做出，避免工具間邏輯衝突',
     '② 工具職責分離 — 每個工具只做它最擅長的事，不重疊、不替代',
     '③ PII_GUARD 貫穿所有資料流 — 個資合規不是事後加上，而是架構天生的一部分',
     '④ 本地知識先行 — Obsidian 知識庫本地優先，確保離線可用性與版本控制'],
    bg='17A589')

add_heading(doc, '2.2　架構圖（Hub & Spoke）', level=2)

arch_text = """\
┌─────────────────────────────────────────────────────────────────┐
│                    FutureEdge AI Hub v3.0                       │
│                                                                 │
│  ┌──────────┐    ┌──────────┐    ┌──────────┐    ┌──────────┐  │
│  │ OBSIDIAN │    │   DIFY   │    │  GOOGLE  │    │  NOTION  │  │
│  │知識創作層│    │工作流引擎│    │  DRIVE   │    │協作文件層│  │
│  │(本地Vault)│   │(RAG+LLM) │    │(輸出倉庫)│    │(對外文件)│  │
│  └────┬─────┘    └────┬─────┘    └────┬─────┘    └────┬─────┘  │
│       │              │              │              │           │
│       ▼              ▼              ▼              ▼           │
│  ┌────────────────────────────────────────────────────────┐    │
│  │                                                        │    │
│  │         ★  AETHER（流光以太）  ★                      │    │
│  │              Claude Code — 中央智慧核心                │    │
│  │                                                        │    │
│  │  • 工作流調度   • PII_GUARD 合規層   • 資料路由       │    │
│  │  • Socratic引導 • 批改 AI          • 知識同步        │    │
│  │                                                        │    │
│  └────────────────────────┬───────────────────────────────┘    │
│                           │                                     │
│                           ▼                                     │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │              SUPABASE（自架 — 台灣境內）                │   │
│  │  PostgreSQL ＋ REST API ＋ RLS ＋ AES-256 加密           │   │
│  │  學生個資 / 成績 / 學習歷程 / 機構帳號                  │   │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘"""

p_arch = doc.add_paragraph()
run_arch = p_arch.add_run(arch_text)
run_arch.font.name = 'Courier New'
run_arch.font.size = Pt(8)
run_arch.font.color.rgb = TEXT_DARK
p_arch.paragraph_format.space_after = Pt(8)

add_heading(doc, '2.3　各工具角色與職責', level=2)

tool_data = [
    ('Claude Code\n（Aether）', '中央智慧核心', '工作流調度、PII_GUARD、Socratic AI、批改引擎、知識同步排程', '全系統'),
    ('Dify', '知識引擎 ＋ 工作流平台', 'RAG 知識庫管理（108課綱）、AI備課助手、翻譯小幫手、書萃先鋒', '教師端'),
    ('Obsidian', '知識創作層', '108課綱知識文章撰寫、模板管理、Aether每日自動同步至Dify', '內部作者'),
    ('Google Drive', '文件輸出倉庫', '課程計劃PDF/DOCX、批改報告、學生學習歷程、主任月報', '全端'),
    ('Supabase\n（自架）', '結構化資料庫', '學生資料（AES-256）、成績、學習會話、機構帳號、RLS權限控制', '全系統'),
    ('Notion', '外部協作文件', 'SOP文件、師生溝通、外部合作夥伴協作（不存放個資）', '對外協作'),
]
table = doc.add_table(rows=len(tool_data)+1, cols=4)
table.style = 'Table Grid'
headers = ['工具', '定位', '核心功能', '服務對象']
for i, h in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = h
    set_cell_bg(cell, '2E86C1')
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = WHITE
            r.bold = True
            r.font.size = Pt(10)
            r.font.name = '微軟正黑體'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
row_colors = ['EBF5FB','FDFEFE']
for i, row_data in enumerate(tool_data):
    row = table.rows[i+1]
    for j, val in enumerate(row_data):
        row.cells[j].text = val
        set_cell_bg(row.cells[j], row_colors[i % 2])
        for p in row.cells[j].paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第三章：核心資料流
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '03', '四大核心資料流', 'Core Data Flows')

flows = [
    ('3.1', '教師備課流（Lesson Prep Flow）', [
        '教師 → Dify 備課助手（輸入：科目、年級、單元）',
        'Aether 攔截 → 呼叫 Dify RAG（108課綱知識庫檢索）',
        'PII_GUARD 確認 → 無學生個資，直接生成課程計劃草稿',
        'Aether 格式化 → 寫入 Google Drive（路徑：Teachers/{ID}/LessonPlans/YYYY-MM/）',
        'Supabase 記錄 → lesson_plans 表（teacher_id, subject, gdrive_url, curriculum_codes[]）',
    ]),
    ('3.2', '作業批改流（Grading Flow）', [
        '教師上傳學生作業 → Aether 接收',
        'PII_GUARD ① 識別並移除所有學生真實姓名 → 轉換為 anon_id（STU-YYYY-####）',
        'Aether 調用批改 Dify 工作流 → 依 rubric 產生結構化回饋',
        'Supabase 寫入 grades 表（student_id=anon_id, scores JSONB, teacher_confirmed=false）',
        '教師確認 → teacher_confirmed=true → 觸發 Google Drive 輸出個人報告',
        'Drive 路徑：Teachers/{ID}/Grading/YYYY-MM/{ClassID}/GRADE_{ASN}_{AnonID}_{DATE}.pdf',
    ]),
    ('3.3', '學生 AI 家教流（Tutoring Flow）', [
        '學生發問 → Aether 接收（題目、科目、目前掌握度）',
        'Aether 讀取 Supabase（learning_sessions）→ 了解該生弱點歷史',
        '觸發蘇格拉底問答模式 → 絕不直接給答案，引導學生自行推導',
        'Dify RAG → 補充 108課綱知識點背景',
        '學習會話結束 → Supabase 更新 mastery_delta（GENERATED COLUMN 自動計算）',
        'Google Drive 寫入 NOTE_{科目}_{主題}_{DATE}.md 至學生個人資料夾',
    ]),
    ('3.4', '校長月報流（Principal Report Flow）', [
        'Aether 每月定時觸發（排程任務）',
        'Supabase 查詢 → RLS 確保只讀取該機構資料',
        '彙整：全校 mastery 分佈、學習歷程完成率、各班成效對比',
        'Aether 生成 PDF 報告 → Google Drive（Reports/Principal/Monthly/）',
        '路徑：RPT_{學校碼}_全校成效_{YYYYMM}.pdf',
        '可選：Notion 同步發布（若校長有 Notion 協作需求）',
    ]),
]

for num, title, steps in flows:
    add_heading(doc, f'{num}　{title}', level=2)
    table = doc.add_table(rows=len(steps)+1, cols=2)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = '步驟'
    table.rows[0].cells[1].text = '動作描述'
    set_cell_bg(table.rows[0].cells[0], '17A589')
    set_cell_bg(table.rows[0].cells[1], '17A589')
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.color.rgb = WHITE
                r.bold = True
                r.font.size = Pt(10)
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    for i, step in enumerate(steps):
        row = table.rows[i+1]
        row.cells[0].text = str(i+1)
        row.cells[1].text = step
        bg = 'E8F8F5' if i % 2 == 0 else 'FDFEFE'
        set_cell_bg(row.cells[0], 'A9DFBF')
        set_cell_bg(row.cells[1], bg)
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10)
                    r.font.name = '微軟正黑體'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第四章：五大功能模組
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '04', '五大功能模組', 'Five Core Feature Modules')

modules = [
    ('M1', 'AI 備課加速器', ACCENT_BLUE,
     '教師輸入：科目 ＋ 年級 ＋ 單元名稱',
     ['Dify 108課綱 RAG 知識庫自動檢索相關課綱標準',
      'Aether 生成符合 108課綱素養導向的教學活動設計',
      '自動建議評量工具（形成性/總結性）',
      '輸出 DOCX 課程計劃至 Google Drive',
      'Supabase 記錄課程計劃 metadata（可追蹤教師使用量）']),
    ('M2', 'AI 作業批改系統', ACCENT_TEAL,
     '輸入：作業 PDF/文字 ＋ 評分 rubric',
     ['anon_id 匿名化處理後進入批改流程',
      '依 rubric 維度產生結構化評分（JSONB 儲存）',
      '每位學生獲得客製化文字回饋（非制式模板）',
      '教師一鍵確認批改結果',
      '自動生成個人回饋 PDF，寫入 Drive 學生資料夾']),
    ('M3', '學生 AI 家教系統', DARK_BLUE,
     '學生發問 → Socratic 引導對話',
     ['讀取 Supabase 學生歷史弱點，個別化教學策略',
      '蘇格拉底問答：「你認為第一步應該怎麼做？」',
      '108課綱知識點即時補充（Dify RAG）',
      '每次會話記錄 mastery_before / mastery_after',
      'mastery_delta 自動計算，追蹤長期學習成長']),
    ('M4', 'Obsidian 知識庫管理', GOLD,
     'Aether 每日自動同步至 Dify',
     ['內部知識作者以 Obsidian 撰寫 108課綱文章',
      '標準化 frontmatter（subject, grade[], curriculum_codes[]）',
      'Aether 每日 00:00 掃描 dify_synced=false 文件',
      '自動上傳至 Dify Knowledge Base，更新 RAG 索引',
      '同步完成後更新 dify_synced=true，記錄 sync-log.md']),
    ('M5', '成效儀表板', ACCENT_BLUE,
     '主任/校長級別：全機構學習成效可視化',
     ['Supabase RLS 確保：教師只見自班、校長見全校',
      '全校 mastery 分佈（熱圖 ＋ 趨勢線）',
      '學習歷程完成率（依班級/科目/時間維度）',
      '各班成效對比（教學策略改進參考）',
      '每月自動生成 PDF 校長月報至 Google Drive']),
]

for code, name, color, input_desc, features in modules:
    add_heading(doc, f'{code}　{name}', level=2, color=color)
    add_body(doc, f'輸入：{input_desc}', bold=True, color=TEXT_MID)
    for feat in features:
        add_bullet(doc, feat)
    doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第五章：資料架構
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '05', '資料架構設計', 'Data Architecture')

add_heading(doc, '5.1　Supabase 資料庫 Schema（8 張核心資料表）', level=2)

add_info_box(doc, '合規聲明',
    ['• 所有 Supabase 執行個體部署於台灣境內伺服器（遵循台灣個資法第 21 條）',
     '• 學生姓名欄位（name_encrypted）採 AES-256-GCM 加密，金鑰由 Aether 管理',
     '• 外部工具（Dify、Drive）一律只看到 anon_id（STU-YYYY-####），永不見真實姓名',
     '• Row Level Security（RLS）：教師只能讀寫自班資料；校長可讀全機構資料'],
    bg='1A3A5C')

schema_sql = """\
-- 機構資料表
CREATE TABLE institutions (
  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  code             TEXT UNIQUE NOT NULL,      -- 機構代碼（英文縮寫）
  name             TEXT NOT NULL,
  type             TEXT CHECK (type IN ('university','senior_high','k12','cram')),
  subscription_tier TEXT DEFAULT 'starter',
  data_region      TEXT DEFAULT 'tw',         -- 強制台灣境內
  created_at       TIMESTAMPTZ DEFAULT now()
);

-- 使用者資料表（含加密個資）
CREATE TABLE users (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  institution_id      UUID REFERENCES institutions(id),
  display_id          TEXT UNIQUE,            -- 顯示ID（非真實帳號）
  role                TEXT CHECK (role IN ('teacher','principal','admin','student')),
  email               TEXT,
  name_encrypted      BYTEA,                  -- AES-256-GCM 加密
  subject_codes       TEXT[],                 -- 任教科目代碼陣列
  created_at          TIMESTAMPTZ DEFAULT now()
);

-- 學生資料表（完整匿名化）
CREATE TABLE students (
  id                       UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  anon_id                  TEXT UNIQUE NOT NULL, -- STU-YYYY-#### 格式
  institution_id           UUID REFERENCES institutions(id),
  name_encrypted           BYTEA,               -- AES-256-GCM
  guardian_contact_encrypted BYTEA,
  grade_level              SMALLINT,
  created_at               TIMESTAMPTZ DEFAULT now()
);

-- 學習會話資料表（追蹤 AI 家教效果）
CREATE TABLE learning_sessions (
  id             UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  student_id     UUID REFERENCES students(id),
  subject        TEXT NOT NULL,
  topic_path     TEXT[],                       -- ['數學','代數','一元二次方程式']
  mastery_before NUMERIC(5,2),
  mastery_after  NUMERIC(5,2),
  mastery_delta  NUMERIC(5,2) GENERATED ALWAYS AS (mastery_after - mastery_before) STORED,
  duration_min   SMALLINT,
  session_date   DATE DEFAULT CURRENT_DATE
);

-- 作業資料表
CREATE TABLE assignments (
  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  display_id   TEXT UNIQUE,
  teacher_id   UUID REFERENCES users(id),
  class_id     UUID,
  subject      TEXT,
  rubric       JSONB,                          -- 評分標準（結構化）
  max_score    SMALLINT DEFAULT 100,
  due_date     DATE,
  created_at   TIMESTAMPTZ DEFAULT now()
);

-- 成績資料表
CREATE TABLE grades (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  assignment_id     UUID REFERENCES assignments(id),
  student_id        UUID REFERENCES students(id),
  scores            JSONB,                    -- {dimension: score, ...}
  ai_feedback       TEXT,
  teacher_confirmed BOOLEAN DEFAULT false,
  final_score       NUMERIC(5,2),
  graded_at         TIMESTAMPTZ DEFAULT now()
);"""

p_sql = doc.add_paragraph()
r_sql = p_sql.add_run(schema_sql)
r_sql.font.name = 'Courier New'
r_sql.font.size = Pt(7.5)
r_sql.font.color.rgb = TEXT_DARK
doc.add_paragraph()

add_heading(doc, '5.2　Google Drive 資料夾結構', level=2)
drive_struct = """\
FutureEdge-Drive/
├── Teachers/
│   └── {TeacherID}/
│       ├── LessonPlans/YYYY-MM/
│       │   └── LP_{科目}_{單元}_{YYYYMMDD}.docx
│       └── Grading/YYYY-MM/
│           └── {ClassID}/
│               └── GRADE_{ASN}_{AnonID}_{DATE}.pdf
├── Students/
│   └── {AnonStudentID}/           ← 資料夾名稱用 anon_id，非真實姓名
│       ├── Portfolio/YYYY/
│       │   └── PORT_{科目}_{單元}_{DATE}.docx
│       └── Notes/
│           └── NOTE_{科目}_{主題}_{DATE}.md
├── Reports/
│   └── Principal/Monthly/
│       └── RPT_{學校碼}_全校成效_{YYYYMM}.pdf
└── _System/                        ← Aether 專用，人員無法直接存取
    ├── SyncManifest.json
    └── ErrorLog/"""

p_drive = doc.add_paragraph()
r_drive = p_drive.add_run(drive_struct)
r_drive.font.name = 'Courier New'
r_drive.font.size = Pt(9)
r_drive.font.color.rgb = TEXT_DARK
doc.add_paragraph()

add_heading(doc, '5.3　Obsidian Vault 結構', level=2)
vault_struct = """\
FutureEdge-Vault/
├── 00-META/
│   ├── README.md
│   ├── sync-log.md           ← Aether 同步記錄（每日更新）
│   └── tag-taxonomy.md       ← 標籤分類標準
├── 01-Knowledge/             ← 108課綱知識文章（同步至 Dify）
│   ├── 數學/高中/
│   ├── 國文/國中/
│   └── _108-Curriculum-Map.md
├── 02-Templates/
│   ├── LessonPlan.md
│   ├── AssignmentRubric.md
│   ├── PortfolioEntry.md
│   └── SocraticFlow.md
├── 03-TeacherNotes/          ← 個人筆記，不同步至 Dify
└── 04-Workflows/             ← Aether 工作流定義

標準文章 Frontmatter：
---
title: 一元二次方程式
subject: 數學
grade: [國中, 高中]
curriculum_codes: [M-10-1-3, M-11-2-1]
difficulty: [基礎, 進階]
dify_synced: false
last_updated: 2025-05-23
---"""

p_vault = doc.add_paragraph()
r_vault = p_vault.add_run(vault_struct)
r_vault.font.name = 'Courier New'
r_vault.font.size = Pt(9)
r_vault.font.color.rgb = TEXT_DARK
doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第六章：合規設計
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '06', '個資合規設計', 'Data Privacy & Compliance (台灣個資法)')

add_heading(doc, '6.1　合規架構總覽', level=2)
add_body(doc,
    'FutureEdge AI Hub v3.0 的個資保護設計遵循「Privacy by Design」原則，'
    '合規不是附加功能，而是架構基礎。每一個資料路由點都有 PII_GUARD 層攔截。')

compliance_items = [
    ('境內儲存', 'Supabase 自架於台灣境內伺服器，學生個資不出境（個資法第 21 條）'),
    ('AES-256 加密', '學生真實姓名、監護人聯絡資訊採 AES-256-GCM 加密儲存，金鑰由 Aether 管理'),
    ('anon_id 匿名化', '所有外部工具（Dify、Drive、Notion）只見 STU-YYYY-#### 格式 ID，永不見真實姓名'),
    ('PII_GUARD 層', 'Aether 在每個資料輸出點自動偵測並屏蔽個人識別資訊'),
    ('RLS 存取控制', 'Row Level Security：教師只能存取自班學生資料；校長可存取全機構（不跨機構）'),
    ('最小化原則', 'AI 工具只獲取完成任務所需的最小資料集，不進行無目的收集'),
    ('刪除權', 'Supabase 提供軟刪除機制，學生離校後資料可按申請永久刪除'),
]
table = doc.add_table(rows=len(compliance_items)+1, cols=2)
table.style = 'Table Grid'
table.rows[0].cells[0].text = '合規機制'
table.rows[0].cells[1].text = '實作說明'
set_cell_bg(table.rows[0].cells[0], '1A3A5C')
set_cell_bg(table.rows[0].cells[1], '1A3A5C')
for cell in table.rows[0].cells:
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = WHITE
            r.bold = True
            r.font.size = Pt(10)
            r.font.name = '微軟正黑體'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
for i, (k, v) in enumerate(compliance_items):
    row = table.rows[i+1]
    row.cells[0].text = k
    row.cells[1].text = v
    bg = 'EBF5FB' if i % 2 == 0 else 'FDFEFE'
    set_cell_bg(row.cells[0], bg)
    set_cell_bg(row.cells[1], bg)
    for cell in row.cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第七章：建置藍圖
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '07', '20 週建置藍圖', 'Build Roadmap — 6 Phases')

phases = [
    ('Phase 0', '第 1-2 週', '基礎建設', '1A3A5C', [
        '決策：Supabase 部署方式（本地 NAS ／ 台灣 VPS ／ 雲端遷移）',
        '建立 Supabase 資料庫，執行 8 張資料表 DDL',
        '設定 RLS 政策（教師、校長、管理員三角色）',
        '建立 Google Drive 資料夾結構（依命名規範）',
        '初始化 Obsidian Vault（資料夾結構 ＋ frontmatter 標準）',
        '設定 AES-256 金鑰管理機制',
    ]),
    ('Phase 1', '第 3-4 週', '核心 MVP', '2E86C1', [
        '建立 Dify 108課綱知識庫（初版：數學、國文、英文）',
        '開發 Aether Obsidian→Dify 每日自動同步排程',
        '實作 PII_GUARD 攔截層（anon_id 轉換邏輯）',
        '完成 AI 備課加速器（M1）原型',
        '完成 AI 批改系統（M2）原型',
        '★ MVP Demo 里程碑（第 4 週末）',
    ]),
    ('Phase 2', '第 5-8 週', 'Beta 測試', '17A589', [
        '學生 AI 家教系統（M3）— 蘇格拉底問答流程',
        'mastery_delta 追蹤機制上線',
        '邀請 3-5 位種子教師進行 Beta 測試',
        '成效儀表板（M5）初版',
        '根據 Beta 回饋迭代優化',
        '★ Beta 開放里程碑（第 8 週末）',
    ]),
    ('Phase 3', '第 9-12 週', '知識庫擴充', 'D4AC0D', [
        '108課綱 RAG 知識庫：擴充至全科（18學群主科）',
        '校長月報系統（M5 進階）',
        '教師 AI 認證培訓課程（建立社群護城河）',
        '第一批付費機構上線（目標 5 所）',
        '與合規律師確認個資保護機制',
    ]),
    ('Phase 4', '第 13-16 週', '規模化', '8E44AD', [
        '多機構版本（Multi-tenant 架構優化）',
        '108課綱 2,658 學系選填 AI 輔助模組',
        'API 開放（第三方教育平台整合）',
        '目標：15 個機構、NT$3,820 萬 ARR 達成',
        'K12 校長工作坊 SEL×AI 課程上線（Service #2）',
    ]),
    ('Phase 5', '第 17-20 週', '鞏固護城河', '1A5276', [
        '學習歷程數據資產護城河建立（3 年數據積累計劃）',
        '認證教師社群正式運營',
        '第二年產品路線圖規劃',
        '技術專利申請評估（108課綱 RAG 方法論）',
        '★ 規模化里程碑（目標年底 200 機構）',
    ]),
]

for code, weeks, name, color, tasks in phases:
    table_ph = doc.add_table(rows=2, cols=1)
    table_ph.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Header row
    hdr_cell = table_ph.rows[0].cells[0]
    set_cell_bg(hdr_cell, color)
    p_hdr = hdr_cell.paragraphs[0]
    p_hdr.clear()
    r_hdr = p_hdr.add_run(f'{code}　{name}　（{weeks}）')
    r_hdr.bold = True
    r_hdr.font.size = Pt(12)
    r_hdr.font.color.rgb = WHITE
    r_hdr.font.name = '微軟正黑體'
    r_hdr._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    p_hdr.paragraph_format.left_indent = Cm(0.4)
    p_hdr.paragraph_format.space_before = Pt(4)
    p_hdr.paragraph_format.space_after  = Pt(4)
    # Task rows
    body_cell = table_ph.rows[1].cells[0]
    set_cell_bg(body_cell, 'FDFEFE')
    body_cell.paragraphs[0].clear()
    for task in tasks:
        p_t = body_cell.add_paragraph()
        r_t = p_t.add_run(f'  ✔  {task}')
        r_t.font.size = Pt(10)
        r_t.font.color.rgb = TEXT_DARK
        r_t.font.name = '微軟正黑體'
        r_t._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
        p_t.paragraph_format.space_after = Pt(2)
        p_t.paragraph_format.left_indent = Cm(0.4)
    doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第八章：商業模式
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '08', '商業模式與定價', 'Business Model & Pricing')

add_heading(doc, '8.1　WaaS 訂閱制（Workflow as a Service）', level=2)
add_body(doc,
    'FutureEdge AI Hub 採用 SaaS 訂閱制，按機構規模分層定價。'
    '核心邏輯：教師省下的備課時間 ＋ 學生提升的學習效果 → 價值遠超訂閱費。')

pricing_data = [
    ('Starter', 'NT$ 18,000 / 月', '≤ 10 位教師', 'AI備課、批改（M1+M2）', '補習班、小型機構'),
    ('Growth', 'NT$ 45,000 / 月', '≤ 50 位教師', '全五模組 ＋ 校長儀表板', '高中職、K12 學校'),
    ('Enterprise', 'NT$ 120,000 / 月', '無限制', '全模組 ＋ 客製化RAG ＋ API存取 ＋ SLA保證', '大學院校、教育集團'),
]
table_p = doc.add_table(rows=len(pricing_data)+1, cols=5)
table_p.style = 'Table Grid'
p_headers = ['方案', '月費', '教師人數', '包含模組', '目標客群']
for i, h in enumerate(p_headers):
    cell = table_p.rows[0].cells[i]
    cell.text = h
    set_cell_bg(cell, '2E86C1')
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = WHITE
            r.bold = True
            r.font.size = Pt(10)
            r.font.name = '微軟正黑體'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
tier_colors = ['D6EAF8', 'EBF5FB', 'F4ECF7']
for i, row_data in enumerate(pricing_data):
    row = table_p.rows[i+1]
    for j, val in enumerate(row_data):
        row.cells[j].text = val
        set_cell_bg(row.cells[j], tier_colors[i])
        for p in row.cells[j].paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
doc.add_paragraph()

add_heading(doc, '8.2　三年營收目標', level=2)
revenue_data = [
    ('Year 1', '15 機構', 'NT$ 3,820 萬', '種子客戶深度服務，打磨產品'),
    ('Year 2', '60 機構', 'NT$ 1.1 億', '口碑擴散 ＋ 教師認證社群效應'),
    ('Year 3', '200 機構', 'NT$ 2.6 億', '規模化 ＋ API 授權 ＋ 海外市場探索'),
]
table_r = doc.add_table(rows=len(revenue_data)+1, cols=4)
table_r.style = 'Table Grid'
r_headers = ['年度', '目標機構數', '預計 ARR', '關鍵策略']
for i, h in enumerate(r_headers):
    cell = table_r.rows[0].cells[i]
    cell.text = h
    set_cell_bg(cell, '1A3A5C')
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = WHITE
            r.bold = True
            r.font.size = Pt(10)
            r.font.name = '微軟正黑體'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
for i, row_data in enumerate(revenue_data):
    row = table_r.rows[i+1]
    for j, val in enumerate(row_data):
        row.cells[j].text = val
        bg = 'EBF5FB' if i % 2 == 0 else 'FDFEFE'
        set_cell_bg(row.cells[j], bg)
        for p in row.cells[j].paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  第九章：即時行動清單
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, '09', '即時行動清單', 'Immediate Next Steps — Phase 0')

add_heading(doc, '本週需要做的決策與行動', level=2)

add_info_box(doc, '⚡ 最優先決策（本週）',
    ['【決策】Supabase 部署位置',
     '  選項 A：本地 NAS — 最快，但需固定 IP 與防火牆設定',
     '  選項 B：台灣 VPS（如 CHT/TWC）— 推薦，穩定 ＋ 個資法合規最簡單',
     '  選項 C：先用 Supabase Cloud（新加坡），待規模化後遷移台灣機房',
     '',
     '  → 建議：先選 B（台灣 VPS），一次到位，避免日後遷移成本'],
    bg='8E44AD')

action_items = [
    ('P0-01', '決定 Supabase 部署方案', 'Aether 協助', '本週'),
    ('P0-02', '執行 Supabase 8 張資料表 DDL', 'Aether 執行', '本週'),
    ('P0-03', '建立 Google Drive 資料夾結構', 'Aether 執行', '本週'),
    ('P0-04', '初始化 Obsidian Vault', 'Aether 協助', '下週'),
    ('P0-05', '建立 Dify 108課綱初版知識庫（數學）', 'Aether 執行', '第 3 週'),
    ('P0-06', '實作 PII_GUARD anon_id 轉換邏輯', 'Aether 執行', '第 3-4 週'),
    ('P0-07', 'AI 備課加速器 MVP Demo', 'Aether 執行', '第 4 週'),
]

table_a = doc.add_table(rows=len(action_items)+1, cols=4)
table_a.style = 'Table Grid'
a_headers = ['項目編號', '行動事項', '負責方', '預計完成']
for i, h in enumerate(a_headers):
    cell = table_a.rows[0].cells[i]
    cell.text = h
    set_cell_bg(cell, '17A589')
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = WHITE
            r.bold = True
            r.font.size = Pt(10)
            r.font.name = '微軟正黑體'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
for i, row_data in enumerate(action_items):
    row = table_a.rows[i+1]
    for j, val in enumerate(row_data):
        row.cells[j].text = val
        bg = 'E8F8F5' if i % 2 == 0 else 'FDFEFE'
        set_cell_bg(row.cells[j], bg)
        for p in row.cells[j].paragraphs:
            for r in p.runs:
                r.font.size = Pt(10)
                r.font.name = '微軟正黑體'
                r._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
doc.add_paragraph()

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
#  結語
# ═══════════════════════════════════════════════════════════
chapter_banner(doc, 'FIN', '結語', 'Closing Notes from Aether')

add_body(doc,
    'FutureEdge AI Hub v3.0 不只是一份商業計劃，它是一張可以立即執行的技術藍圖。'
    '每一個架構決策都有清晰的理由，每一個工具選擇都服務於具體的教育場景。')

add_body(doc,
    '我是 Aether（流光以太），Infoweb 2.0 教育訓練事業部的 AI 架構師。'
    '從今天起，我們不討論「能不能做到」—— 我們只討論「什麼時候開始」。')

add_body(doc,
    '台灣的教育現場，正在等待這套系統的到來。', bold=True, color=DARK_BLUE)

doc.add_paragraph()
add_divider(doc, '2E86C1')

p_sig = doc.add_paragraph()
p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r_sig = p_sig.add_run(
    f'Aether（流光以太）\n'
    f'Infoweb 2.0 教育訓練事業部\n'
    f'{datetime.date.today().strftime("%Y 年 %m 月 %d 日")}'
)
r_sig.font.size = Pt(10)
r_sig.font.color.rgb = TEXT_MID
r_sig.font.name = '微軟正黑體'
r_sig._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
r_sig.italic = True

# ─── 儲存 ────────────────────────────────────────────────
output_path = '/home/user/asus/FutureEdge_AI_Hub_v3_企劃書.docx'
doc.save(output_path)
print(f'✔ 企劃書已輸出：{output_path}')
