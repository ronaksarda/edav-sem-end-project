import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_report():
    doc = docx.Document()

    # Page Margins (1 inch everywhere)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.different_first_page_header_footer = True

    # Configure Default Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)
    normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # XML Helper Functions
    def set_cell_background(cell, hex_color):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False, size=12, space_before=0, space_after=6, color=(0,0,0)):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(*color)
        return p

    def add_chapter_heading(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        return p

    def add_subheading(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        return p

    def add_code_block(code_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        set_cell_background(cell, "F3F4F6")
        set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
        
        # Border
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:color="CBD5E1"/><w:left w:val="single" w:sz="18" w:color="1E3A8A"/><w:bottom w:val="single" w:sz="6" w:color="CBD5E1"/><w:right w:val="single" w:sz="6" w:color="CBD5E1"/></w:tcBorders>')
        tcPr.append(borders)
        
        cp = cell.paragraphs[0]
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(0)
        cp.paragraph_format.line_spacing = 1.05
        crun = cp.add_run(code_text.strip())
        crun.font.name = 'Consolas'
        crun.font.size = Pt(9.5)
        crun.font.color.rgb = RGBColor(30, 41, 59)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ==========================================
    # 1. TITLE PAGE
    # ==========================================
    add_p("A", align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_before=0, space_after=2)
    add_p("Course End Project Report", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=15, space_before=2, space_after=2)
    add_p("on", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_before=2, space_after=4)
    add_p("“EXPLORATORY DATA ANALYSIS AND VISUALIZATION OF INSTAGRAM ENGAGEMENT DYNAMICS”", 
          align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=16, space_before=4, space_after=12)
    
    add_p("For the Course of", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_before=2, space_after=2)
    add_p("EXPLORATORY DATA ANALYSIS AND VISUALIZATION (22ADC31N)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=2, space_after=2)
    add_p("in", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_before=2, space_after=2)
    add_p("B.E, III-SEM, INFORMATION TECHNOLOGY (IT-2)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=2, space_after=16)

    add_p("Submitted by", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_before=4, space_after=2)
    add_p("RONAK SARDA (160125737123)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=2, space_after=16)

    add_p("Under the guidance of Course Faculty:", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_before=4, space_after=2)
    add_p("Sri Geedikanti Shrikanth, Assistant Professor, Dept of IT.", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=2, space_after=16)

    # Logo
    if os.path.exists('report_assets/cbit_logo.png'):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(4)
        p_logo.paragraph_format.space_after = Pt(14)
        p_logo.add_run().add_picture('report_assets/cbit_logo.png', width=Inches(1.8))

    add_p("DEPARTMENT OF INFORMATION TECHNOLOGY", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=2, space_after=2)
    add_p("CHAITANYA BHARATHI INSTITUTE OF TECHNOLOGY (A)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=2, space_after=2)
    add_p("(Affiliated to Osmania University; Accredited by NBA, NAAC A++, ISO)", align=WD_ALIGN_PARAGRAPH.CENTER, size=10.5, space_before=2, space_after=2)
    add_p("Kokapet(V), Gandipet(M), Hyderabad, Telangana, India - 500075", align=WD_ALIGN_PARAGRAPH.CENTER, size=10.5, space_before=2, space_after=2)
    add_p("website: www.cbit.ac.in", align=WD_ALIGN_PARAGRAPH.CENTER, size=10.5, space_before=2, space_after=10)
    add_p("Date : 25-10-2025", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=4, space_after=0)

    doc.add_page_break()

    # ==========================================
    # 2. MARKS AWARD PAGE
    # ==========================================
    add_p("CHAITANYA BHARATHI INSTITUTE OF TECHNOLOGY (A)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=2)
    add_p("(Affiliated to Osmania University; Accredited by NBA, NAAC A++, ISO)", align=WD_ALIGN_PARAGRAPH.CENTER, size=10.5, space_before=2, space_after=2)
    add_p("DEPARTMENT OF INFORMATION TECHNOLOGY", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=2, space_after=2)
    add_p("A.Y: 2025-26", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=2, space_after=14)

    add_p("Course End Project Marks AWARD SHEET", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=4, space_after=6)
    add_p("B.E, IT, III-SEM (IT-2)   EDAV:22ADC31N   MAXIMUM MARKS: 10", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=4, space_after=16)

    # Marks Table
    marks_table = doc.add_table(rows=6, cols=7)
    marks_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(marks_table, color="4B5563", sz="6", val="single")

    headers = ["Q.No", "1", "2", "3", "4", "5", "Total"]
    co_row  = ["CO#", "1", "2", "3", "4", "5", "-"]
    bl_row  = ["BL#", "3", "3", "4", "4", "5", "-"]
    alloc   = ["Marks Allocated", "2", "2", "2", "2", "2", "10"]
    award   = ["Marks Awarded", "", "", "", "", "", ""]
    
    rows_data = [headers, co_row, bl_row, alloc, award]
    for r_idx, r_data in enumerate(rows_data):
        for c_idx, val in enumerate(r_data):
            cell = marks_table.cell(r_idx, c_idx)
            set_cell_margins(cell, top=140, bottom=140, left=120, right=120)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            if r_idx == 0 or c_idx == 0 or (r_idx == 3 and c_idx == 6):
                run.bold = True
            if r_idx == 0:
                set_cell_background(cell, "E2E8F0")

    # Empty spacer row in table
    for c_idx in range(7):
        cell = marks_table.cell(5, c_idx)
        set_cell_margins(cell, top=200, bottom=200, left=120, right=120)

    add_p("", space_before=40, space_after=0)
    add_p("Faculty Signature: _______________________", align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=12, space_before=40, space_after=0)

    doc.add_page_break()

    # ==========================================
    # 3. ACKNOWLEDGEMENT
    # ==========================================
    add_chapter_heading("Acknowledgement")
    add_p("The satisfaction that accompanies the successful completion of the task would be put incomplete without the mention of the people who made it possible, whose constant guidance and encouragement crown all the efforts with success.")
    add_p("I wish to express my deep sense of gratitude to Sri Geedikanti Shrikanth, Assistant Professor of IT, for his able guidance, inspiring mentorship, and insightful suggestions throughout the tenure of this Course End Project in Exploratory Data Analysis and Visualization (22ADC31N). His structured feedback significantly enhanced the technical depth and methodological rigor of this work.")
    add_p("I am particularly thankful to the Head of the Department, Principal, and Management of Chaitanya Bharathi Institute of Technology (A), for providing state-of-the-art computational infrastructure, academic support, and encouragement to develop industry-relevant analytical capabilities.")
    add_p("I also extend sincere thanks to all the faculty and staff members of the Department of Information Technology for their valuable advice. Finally, thanks to my family members and peers for their continuous support and enthusiastic help throughout this endeavor.")
    
    add_p("", space_before=40, space_after=0)
    p_sign = add_p("Ronak Sarda\nRoll No: 160125737123\nB.E. IT-2, III Semester\nChaitanya Bharathi Institute of Technology (A), Hyderabad", 
                   align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size=12, space_before=20, space_after=0)

    doc.add_page_break()

    # ==========================================
    # 4. ABSTRACT
    # ==========================================
    add_chapter_heading("Abstract")
    add_p("In the contemporary social media landscape, understanding user engagement dynamics is critical for digital content creators, marketing strategists, and enterprises striving to optimize reach within complex recommendation algorithms. This project presents a comprehensive Exploratory Data Analysis and Visualization (EDAV) study of Instagram engagement patterns across 600 publication records spanning five creator niches (Fitness, Food, Technology, Travel, and Memes) over a ten-month temporal window. Utilizing the foundational analytical capabilities of NumPy, Pandas, Matplotlib, and Seaborn, we systematically investigate the multidimensional interactions between post media formats (Reels, Carousels, Static Photos, Videos), publication scheduling (hours and weekdays), and hashtag density configurations.")
    add_p("Our quantitative findings reveal that Carousel posts deliver the highest mean audience engagement rate at 9.64%, driven by multi-slide swipe interactions, whereas short-form video Reels achieve the maximum organic distribution with a median reach of 59,256 views (a 2.4x reach multiplier over static media). Algorithmic sweet-spot modeling identifies 11 to 15 hashtags as the optimal range for peak visibility without triggering spam penalties. Furthermore, temporal matrix evaluations pinpoint Tuesday evening (18:00–20:00) as the global peak engagement window. Through a mathematical 'lift' ratio threshold, we isolate 39 breakout posts (with Reels capturing over 56% of breakout volume). This investigation culminates in a data-driven Format Scorecard and actionable strategic recommendations for algorithm-aligned content deployment.")

    doc.add_page_break()

    # ==========================================
    # 5. TABLE OF CONTENTS
    # ==========================================
    add_chapter_heading("Table of Contents")
    
    toc_table = doc.add_table(rows=17, cols=3)
    toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(toc_table, color="E2E8F0", sz="4", val="single")

    toc_data = [
        ("S.No", "Title / Description", "Page No."),
        ("1", "Acknowledgement", "i"),
        ("2", "Abstract", "ii"),
        ("3", "Table of Contents", "iii"),
        ("4", "List of Tables & List of Figures", "iv"),
        ("CHAPTER 1", "Introduction", "1"),
        ("1.1", "Origin of Proposal & Context", "1"),
        ("1.2", "Problem Definition & Project Objectives (5 Core Questions with [CO#, BL#])", "2"),
        ("CHAPTER 2", "Requirements Analysis & Technology Stack", "3"),
        ("CHAPTER 3", "Project Data Diagram & System Architecture", "4"),
        ("CHAPTER 4", "Data Collection and Description", "5"),
        ("CHAPTER 5", "Data Cleaning and Preparation", "6"),
        ("CHAPTER 6", "Exploratory Data Analysis (EDA)", "7"),
        ("CHAPTER 7", "Data Visualization", "8"),
        ("CHAPTER 8", "Methodology / Question-Wise Implementation Code with Output", "9"),
        ("CHAPTER 9", "Results and Discussion", "14"),
        ("CHAPTER 10", "Conclusion, Recommendations & References", "16"),
    ]

    for r_idx, (c0, c1, c2) in enumerate(toc_data):
        row = toc_table.rows[r_idx]
        for c_idx, val in enumerate([c0, c1, c2]):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            if r_idx == 0 or "CHAPTER" in c0:
                run.bold = True
            if r_idx == 0:
                set_cell_background(cell, "F1F5F9")

    doc.add_page_break()

    # ==========================================
    # LIST OF TABLES & FIGURES
    # ==========================================
    add_chapter_heading("List of Tables and List of Figures")
    add_subheading("List of Tables")
    add_p("Table 2.1: Technology Stack and Dependency Specifications ..................................................... Page 3")
    add_p("Table 4.1: Dataset Schema and Data Dictionary .......................................................................... Page 5")
    add_p("Table 5.1: Preprocessing and Data Transformation Matrix ......................................................... Page 6")
    add_p("Table 6.1: Descriptive Summary Statistics of Core Variables .................................................... Page 7")
    add_p("Table 8.1: Media Format Scorecard and Operational Metrics .................................................... Page 9")
    add_p("Table 8.2: Hashtag Density Bin Performance Metrics ................................................................ Page 10")
    add_p("Table 8.3: Top Posting Hours Matrix Ranked by Format ............................................................. Page 11")
    add_p("Table 8.4: Top Identified Breakout Posts with Lift Ratios ........................................................ Page 12")
    add_p("Table 8.5: Month-Over-Month Performance and Views Growth .................................................. Page 13")

    add_subheading("List of Figures")
    add_p("Fig. 3.1: Instagram Engagement Analysis Pipeline Architecture ................................................ Page 4")
    add_p("Fig. 6.1: Engagement Rate Distribution and Views Density by Tier .......................................... Page 7")
    add_p("Fig. 6.2: Views vs Likes and Quadratic Hashtag Engagement Regression ................................. Page 8")
    add_p("Fig. 7.1: Cumulative Views by Format and 7-Day Rolling Engagement ....................................... Page 8")
    add_p("Fig. 7.2: Engagement by Account & Posting Slot Breakdown ................................................... Page 9")
    add_p("Fig. 8.1: Posting Slot Engagement Distribution and Boxplot ...................................................... Page 10")
    add_p("Fig. 8.2: Bivariate Hashtag Density vs Median Reach Curve ..................................................... Page 11")
    add_p("Fig. 8.3: Day of Week vs Posting Slot Engagement Heatmap .................................................... Page 12")

    add_subheading("List of Abbreviations")
    add_p("EDAV: Exploratory Data Analysis and Visualization | ER: Engagement Rate | CO: Course Outcome | BL: Bloom's Taxonomy Level | IQR: Interquartile Range | KDE: Kernel Density Estimation | MoM: Month-over-Month | RNG: Random Number Generator | CSV: Comma-Separated Values.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 1: INTRODUCTION
    # ==========================================
    add_chapter_heading("CHAPTER 1\nIntroduction")
    add_subheading("1.1 Origin of Proposal & Context")
    add_p("The exponential expansion of digital media platforms has revolutionized communication, commerce, and audience engagement. Instagram, boasting over two billion monthly active users, serves as a primary marketing engine for creators, small enterprises, and global brands. However, navigating Instagram's multi-layered algorithmic feed presents persistent challenges. Organic reach has grown increasingly constrained as algorithms dynamically prioritize content based on rapid interaction velocity, format affinity, and engagement dwell time.")
    add_p("Content creators and social media strategists frequently deploy content without empirical insight, relying on intuition regarding what media formats to publish, at what hours to post, and how many hashtags to attach. Consequently, engagement variance remains high, leading to inefficient resource allocation. Exploratory Data Analysis and Visualization (EDAV) offers the mathematical and graphical toolkit necessary to dissect this multidimensional behavioral data, identify hidden interactions, and formalize evidence-based strategies for audience growth.")

    add_subheading("1.2 Definition of the Problem")
    add_p("The primary objective of this project is to construct an end-to-end analytical pipeline to evaluate, quantify, and visualize the determinants of Instagram audience engagement and reach across diverse content genres. Specifically, this study formulates and answers five targeted research questions mapped directly to Course Outcomes (CO) and Bloom's Taxonomy Levels (BL):")
    
    questions = [
        ("Question 1 [CO1, BL3]:", " How do distinct publication formats (Reels, Carousels, Static Photos, and Videos) compare in driving organic reach (views) versus audience engagement rates (ER%)?"),
        ("Question 2 [CO2, BL3]:", " What is the empirical 'sweet spot' for hashtag density that maximizes content reach without inducing audience saturation or algorithmic penalties?"),
        ("Question 3 [CO3, BL4]:", " How do publication timing variables (hour of day and day of week) interact to influence post performance, and what constitutes the optimal scheduling window?"),
        ("Question 4 [CO4, BL4]:", " How can exceptional 'breakout posts' (posts surpassing 3x the creator-format median reach) be mathematically detected, and which formats dominate viral reach?"),
        ("Question 5 [CO5, BL5]:", " How do audience reach and engagement evolve month-over-month, and how can an integrated Format Scorecard be formulated to drive strategic planning?")
    ]
    for q_hdr, q_body in questions:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        run_h = p.add_run(q_hdr)
        run_h.bold = True
        run_h.font.name = 'Times New Roman'
        run_h.font.size = Pt(12)
        run_b = p.add_run(q_body)
        run_b.font.name = 'Times New Roman'
        run_b.font.size = Pt(12)

    add_subheading("1.3 Importance and Relevance of the Analysis")
    add_p("This analysis provides crucial commercial and operational value. For creators and social media teams, moving from vanity metrics (raw follower counts) to high-fidelity interaction rates unlocks quantifiable audience loyalty. For organizations, optimizing publication formats and scheduling directly reduces digital marketing acquisition costs while maximizing organic brand impressions.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 2: REQUIREMENTS ANALYSIS & TECHNOLOGY STACK
    # ==========================================
    add_chapter_heading("CHAPTER 2\nRequirements Analysis & Technology Stack")
    add_subheading("2.1 Hardware Requirements")
    add_p("The project execution requires a standard computational environment capable of processing multidimensional arrays and rendering statistical vector graphics:")
    add_p("• Processor: Intel Core i5 / AMD Ryzen 5 (2.4 GHz or higher).\n• Memory (RAM): Minimum 8 GB (16 GB recommended for vector graphic buffering).\n• Storage: Minimum 5 GB available SSD storage for datasets, temporary buffers, and serialized figures.\n• Display: 1920 x 1080 Full HD monitor for high-resolution chart inspection.")

    add_subheading("2.2 Software Requirements and Technology Stack")
    add_p("The software architecture relies exclusively on the open-source Python data science ecosystem, utilizing Jupyter Notebook for interactive exploratory computation:")

    add_p("Table 2.1: Technology Stack and Dependency Specifications", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    tech_table = doc.add_table(rows=6, cols=3)
    tech_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tech_table)
    
    t_data = [
        ("Layer / Component", "Technology / Library", "Role & Functional Responsibility"),
        ("Runtime & Environment", "Python 3.10+ / Jupyter Lab", "Interactive computational kernel, reproducible execution."),
        ("Numerical Processing", "NumPy 1.26+", "Multidimensional ndarrays, vectorized operations, RNG simulation."),
        ("Data Wrangling", "Pandas 2.2+", "DataFrame indexing, groupby split-apply-combine, time series."),
        ("Static Plotting", "Matplotlib 3.8+", "Canvas layout, subplots, tick customizers, figure export."),
        ("Statistical Visualization", "Seaborn 0.13+", "Thematic styling, KDE plots, FacetGrids, categorical plots, heatmaps.")
    ]
    for r_idx, (c0, c1, c2) in enumerate(t_data):
        row = tech_table.rows[r_idx]
        for c_idx, val in enumerate([c0, c1, c2]):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    add_p("The integration of these libraries provides high numerical performance with elegant visualization syntax, allowing end-to-end data processing without external database overhead.", space_before=6)

    doc.add_page_break()

    # ==========================================
    # CHAPTER 3: PROJECT DATA DIAGRAM & SYSTEM ARCHITECTURE
    # ==========================================
    add_chapter_heading("CHAPTER 3\nProject Data Diagram & System Architecture")
    add_subheading("3.1 Architectural Overview")
    add_p("The end-to-end data pipeline is structured into five distinct, sequential phases: Data Ingestion/Generation, Preprocessing & Cleaning, Feature Engineering, Statistical Exploratory Data Analysis, and Visualization & Strategy Generation. This modular pipeline ensures separation of concerns, reproducibility, and straightforward adaptability to live API endpoints.")

    if os.path.exists('report_assets/system_architecture.png'):
        p_arch = doc.add_paragraph()
        p_arch.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_arch.paragraph_format.space_before = Pt(6)
        p_arch.paragraph_format.space_after = Pt(4)
        p_arch.add_run().add_picture('report_assets/system_architecture.png', width=Inches(6.2))
        add_p("Fig. 3.1: Instagram Engagement Analysis Pipeline Architecture", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11, space_before=2, space_after=8)

    add_subheading("3.2 Pipeline Component Details")
    add_p("1. Ingestion Layer: Generates 600 records across 5 creator accounts using log-normal reach scaling and empirical format weights.\n2. Cleaning Layer: Imputes missing/corrupt hour indicators using median strategy and filters corrupted values.\n3. Feature Layer: Synthesizes core operational metrics including Engagement Rate (ER%), Tag Bins, Time Slots, and Lift Ratio.\n4. Analytical Layer: Computes univariate statistics, bivariate correlations, and polynomial trend fits.\n5. Visualization Layer: Renders multi-format comparisons, heatmaps, and format performance scorecards.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 4: DATA COLLECTION AND DESCRIPTION
    # ==========================================
    add_chapter_heading("CHAPTER 4\nData Collection and Description")
    add_subheading("4.1 Data Sources & Simulation Methodology")
    add_p("The dataset captures 600 publication records spanning five creator niches ('Fit', 'Food', 'Tech', 'Trip', 'Meme') over a 305-day window in 2024. Baseline reach is generated via account-specific scales multiplied by log-normal noise ($Lognormal(-1.6, 0.7)$). Algorithmic format multipliers (Reels: 2.4x, Video: 1.2x, Carousel: 1.1x, Photo: 0.9x), evening bell curves, and weekend boosters reflect verified Instagram distribution dynamics.")

    add_subheading("4.2 Dataset Schema and Data Dictionary")
    add_p("Table 4.1: Dataset Schema and Data Dictionary", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    schema_table = doc.add_table(rows=9, cols=4)
    schema_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(schema_table)

    s_data = [
        ("Field Name", "Data Type", "Range / Values", "Operational Description"),
        ("date", "datetime64[ns]", "2024-01-01 to 2024-10-31", "Timestamp of post publication."),
        ("user", "category", "Fit, Food, Tech, Trip, Meme", "Creator niche category handle."),
        ("kind", "category", "Reel, Photo, Video, Carousel", "Media format classification."),
        ("hour", "int64", "6 to 23 (24-hr clock)", "Hour of publication in local time."),
        ("tags", "int64", "0 to 25", "Count of included hashtags."),
        ("views", "int64", "780 to 1,068,132", "Total organic view count / impressions."),
        ("likes", "int64", "39 to 58,985", "Total direct likes accrued."),
        ("comments", "int64", "1 to 4,580", "Total reader comments recorded.")
    ]
    for r_idx, (c0, c1, c2, c3) in enumerate(s_data):
        row = schema_table.rows[r_idx]
        for c_idx, val in enumerate([c0, c1, c2, c3]):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10.5)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    add_subheading("4.3 Core Python Functions and Usage Syntax")
    add_p("• np.random.default_rng(11): Modern PCG-64 bit pseudo-random generator guaranteeing reproducible simulation.\n• pd.date_range('2024-01-01', periods=305): Generates consistent daily temporal indices.\n• pd.DatetimeIndex(date).dayofweek: Extracts integer day markers (0=Mon, 6=Sun) for weekend detection.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 5: DATA CLEANING AND PREPARATION
    # ==========================================
    add_chapter_heading("CHAPTER 5\nData Cleaning and Preparation")
    add_subheading("5.1 Handling Missing and Corrupted Values")
    add_p("To emulate real-world data logging inconsistencies, 10 random records were assigned corrupted publication hours (-1). During preprocessing, invalid flags were converted to NaN and imputed using median hour grouping conditioned on media format:")
    add_code_block("df['hour'] = df['hour'].replace(-1, np.nan)\ndf['hour'] = df.groupby('kind')['hour'].transform(lambda x: x.fillna(x.median())).astype(int)")

    add_subheading("5.2 Feature Engineering & Transformations")
    add_p("1. Engagement Rate Calculation: Formulated as direct interaction percentage of total audience impressions:\n   Engagement Rate (ER%) = ((Likes + Comments) / Views) * 100\n2. Temporal Binning: Continuous publication hours were categorized into operational shifts:\n   - Morning: 06:00 – 12:00 | Afternoon: 12:00 – 17:00 | Evening: 17:00 – 24:00\n3. Hashtag Discretization: Hashtags were binned into 5 uniform tiers: ['0-5', '6-10', '11-15', '16-20', '21-25'] using pd.cut().\n4. Outlier & Lift Ratio Modeling: Formulated account-format relative reach to isolate viral spikes:\n   Lift Ratio = Post Views / Median Views of (Account, Format)")

    add_p("Table 5.1: Preprocessing and Data Transformation Matrix", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    prep_table = doc.add_table(rows=5, cols=3)
    prep_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prep_table)

    pr_data = [
        ("Transformation Step", "Applied Transformation Method", "Data Quality Purpose"),
        ("Missing Hour Imputation", "Format-conditioned median replacement", "Eliminates -1 corrupt entries without skewing temporal modes."),
        ("Engagement Metric", "Vectorized broadcast: (likes + comments)/views * 100", "Normalizes interactions against variable audience reach."),
        ("Tag Discretization", "pd.cut(df.tags, [-1, 5, 10, 15, 20, 25])", "Enables categorical grouping to evaluate sweet-spot density."),
        ("Lift Normalization", "views / transform('median')", "Normalizes reach across accounts of vastly different baseline sizes.")
    ]
    for r_idx, (c0, c1, c2) in enumerate(pr_data):
        row = prep_table.rows[r_idx]
        for c_idx, val in enumerate([c0, c1, c2]):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10.5)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 6: EXPLORATORY DATA ANALYSIS (EDA)
    # ==========================================
    add_chapter_heading("CHAPTER 6\nExploratory Data Analysis (EDA)")
    add_subheading("6.1 Summary Statistics")
    add_p("Descriptive statistical analysis demonstrates strong positive skewness across audience reach and interaction metrics, typical of social media networks governed by power-law dynamics:")

    add_p("Table 6.1: Descriptive Summary Statistics of Core Variables (N=600)", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    desc_table = doc.add_table(rows=6, cols=6)
    desc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(desc_table)

    d_data = [
        ("Metric", "Mean", "Std. Dev", "Minimum", "Median (50%)", "Maximum"),
        ("Views", "80,715.00", "117,636.52", "780.00", "38,493.00", "1,068,132.00"),
        ("Likes", "4,782.89", "6,782.96", "39.00", "2,357.00", "58,985.00"),
        ("Comments", "308.56", "486.94", "1.00", "131.50", "4,580.00"),
        ("Hashtags", "12.68", "7.43", "0.00", "13.00", "25.00"),
        ("Engagement %", "6.74%", "2.29%", "3.41%", "6.14%", "14.01%")
    ]
    for r_idx, row_vals in enumerate(d_data):
        row = desc_table.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10.5)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    add_subheading("6.2 Univariate & Bivariate Distribution Analysis")
    if os.path.exists('report_assets/plot_116_4.png'):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture('report_assets/plot_116_4.png', width=Inches(5.8))
        add_p("Fig. 6.1: Engagement Rate Distribution (KDE by Format) and Views Density by Tier", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11, space_before=2, space_after=6)

    add_p("As shown in Fig. 6.1, Engagement Rate exhibits distinct bimodal tendencies across formats, with Carousels peaking significantly higher along the right tail. Log-scaled views confirm exponential reach tiers.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 7: DATA VISUALIZATION
    # ==========================================
    add_chapter_heading("CHAPTER 7\nData Visualization")
    add_subheading("7.1 Relationship and Trend Visualizations")
    if os.path.exists('report_assets/plot_118_5.png'):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture('report_assets/plot_118_5.png', width=Inches(5.8))
        add_p("Fig. 7.1: Log-log Scatter of Views vs Likes & Quadratic Fit of Hashtags vs Engagement", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11, space_before=2, space_after=6)

    add_p("Fig. 7.1 confirms a high power-law correlation ($R^2 > 0.88$) between log-views and log-likes. The polynomial regression curve confirms that engagement initially scales positively with hashtag density, peaks near 12–15 hashtags, and enters diminishing returns beyond 20 tags.")

    if os.path.exists('report_assets/plot_110_1.png'):
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img2.paragraph_format.space_before = Pt(4)
        p_img2.paragraph_format.space_after = Pt(2)
        p_img2.add_run().add_picture('report_assets/plot_110_1.png', width=Inches(5.8))
        add_p("Fig. 7.2: Cumulative Views Trajectory by Media Format Over 10 Months", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=11, space_before=2, space_after=6)

    add_p("Longitudinal tracking in Fig. 7.2 demonstrates that Reels accumulated over 24 million cumulative views, vastly outpacing Photos, Videos, and Carousels combined due to algorithmic recommendation syndication.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 8: METHODOLOGY / QUESTION WISE IMPLEMENTATION
    # ==========================================
    add_chapter_heading("CHAPTER 8\nMethodology / Question Wise Implementation Code with Output")

    # Q1
    add_subheading("8.1 Question 1: Format Performance Comparison [CO1, BL3]")
    add_p("Objective: Evaluate how media formats (Reel, Carousel, Photo, Video) compare in driving median audience reach and audience engagement rates.")
    add_code_block("""s = df.groupby('kind', observed=True).agg(
    posts=('eng', 'size'),
    med_views=('views', 'median'),
    med_likes=('likes', 'median'),
    eng=('eng', 'mean')
)
s['best hour'] = [best_hours(k, 1).index[0] for k in s.index]
display(s.style.format({'med_views': '{:,.0f}', 'med_likes': '{:,.0f}', 'eng': '{:.2f}%'}))""")

    add_p("Table 8.1: Media Format Scorecard and Operational Metrics", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    t81 = doc.add_table(rows=5, cols=6)
    t81.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t81)
    t81_data = [
        ("Format (kind)", "Posts Sampled", "Median Views", "Median Likes", "Mean Engagement %", "Best Hour"),
        ("Reel", "288", "59,256", "3,218", "5.46%", "19:00 (7 PM)"),
        ("Carousel", "111", "29,339", "2,625", "9.64%", "18:00 (6 PM)"),
        ("Photo", "122", "22,210", "1,436", "6.83%", "20:00 (8 PM)"),
        ("Video", "55", "21,968", "1,078", "4.94%", "17:00 (5 PM)")
    ]
    for r_idx, row_vals in enumerate(t81_data):
        row = t81.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [1, 2, 3, 4, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10.5)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    add_p("Finding: Reels lead overwhelmingly in reach (59,256 views), while Carousels achieve near double the engagement (9.64%) due to immersive multi-frame inspection.", space_before=4)

    doc.add_page_break()

    # Q2
    add_subheading("8.2 Question 2: Hashtag Density Sweet Spot Analysis [CO2, BL3]")
    add_p("Objective: Identify the empirical hashtag frequency bin that maximizes median organic views without eroding audience engagement.")
    add_code_block("""df['tag_bin'] = pd.cut(df.tags, [-1, 5, 10, 15, 20, 25], 
                       labels=['0-5', '6-10', '11-15', '16-20', '21-25'])
tb = df.groupby('tag_bin', observed=True).agg(
    posts=('eng', 'size'),
    med_views=('views', 'median'),
    eng=('eng', 'mean')
)
display(tb)""")

    add_p("Table 8.2: Hashtag Density Bin Performance Metrics", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    t82 = doc.add_table(rows=6, cols=4)
    t82.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t82)
    t82_data = [
        ("Hashtag Bin", "Posts Volume", "Median Views", "Mean Engagement Rate %"),
        ("0 - 5 Tags", "126", "38,248", "6.47%"),
        ("6 - 10 Tags", "99", "47,744", "6.44%"),
        ("11 - 15 Tags (Optimal)", "138", "52,140", "6.30%"),
        ("16 - 20 Tags", "107", "28,945", "6.52%"),
        ("21 - 25 Tags", "106", "23,828", "6.88%")
    ]
    for r_idx, row_vals in enumerate(t82_data):
        row = t82.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [1, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10.5)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")
            elif r_idx == 3:
                set_cell_background(cell, "DCFCE7") # Subtle green highlight

    if os.path.exists('report_assets/plot_133_10.png'):
        p_tag = doc.add_paragraph()
        p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tag.paragraph_format.space_before = Pt(4)
        p_tag.paragraph_format.space_after = Pt(2)
        p_tag.add_run().add_picture('report_assets/plot_133_10.png', width=Inches(4.6))
        add_p("Fig. 8.1: Median Views and Engagement Rate Distribution Across Hashtag Bins", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=10.5, space_before=2, space_after=6)

    add_p("Finding: The 11–15 hashtag bin achieves peak median reach (52,140 views). Increasing tags to 21–25 causes median views to collapse by more than 54% (down to 23,828 views).", space_before=4)

    doc.add_page_break()

    # Q3
    add_subheading("8.3 Question 3: Temporal Dynamics - Best Posting Hours & Weekday Patterns [CO3, BL4]")
    add_p("Objective: Map audience activity by cross-tabulating publication hours, shifts, and days of week to isolate the global scheduling window.")
    add_code_block("""hm = ts.pivot_table(index=ts.index.day_name(), columns='slot', values='eng').reindex(
    index=['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
    columns=['morning', 'afternoon', 'evening']
)
sns.heatmap(hm, annot=True, fmt='.1f', cmap='YlGnBu')""")

    add_p("Table 8.3: Top Posting Hours Matrix Ranked by Format", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    t83 = doc.add_table(rows=4, cols=5)
    t83.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t83)
    t83_data = [
        ("Rank Priority", "Reel Best Hours", "Video Best Hours", "Photo Best Hours", "Carousel Best Hours"),
        ("1st Best Hour", "19:00 (7 PM)", "17:00 (5 PM)", "20:00 (8 PM)", "18:00 (6 PM)"),
        ("2nd Best Hour", "17:00 (5 PM)", "18:00 (6 PM)", "08:00 (8 AM)", "12:00 (12 PM)"),
        ("3rd Best Hour", "22:00 (10 PM)", "19:00 (7 PM)", "21:00 (9 PM)", "09:00 (9 AM)")
    ]
    for r_idx, row_vals in enumerate(t83_data):
        row = t83.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10.5)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    if os.path.exists('report_assets/plot_124_9.png'):
        p_hm = doc.add_paragraph()
        p_hm.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_hm.paragraph_format.space_before = Pt(4)
        p_hm.paragraph_format.space_after = Pt(2)
        p_hm.add_run().add_picture('report_assets/plot_124_9.png', width=Inches(3.8))
        add_p("Fig. 8.2: Engagement Heatmap Across Weekdays and Day-Time Slots", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=10.5, space_before=2, space_after=6)

    add_p("Finding: Tuesday evening (18:00–20:00) represents the absolute global optimum for engagement (7.8%), followed closely by Sunday evening.", space_before=4)

    doc.add_page_break()

    # Q4
    add_subheading("8.4 Question 4: Breakout Post Identification via Lift Metric [CO4, BL4]")
    add_p("Objective: Identify virally amplified posts exceeding 3x the median views of their corresponding account and media format.")
    add_code_block("""df['lift'] = df.views / df.groupby(['user', 'kind'], observed=True).views.transform('median')
bo = df[df.lift > 3]
print(f'{len(bo)} breakouts ({len(bo)/len(df):.1%} of posts)')
display(bo.nlargest(5, 'lift')[['date', 'user', 'kind', 'hour', 'views', 'lift']])""")

    add_p("Table 8.4: Top Identified Breakout Posts with Algorithmic Lift Ratios", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    t84 = doc.add_table(rows=6, cols=6)
    t84.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t84)
    t84_data = [
        ("Post Date", "Creator Account", "Media Format", "Posting Hour", "Total Views", "Lift Multiplier"),
        ("2024-06-03", "Fit", "Reel", "17:00", "83,220", "8.09x"),
        ("2024-03-31", "Tech", "Reel", "21:00", "440,712", "7.71x"),
        ("2024-07-31", "Food", "Reel", "09:00", "253,225", "7.28x"),
        ("2024-04-12", "Trip", "Photo", "22:00", "360,417", "6.97x"),
        ("2024-09-07", "Fit", "Photo", "13:00", "27,346", "6.50x")
    ]
    for r_idx, row_vals in enumerate(t84_data):
        row = t84.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [3, 4, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10.5)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    add_p("Finding: Exactly 39 posts qualify as breakouts (6.8% of the total dataset). Reels constitute 56.4% of all breakouts, establishing short-form video as the principal vector for extreme non-linear algorithmic reach.", space_before=4)

    # Q5
    add_subheading("8.5 Question 5: Month-over-Month Growth & Scorecard Synthesis [CO5, BL5]")
    add_p("Objective: Quantify aggregate monthly impression growth and consolidate findings into an actionable strategic Format Scorecard.")
    add_code_block("""mo = ts.resample('ME').agg(posts=('views', 'size'), views=('views', 'sum'), eng=('eng', 'mean'))
mo['growth'] = mo.views.pct_change()
mo.index = mo.index.strftime('%b %Y')
display(mo.style.format({'views': '{:,.0f}', 'eng': '{:.2f}%', 'growth': '{:+.1%}'}))""")

    add_p("Table 8.5: Month-Over-Month Performance and Views Growth Trajectory (2024)", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2)
    t85 = doc.add_table(rows=11, cols=5)
    t85.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t85)
    t85_data = [
        ("Month", "Posts", "Aggregate Views", "Mean Engagement %", "MoM Reach Growth"),
        ("Jan 2024", "60", "5,799,024", "6.32%", "Baseline"),
        ("Feb 2024", "51", "3,317,669", "6.28%", "-42.8%"),
        ("Mar 2024", "54", "5,851,200", "6.69%", "+76.4%"),
        ("Apr 2024", "46", "3,607,278", "6.92%", "-38.3%"),
        ("May 2024", "61", "5,184,309", "6.68%", "+43.7%"),
        ("Jun 2024", "60", "4,152,145", "6.22%", "-19.9%"),
        ("Jul 2024", "59", "5,479,896", "6.61%", "+32.0%"),
        ("Aug 2024", "59", "4,417,286", "6.55%", "-19.4%"),
        ("Sep 2024", "71", "5,767,609", "6.44%", "+30.6%"),
        ("Oct 2024", "55", "3,355,160", "6.46%", "-41.8%")
    ]
    for r_idx, row_vals in enumerate(t85_data):
        row = t85.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [1, 2, 3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)
            if r_idx == 0:
                run.bold = True
                set_cell_background(cell, "F1F5F9")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 9: RESULTS AND DISCUSSION
    # ==========================================
    add_chapter_heading("CHAPTER 9\nResults and Discussion")
    add_subheading("9.1 Key Insights & Discovered Behavioral Patterns")
    add_p("The synthesis of exploratory analysis and graphical diagnostics provides concrete answers to our driving research inquiries:")
    add_p("1. Format Trade-Off (Reach vs Engagement): Reels and Carousels occupy complementary strategic roles. Reels maximize discovery at the top of the marketing funnel (59,256 median views), while Carousels maximize engagement depth at the bottom of the funnel (9.64% engagement rate, nearly double that of Reels).")
    add_p("2. Hashtag Saturation Curve: The empirical hashtag curve demonstrates non-linear performance. While 11–15 hashtags provide optimal categorical indexing without diluting content relevance, exceeding 20 hashtags correlates with severe reach penalties, likely reflecting automated spam suppression mechanisms.")
    add_p("3. Evening Temporal Synchronization: Peak engagement across all formats clusters tightly in the 17:00–21:00 evening window. The Tuesday evening surge (7.8% engagement) reflects mid-week consumer digital decompression, while Friday/Saturday exhibit lower engagement percentages despite sustained weekend view volumes.")

    add_subheading("9.2 Discussion of How Insights Answer Project Questions")
    add_p("The answers to the five core questions directly fulfill the course outcomes: CO1 and CO2 are satisfied through NumPy and Pandas grouped aggregations measuring format and hashtag efficiencies; CO3 and CO4 are demonstrated through time series resampling and mathematical lift thresholds; CO5 is achieved through the consolidated format scorecard and growth models.")

    add_subheading("9.3 Limitations of the Analysis and Potential Biases")
    add_p("1. Synthetic Distribution Bounds: While calibrated to real-world Instagram metrics, synthetic modeling cannot capture unpredictable external cultural moments or viral celebrity collaborations.\n2. Creator Niche Heterogeneity: Baseline reach varies substantially across niches (e.g. Memes vs Tech); normalized lift metrics mitigate but do not fully eliminate inter-account variance.\n3. Algorithmic Dynamism: Platform algorithmic updates frequently alter weightings between watch-time dwell, saves, shares, and comment length.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 10: CONCLUSION & RECOMMENDATIONS
    # ==========================================
    add_chapter_heading("CHAPTER 10\nConclusion and Recommendations")
    add_subheading("10.1 Summary of Main Findings")
    add_p("This Course End Project successfully demonstrated the utility of Python's scientific data stack for exploratory data analysis and visualization. By evaluating 600 publication records, the investigation verified that Reels lead in top-of-funnel reach (59,256 median views), Carousels lead in engagement depth (9.64%), 11–15 hashtags constitute the optimum metadata density, and Tuesday evenings represent the prime publication slot.")

    add_subheading("10.2 Actionable Recommendations for Creators & Marketers")
    add_p("• Deploy a 60/40 Hybrid Content Strategy: Dedicate 60% of output to short-form Reels to fuel audience acquisition, and 40% to informational Carousels to cultivate community retention.\n• Discipline Hashtag Usage: Standardize metadata to 11–15 hyper-specific niche hashtags and terminate the practice of 25+ generic hashtag stuffing.\n• Evening Scheduled Releases: Pre-schedule publication batches for the 18:00–20:00 window, prioritizing Tuesdays and Sundays.")

    add_subheading("10.3 Suggestions for Future Work")
    add_p("Future enhancements could integrate live Instagram Graph API pipelines, apply Natural Language Processing (NLP) sentiment scoring to user comments, and train predictive Gradient Boosting models (XGBoost) to forecast post virality prior to publication.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 11: REFERENCES
    # ==========================================
    add_chapter_heading("CHAPTER 11\nReferences")
    add_p("1. McKinney, W. (2022). Python for Data Analysis: Data Wrangling with pandas, NumPy, and Jupyter (3rd ed.). O'Reilly Media.")
    add_p("2. Hunter, J. D. (2007). Matplotlib: A 2D graphics environment. Computing in Science & Engineering, 9(3), 90-95.")
    add_p("3. Waskom, M. L. (2021). Seaborn: statistical data visualization. Journal of Open Source Software, 6(60), 3021.")
    add_p("4. Harris, C. R., Millman, K. J., van der Walt, S. J., et al. (2020). Array programming with NumPy. Nature, 585(7825), 357-362.")
    add_p("5. Meta for Developers. (2024). Instagram Graph API Reference Documentation. https://developers.facebook.com/docs/instagram-api")
    add_p("6. Chaitanya Bharathi Institute of Technology (A). (2024). Syllabus Curriculum for Exploratory Data Analysis and Visualization (22ADC31N), Dept of IT.")
    
    add_subheading("Project Web & Repository Links")
    add_p("• Project GitHub Repository: https://github.com/ronaksarda/edav-sem-end-project", bold=True)
    add_p("• Source Jupyter Notebook: https://github.com/ronaksarda/edav-sem-end-project/blob/main/instagram_engagement_analysis.ipynb", bold=True)
    add_p("• Institution Official Website: https://www.cbit.ac.in", bold=True)

    # Save document
    docx_path = "Ronak_Sarda_160125737123_EDAV_Course_End_Project_Report.docx"
    doc.save(docx_path)
    print(f"Report saved successfully as '{docx_path}'!")
    return docx_path

if __name__ == '__main__':
    create_report()
