import os
import struct
import zlib
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------------------------
# PURE PYTHON PNG RENDERER FOR ER DIAGRAM
# ---------------------------------------------------------------------------
class SimpleCanvas:
    def __init__(self, width, height, bg_color=(248, 250, 252)):
        self.w = width
        self.h = height
        self.bg = bg_color
        # 2D array of (r, g, b)
        self.pixels = [[bg_color for _ in range(width)] for _ in range(height)]

    def set_pixel(self, x, y, color):
        if 0 <= x < self.w and 0 <= y < self.h:
            self.pixels[y][x] = color

    def fill_rect(self, x, y, w, h, color):
        for cy in range(max(0, y), min(self.h, y + h)):
            for cx in range(max(0, x), min(self.w, x + w)):
                self.pixels[cy][cx] = color

    def draw_rect(self, x, y, w, h, color, thickness=1):
        for t in range(thickness):
            for cx in range(x, x + w):
                self.set_pixel(cx, y + t, color)
                self.set_pixel(cx, y + h - 1 - t, color)
            for cy in range(y, y + h):
                self.set_pixel(x + t, cy, color)
                self.set_pixel(x + w - 1 - t, cy, color)

    def draw_line_h(self, x1, x2, y, color, thickness=1):
        xmin, xmax = min(x1, x2), max(x1, x2)
        for cy in range(y, y + thickness):
            for cx in range(xmin, xmax + 1):
                self.set_pixel(cx, cy, color)

    def draw_line_v(self, x, y1, y2, color, thickness=1):
        ymin, ymax = min(y1, y2), max(y1, y2)
        for cx in range(x, x + thickness):
            for cy in range(ymin, ymax + 1):
                self.set_pixel(cx, cy, color)

    def draw_char(self, x, y, char, color, scale=1):
        # 5x7 basic bitmap font for letters and numbers
        # We define a few key characters for diagram labels
        patterns = {
            'A': ["01110","10001","11111","10001","10001"],
            'B': ["11110","10001","11110","10001","11110"],
            'C': ["01111","10000","10000","10000","01111"],
            'D': ["11110","10001","10001","10001","11110"],
            'E': ["11111","10000","11110","10000","11111"],
            'F': ["11111","10000","11110","10000","10000"],
            'G': ["01111","10000","10111","10001","01111"],
            'H': ["10001","10001","11111","10001","10001"],
            'I': ["11111","00100","00100","00100","11111"],
            'J': ["00111","00010","00010","10010","01100"],
            'K': ["10001","10010","11100","10010","10001"],
            'L': ["10000","10000","10000","10000","11111"],
            'M': ["10001","11011","10101","10001","10001"],
            'N': ["10001","11001","10101","10011","10001"],
            'O': ["01110","10001","10001","10001","01110"],
            'P': ["11110","10001","11110","10000","10000"],
            'Q': ["01110","10001","10101","10011","01111"],
            'R': ["11110","10001","11110","10010","10001"],
            'S': ["01111","10000","01110","00001","11110"],
            'T': ["11111","00100","00100","00100","00100"],
            'U': ["10001","10001","10001","10001","01110"],
            'V': ["10001","10001","10001","01010","00100"],
            'W': ["10001","10001","10101","11011","10001"],
            'X': ["10001","01010","00100","01010","10001"],
            'Y': ["10001","01010","00100","00100","00100"],
            'Z': ["11111","00010","00100","01000","11111"],
            '0': ["01110","10011","10101","11001","01110"],
            '1': ["00100","01100","00100","00100","01110"],
            '2': ["01110","10001","00110","01000","11111"],
            '3': ["11110","00001","01110","00001","11110"],
            '4': ["10010","10010","11111","00010","00010"],
            '5': ["11111","10000","11110","00001","11110"],
            '6': ["01110","10000","11110","10001","01110"],
            '7': ["11111","00010","00100","01000","01000"],
            '8': ["01110","10001","01110","10001","01110"],
            '9': ["01110","10001","01111","00001","01110"],
            ':': ["00000","00100","00000","00100","00000"],
            '-': ["00000","00000","11111","00000","00000"],
            '_': ["00000","00000","00000","00000","11111"],
            ' ': ["00000","00000","00000","00000","00000"],
            '(': ["00110","01000","01000","01000","00110"],
            ')': ["01100","00010","00010","00010","01100"],
            '.': ["00000","00000","00000","00000","00100"]
        }
        ch = char.upper()
        grid = patterns.get(ch, patterns[' '])
        for row_i, row in enumerate(grid):
            for col_i, bit in enumerate(row):
                if bit == '1':
                    for sy in range(scale):
                        for sx in range(scale):
                            self.set_pixel(x + col_i*scale + sx, y + row_i*scale + sy, color)

    def draw_text(self, x, y, text, color, scale=1):
        cur_x = x
        for c in text:
            self.draw_char(cur_x, y, c, color, scale)
            cur_x += 6 * scale

    def to_png(self):
        def chunk(tag, data):
            return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data) & 0xffffffff)
        ihdr = struct.pack('>IIBBBBB', self.w, self.h, 8, 2, 0, 0, 0)
        raw_data = bytearray()
        for y in range(self.h):
            raw_data.append(0)
            for x in range(self.w):
                r, g, b = self.pixels[y][x]
                raw_data.extend([r, g, b])
        idat = zlib.compress(bytes(raw_data), 9)
        return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', ihdr) + chunk(b'IDAT', idat) + chunk(b'IEND', b'')

def render_er_diagram_image(out_path):
    canvas = SimpleCanvas(1000, 720, bg_color=(248, 250, 252))
    
    # Outer Border
    canvas.draw_rect(10, 10, 980, 700, (203, 213, 225), thickness=2)
    
    # Diagram Header Title Block
    canvas.fill_rect(12, 12, 976, 50, (30, 58, 138))
    canvas.draw_text(30, 28, "CHANNEL WISE MARKETING PERFORMANCE - ENTITY RELATIONSHIP DIAGRAM", (255, 255, 255), scale=2)
    
    # Entities helper
    def draw_entity(x, y, w, h, name, columns, is_fact=False):
        # Header bar
        header_color = (29, 78, 216) if not is_fact else (185, 28, 28) # Blue for Dim, Crimson for Fact
        canvas.fill_rect(x, y, w, 28, header_color)
        canvas.draw_rect(x, y, w, h, (148, 163, 184), thickness=2)
        canvas.draw_text(x + 12, y + 8, name, (255, 255, 255), scale=2)
        # Body
        canvas.fill_rect(x + 2, y + 28, w - 4, h - 30, (255, 255, 255))
        # Columns
        cur_y = y + 36
        for col_name, key_type in columns:
            key_color = (185, 28, 28) if key_type == "PK" else ((29, 78, 216) if key_type == "FK" else (100, 116, 139))
            if key_type != "-":
                canvas.draw_text(x + 12, cur_y, f"[{key_type}]", key_color, scale=1)
            canvas.draw_text(x + 50, cur_y, col_name, (30, 41, 59), scale=1)
            cur_y += 15

    # 1. dim_channel (Top Center)
    ch_cols = [
        ("channel_id (INT)", "PK"),
        ("channel_name (TEXT)", "-"),
        ("channel_category (TEXT)", "-"),
        ("default_cpc_benchmark (REAL)", "-"),
        ("benchmark_roas (REAL)", "-"),
        ("benchmark_cac (REAL)", "-")
    ]
    draw_entity(370, 80, 260, 145, "DIM_CHANNEL", ch_cols)

    # 2. dim_campaign (Middle Left)
    cmp_cols = [
        ("campaign_id (INT)", "PK"),
        ("channel_id (INT)", "FK"),
        ("campaign_name (TEXT)", "-"),
        ("campaign_type (TEXT)", "-"),
        ("target_roas (REAL)", "-"),
        ("target_cac (REAL)", "-"),
        ("start_date (TEXT)", "-"),
        ("status (TEXT)", "-")
    ]
    draw_entity(50, 270, 260, 175, "DIM_CAMPAIGN", cmp_cols)

    # 3. dim_customer (Middle Right)
    cust_cols = [
        ("customer_id (INT)", "PK"),
        ("customer_name (TEXT)", "-"),
        ("age_group (TEXT)", "-"),
        ("gender (TEXT)", "-"),
        ("city_tier (TEXT)", "-"),
        ("city_name (TEXT)", "-"),
        ("acquisition_channel_id (INT)", "FK"),
        ("signup_date (TEXT)", "-")
    ]
    draw_entity(690, 270, 260, 175, "DIM_CUSTOMER", cust_cols)

    # 4. fact_daily_campaign_performance (Bottom Left)
    fact_cols = [
        ("fact_id (INT)", "PK"),
        ("date (TEXT)", "-"),
        ("campaign_id (INT)", "FK"),
        ("day_of_week / month (INT)", "-"),
        ("spend (REAL)", "-"),
        ("impressions / clicks (INT)", "-"),
        ("conversions (INT)", "-"),
        ("revenue (REAL)", "-"),
        ("utm_source / med / camp", "-"),
        ("is_outlier / reason", "-")
    ]
    draw_entity(50, 480, 280, 205, "FACT_DAILY_PERF", fact_cols, is_fact=True)

    # 5. fact_customer_touchpoint (Bottom Right)
    touch_cols = [
        ("touchpoint_id (INT)", "PK"),
        ("journey_id (INT)", "-"),
        ("customer_id (INT)", "FK"),
        ("channel_id (INT)", "FK"),
        ("touchpoint_sequence (INT)", "-"),
        ("touchpoint_timestamp (TEXT)", "-"),
        ("is_conversion (INT)", "-"),
        ("order_value (REAL)", "-")
    ]
    draw_entity(670, 480, 280, 180, "FACT_TOUCHPOINT", touch_cols, is_fact=True)

    # CONNECTING LINES & CARDINALITIES
    line_color = (71, 85, 105)
    
    # REL-01: dim_channel -> dim_campaign
    canvas.draw_line_v(400, 225, 250, line_color, thickness=2)
    canvas.draw_line_h(180, 400, 250, line_color, thickness=2)
    canvas.draw_line_v(180, 250, 270, line_color, thickness=2)
    canvas.draw_text(260, 235, "1 : N (FUNDS)", (29, 78, 216), scale=1)

    # REL-02: dim_channel -> dim_customer
    canvas.draw_line_v(600, 225, 250, line_color, thickness=2)
    canvas.draw_line_h(600, 820, 250, line_color, thickness=2)
    canvas.draw_line_v(820, 250, 270, line_color, thickness=2)
    canvas.draw_text(650, 235, "1 : N (ACQUIRES)", (29, 78, 216), scale=1)

    # REL-03: dim_campaign -> fact_daily_campaign_performance
    canvas.draw_line_v(180, 445, 480, line_color, thickness=2)
    canvas.draw_text(190, 455, "1 : N (RECORDS)", (185, 28, 28), scale=1)

    # REL-04: dim_customer -> fact_customer_touchpoint
    canvas.draw_line_v(820, 445, 480, line_color, thickness=2)
    canvas.draw_text(830, 455, "1 : N (TOUCHES)", (185, 28, 28), scale=1)

    # REL-05: dim_channel -> fact_customer_touchpoint (central bypass)
    canvas.draw_line_v(500, 225, 560, line_color, thickness=2)
    canvas.draw_line_h(500, 670, 560, line_color, thickness=2)
    canvas.draw_text(515, 545, "1 : N (ATTRIBUTION)", (29, 78, 216), scale=1)

    png_bytes = canvas.to_png()
    with open(out_path, 'wb') as f:
        f.write(png_bytes)
    print(f"ER Diagram PNG generated at: {out_path} ({len(png_bytes)} bytes)")

# ---------------------------------------------------------------------------
# WORD DOCUMENT GENERATION (docs/09_ER_Diagram.docx)
# ---------------------------------------------------------------------------
def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table(table, header_bg="1E3A8A", alt_bg="F1F5F9"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if i == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            for cell in row.cells:
                set_cell_background(cell, header_bg)
                set_cell_margins(cell, 120, 120, 140, 140)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.size = Pt(9.5)
        else:
            bg = alt_bg if i % 2 == 1 else "FFFFFF"
            for cell in row.cells:
                set_cell_background(cell, bg)
                set_cell_margins(cell, 90, 90, 130, 130)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)
                        run.font.color.rgb = RGBColor(30, 41, 59)

def generate_step9_doc():
    doc = docx.Document()
    
    # 1-inch margins
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Title Block
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("ENTITY RELATIONSHIP (ER) DIAGRAM & SCHEMA REPORT\n")
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Project: Channel Wise Marketing Performance Analysis\nStep 9 Milestone | Data Architecture & Structural Invariants\n")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Metadata Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Document Identifier", "DOC-ERD-V1.0 (Step 9 Schema Artifact)"),
        ("Author / Data Architect", "Vaishnavi Singh (Roll No: 2503201450157, MCA 2nd Yr DA, ABES)"),
        ("Project Mentor", "Prof. Kalpana Dwivedi (ABES Engineering College)"),
        ("Diagram Notation", "Crow's Foot Entity Relationship Notation (Mermaid erDiagram)"),
        ("Target Database Engine", "SQLite 3.x (Embedded file: data/app.db)")
    ]
    for idx, (k, v) in enumerate(meta_info):
        ck, cv = meta_table.rows[idx].cells
        ck.text = k
        cv.text = v
        set_cell_background(ck, "F1F5F9")
        set_cell_background(cv, "FFFFFF")
        set_cell_margins(ck, 60, 60, 100, 100)
        set_cell_margins(cv, 60, 60, 100, 100)
        ck.paragraphs[0].runs[0].font.bold = True
        ck.paragraphs[0].runs[0].font.size = Pt(9.5)
        cv.paragraphs[0].runs[0].font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # 1. Mermaid erDiagram Code Block
    h1 = doc.add_heading(level=1)
    h1.add_run("1. Mermaid ER Diagram Code (erDiagram)").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Below is the complete, valid Mermaid 'erDiagram' source code defining every table, attribute datatype, "
        "primary/foreign key constraint, and verb-labeled relationship. This code can be copied directly into "
        "https://mermaid.live to render or export interactive vector/raster graphics:"
    )

    mermaid_code = (
        "erDiagram\n"
        "    DIM_CHANNEL {\n"
        "        int channel_id PK\n"
        "        string channel_name\n"
        "        string channel_category\n"
        "        float default_cpc_benchmark\n"
        "        float benchmark_roas\n"
        "        float benchmark_cac\n"
        "    }\n\n"
        "    DIM_CAMPAIGN {\n"
        "        int campaign_id PK\n"
        "        int channel_id FK\n"
        "        string campaign_name\n"
        "        string campaign_type\n"
        "        float target_roas\n"
        "        float target_cac\n"
        "        string start_date\n"
        "        string status\n"
        "    }\n\n"
        "    DIM_CUSTOMER {\n"
        "        int customer_id PK\n"
        "        string customer_name\n"
        "        string age_group\n"
        "        string gender\n"
        "        string city_tier\n"
        "        string city_name\n"
        "        int acquisition_channel_id FK\n"
        "        string signup_date\n"
        "    }\n\n"
        "    FACT_DAILY_CAMPAIGN_PERFORMANCE {\n"
        "        int fact_id PK\n"
        "        string date\n"
        "        int campaign_id FK\n"
        "        int day_of_week\n"
        "        int month\n"
        "        float spend\n"
        "        int impressions\n"
        "        int clicks\n"
        "        int conversions\n"
        "        float revenue\n"
        "        string utm_source\n"
        "        string utm_medium\n"
        "        string utm_campaign\n"
        "        int is_outlier\n"
        "        string outlier_reason\n"
        "    }\n\n"
        "    FACT_CUSTOMER_TOUCHPOINT {\n"
        "        int touchpoint_id PK\n"
        "        int journey_id\n"
        "        int customer_id FK\n"
        "        int channel_id FK\n"
        "        int touchpoint_sequence\n"
        "        string touchpoint_timestamp\n"
        "        int is_conversion\n"
        "        float order_value\n"
        "    }\n\n"
        "    DIM_CHANNEL ||--o{ DIM_CAMPAIGN : \"finances_and_funds\"\n"
        "    DIM_CHANNEL ||--o{ DIM_CUSTOMER : \"acquires_first_touch\"\n"
        "    DIM_CAMPAIGN ||--o{ FACT_DAILY_CAMPAIGN_PERFORMANCE : \"generates_daily_log\"\n"
        "    DIM_CUSTOMER ||--o{ FACT_CUSTOMER_TOUCHPOINT : \"initiates_journey_step\"\n"
        "    DIM_CHANNEL ||--o{ FACT_CUSTOMER_TOUCHPOINT : \"receives_touchpoint_visit\"\n"
    )

    p_code = doc.add_paragraph()
    r_code = p_code.add_run(mermaid_code)
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(8.5)
    p_code.paragraph_format.left_indent = Inches(0.2)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. Visual ER Diagram PNG Insertion
    h2 = doc.add_heading(level=1)
    h2.add_run("2. Graphical Entity Relationship Diagram").font.color.rgb = RGBColor(30, 58, 138)

    png_path = os.path.abspath("docs/09_ER_Diagram.png")
    if os.path.exists(png_path):
        doc.add_paragraph("The architectural schema rendered below illustrates entity relationships, cardinalities, and key constraints:")
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_img.add_run()
        run_img.add_picture(png_path, width=Inches(6.2))
        
        cap_p = doc.add_paragraph()
        cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap_p.add_run("Figure 9.1: Entity Relationship Architecture for Marketing Performance Platform")
        cap_run.font.size = Pt(9)
        cap_run.font.italic = True
        cap_run.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. Plain-English Relationship Statements
    h3 = doc.add_heading(level=1)
    h3.add_run("3. Plain-English Relationship Descriptions").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "Each relational cardinality in the schema represents an authoritative business dependency:"
    )

    rel_p1 = doc.add_paragraph()
    rel_p1.add_run("1. Channel to Campaign (1 : N): ").bold = True
    rel_p1.add_run("One marketing channel (such as Google Ads or Meta Ads) finances and hosts many marketing campaigns, but each campaign belongs to exactly one marketing channel.")

    rel_p2 = doc.add_paragraph()
    rel_p2.add_run("2. Channel to Customer (1 : N): ").bold = True
    rel_p2.add_run("One marketing channel acquires many customers via first-touch discovery, but each customer records exactly one initial acquisition channel upon registration.")

    rel_p3 = doc.add_paragraph()
    rel_p3.add_run("3. Campaign to Daily Performance Fact (1 : N): ").bold = True
    rel_p3.add_run("One marketing campaign generates many daily performance telemetry records over time, but each daily performance record tracks metrics for exactly one specific campaign.")

    rel_p4 = doc.add_paragraph()
    rel_p4.add_run("4. Customer to Touchpoint Fact (1 : N): ").bold = True
    rel_p4.add_run("One customer initiates many sequential touchpoint interactions across their multi-touch purchase journey, but each touchpoint event belongs to exactly one customer.")

    rel_p5 = doc.add_paragraph()
    rel_p5.add_run("5. Channel to Touchpoint Fact (1 : N): ").bold = True
    rel_p5.add_run("One marketing channel receives many interaction touchpoint visits from various customer journeys, but each touchpoint visit occurs on exactly one marketing channel.")

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 4. Business Rules Enforced by the Design
    h4 = doc.add_heading(level=1)
    h4.add_run("4. Business Rules Enforced by Database Constraints").font.color.rgb = RGBColor(30, 58, 138)

    rules_table = doc.add_table(rows=7, cols=3)
    rules_headers = ["Rule ID", "Target Table & Column(s)", "Database Mechanism & Enforcement Behavior"]
    for j, rh in enumerate(rules_headers):
        rules_table.rows[0].cells[j].text = rh

    rules_data = [
        ("BR-DB-01", "fact_daily_campaign_performance (clicks, impressions)", "CHECK constraint (clicks >= 0 AND clicks <= impressions): Enforces logical physical impossibility of recording more clicks than ad impressions."),
        ("BR-DB-02", "fact_daily_campaign_performance (conversions, clicks)", "CHECK constraint (conversions >= 0 AND conversions <= clicks): Guarantees completed conversion purchases never exceed recorded clicks."),
        ("BR-DB-03", "fact_daily_campaign_performance (spend, revenue)", "CHECK constraint (spend >= 0.0 AND revenue >= 0.0): Strictly prohibits negative financial values at the database engine level."),
        ("BR-DB-04", "dim_channel & dim_campaign (channel_name, campaign_name)", "UNIQUE constraint: Prevents duplicate channel or campaign naming, eliminating ambiguity in reporting and filter queries."),
        ("BR-DB-05", "All Foreign Keys (channel_id, campaign_id, customer_id)", "FOREIGN KEY ... ON DELETE RESTRICT: Prevents deletion of parent master entities if active performance logs or campaigns exist."),
        ("BR-DB-06", "fact_customer_touchpoint (touchpoint_sequence)", "CHECK constraint (touchpoint_sequence >= 1): Guarantees valid positive ordering in multi-touch customer conversion paths.")
    ]
    for r_idx, r_data in enumerate(rules_data, start=1):
        for c_idx, val in enumerate(r_data):
            rules_table.rows[r_idx].cells[c_idx].text = val
    style_table(rules_table)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 5. Design Problems Identified & Fixes Applied
    h5 = doc.add_heading(level=1)
    h5.add_run("5. Database Design Problems Identified & Engineering Fixes").font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph(
        "During architectural review, four structural design pitfalls were evaluated and systematically resolved:"
    )

    issues = [
        ("Problem 1: Many-to-Many Relationship Between Customers and Marketing Channels",
         "A customer interacts with multiple channels prior to purchasing, and each channel reaches multiple customers. A direct M:N relationship creates update anomalies and data duplication.",
         "ENGINEERING FIX: Introduced 'fact_customer_touchpoint' as an associative bridge entity capturing (customer_id, channel_id, journey_id, touchpoint_sequence, timestamp). This decomposes the M:N relationship into two clean 1:N relationships."),

        ("Problem 2: Granularity Mismatch Between Macro Ad Telemetry and Micro Journey Orders",
         "Ad platforms report daily aggregated spend, clicks, and impressions per campaign, whereas attribution models require event-level individual user journey sequences. Forcing these into a single fact table causes severe null proliferation and Cartesian multiplication errors.",
         "ENGINEERING FIX: Separated transactional telemetry into two dedicated fact tables: 'fact_daily_campaign_performance' (macro grain) and 'fact_customer_touchpoint' (micro grain)."),

        ("Problem 3: Redundant Computed Ratio Columns in Physical Storage (3NF Violation)",
         "Storing derived financial metrics (ROAS, CAC, CTR, CVR, Profit) physically in the fact table wastes storage and risks data drift if underlying spend or revenue records are updated.",
         "ENGINEERING FIX: Stored only atomic, uncalculated base facts in SQLite tables, and engineered a 38-column joined view ('vw_marketing_analysis') that computes ratios dynamically in SQL using NULLIF protection."),

        ("Problem 4: Division by Zero Crashes During Ratio Computation",
         "Days or campaigns with zero impressions, zero clicks, or zero conversions crash division queries ('spend / clicks') with fatal SQLite/Python exceptions.",
         "ENGINEERING FIX: Wrapped all divisors in 'NULLIF(denominator, 0)' across all SQL queries and view definitions, ensuring zero divisors safely evaluate to NULL without breaking queries.")
    ]

    for title, desc, fix in issues:
        p_iss = doc.add_paragraph()
        r_iss_t = p_iss.add_run(f"• {title}\n")
        r_iss_t.bold = True
        r_iss_t.font.color.rgb = RGBColor(30, 58, 138)
        
        p_desc = doc.add_paragraph()
        p_desc.add_run("  Identified Issue: ").bold = True
        p_desc.add_run(desc + "\n")
        p_desc.add_run("  Architectural Resolution: ").bold = True
        p_desc.add_run(fix)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Save Word document
    out_word = os.path.abspath("docs/09_ER_Diagram.docx")
    doc.save(out_word)
    print(f"Saved: {out_word}")

if __name__ == "__main__":
    png_file = os.path.abspath("docs/09_ER_Diagram.png")
    render_er_diagram_image(png_file)
    generate_step9_doc()
    # cleanup temp test files
    for tmp in ['test.png', 'test_doc.docx']:
        if os.path.exists(tmp):
            os.remove(tmp)
