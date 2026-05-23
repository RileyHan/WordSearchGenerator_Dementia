import random
import string
import os
from openpyxl import load_workbook  # type: ignore
from docx import Document  # type: ignore
from docx.shared import Pt, Inches, RGBColor  # type: ignore
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER  # type: ignore
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE  # type: ignore
from docx.enum.section import WD_ORIENT  # type: ignore
from docx.oxml import OxmlElement  # type: ignore
from docx.oxml.ns import qn  # type: ignore

ROWS = 8
COLS = 12

EXCEL_PATH = r"C:\Users\riley\OneDrive\Desktop\wordsearch_booklet\Book2.xlsx"

TITLE_FONT_SIZE = 26
GRID_FONT_SIZE = 36
WORD_FONT_SIZE = 20

GRID_CELL_WIDTH = 56
GRID_ROW_HEIGHT = 48
WORD_BANK_ROW_HEIGHT = 20
WORD_BANK_COL_WIDTH = 150

TOP_MARGIN = 0.2
BOTTOM_MARGIN = 0.2
LEFT_MARGIN = 0.25
RIGHT_MARGIN = 0.25

ANSWER_TITLE_SIZE = 16
ANSWER_GRID_FONT_SIZE = 14
ANSWER_CELL_SIZE = 23
ANSWER_BLOCK_HEIGHT = 235

def load_puzzles_from_excel(file_path):
    wb = load_workbook(file_path, data_only=True)
    ws = wb.worksheets[0]  # sheet 1 only

    puzzle_titles = []
    puzzles = []

    for excel_row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or row[0] is None:
            continue

        title = str(row[0]).strip()
        if title == "":
            continue

        words = []
        for cell in row[1:]:
            if cell is None:
                continue

            word = str(cell).strip().upper()
            if word == "":
                continue

            if len(word) > COLS:
                print(f"Skipping word '{word}' in row {excel_row_num}: too long for {COLS} columns")
                continue

            words.append((word, 1, 1))

        if len(words) > ROWS:
            print(f"Row {excel_row_num} has more than {ROWS} words. Extra words will be ignored.")
            words = words[:ROWS]

        if words:
            puzzle_titles.append(title)
            puzzles.append(words)

    return puzzle_titles, puzzles

def build_grid(words):
    grid = [["" for _ in range(COLS)] for _ in range(ROWS)]

    shuffled_rows = list(range(ROWS))
    random.shuffle(shuffled_rows)

    placed_words = []

    for index, (word, _, _) in enumerate(words):
        r = shuffled_rows[index]
        max_start = COLS - len(word)
        c = random.randint(0, max_start)

        for i, ch in enumerate(word):
            grid[r][c + i] = ch

        placed_words.append((word, r + 1, c + 1))

    for r in range(ROWS):
        for c in range(COLS):
            if grid[r][c] == "":
                grid[r][c] = random.choice(string.ascii_uppercase)

    return grid, placed_words

def get_answer_positions(words):
    positions = set()
    for word, r, c in words:
        r -= 1
        c -= 1
        for i in range(len(word)):
            if 0 <= r < ROWS and 0 <= c + i < COLS:
                positions.add((r, c + i))
    return positions

def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr

    borders = OxmlElement("w:tblBorders")
    for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        elem = OxmlElement(f"w:{edge}")
        elem.set(qn("w:val"), "nil")
        borders.append(elem)

    tblPr.append(borders)

def style_grid_cell(cell):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.size = Pt(GRID_FONT_SIZE)
            run.font.name = "Arial"

def style_word_cell(cell):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.size = Pt(WORD_FONT_SIZE)
            run.font.name = "Arial"

def add_title(doc, title_text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    run = p.add_run(f'"{title_text}"')
    run.bold = True
    run.font.size = Pt(TITLE_FONT_SIZE)
    run.font.name = "Arial Rounded MT Bold"

def add_spacer(doc, points_after):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(points_after)

def setup_footer():
    section = doc.sections[0]
    footer = section.footer
    paragraph = footer.paragraphs[0]

    p = paragraph._element
    for child in list(p):
        p.remove(child)

    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.left_indent = Inches(1)
    paragraph.paragraph_format.right_indent = Inches(0.2)

    tab_stops = paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Inches(5.0), WD_TAB_ALIGNMENT.CENTER, WD_TAB_LEADER.SPACES)
    tab_stops.add_tab_stop(Inches(9.6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)

    left_run = paragraph.add_run('"Dementia Friendly Wordsearch"')
    left_run.font.name = "Arial"
    left_run.font.size = Pt(10)

    paragraph.add_run("\t")

    page_run = paragraph.add_run()
    page_run.font.name = "Arial"
    page_run.font.size = Pt(10)

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")

    page_run._r.append(fld_char_begin)
    page_run._r.append(instr_text)
    page_run._r.append(fld_char_end)

    paragraph.add_run("\t")

    right_run = paragraph.add_run('"Puzzles by Riley"')
    right_run.font.name = "Arial"
    right_run.font.size = Pt(10)

desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
folder = os.path.join(desktop, "wordsearch_booklet")
os.makedirs(folder, exist_ok=True)

PUZZLE_TITLES, PUZZLES = load_puzzles_from_excel(EXCEL_PATH)

doc = Document()

section = doc.sections[0]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width, section.page_height = section.page_height, section.page_width
section.left_margin = Inches(LEFT_MARGIN)
section.right_margin = Inches(RIGHT_MARGIN)
section.top_margin = Inches(TOP_MARGIN)
section.bottom_margin = Inches(BOTTOM_MARGIN)

setup_footer()

answer_keys = []

for i, puzzle in enumerate(PUZZLES, start=1):
    title = PUZZLE_TITLES[i - 1]

    grid, placed_words = build_grid(puzzle)
    answer_keys.append((title, grid, get_answer_positions(placed_words)))

    add_spacer(doc, 2)
    add_title(doc, title)
    add_spacer(doc, 3)

    table = doc.add_table(rows=ROWS, cols=COLS)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(table)

    for row in table.rows:
        row.height = Pt(GRID_ROW_HEIGHT)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for cell in row.cells:
            cell.width = Pt(GRID_CELL_WIDTH)

    for r in range(ROWS):
        for c in range(COLS):
            cell = table.cell(r, c)
            cell.text = grid[r][c]
            style_grid_cell(cell)

    add_spacer(doc, 3)

    words_only = [word for word, _, _ in puzzle]
    word_bank = doc.add_table(rows=2, cols=4)
    word_bank.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(word_bank)

    for row in word_bank.rows:
        row.height = Pt(WORD_BANK_ROW_HEIGHT)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for cell in row.cells:
            cell.width = Pt(WORD_BANK_COL_WIDTH)

    index = 0
    for r in range(2):
        for c in range(4):
            cell = word_bank.cell(r, c)
            if index < len(words_only):
                cell.text = words_only[index]
                style_word_cell(cell)
            index += 1

    if i < len(PUZZLES):
        doc.add_page_break()

if len(answer_keys) > 0:
    doc.add_page_break()

for start in range(0, len(answer_keys), 4):
    if start > 0:
        doc.add_page_break()

    block = answer_keys[start:start + 4]

    answer_page = doc.add_table(rows=2, cols=2)
    answer_page.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(answer_page)

    for row in answer_page.rows:
        row.height = Pt(ANSWER_BLOCK_HEIGHT)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    slot = 0
    for rr in range(2):
        for cc in range(2):
            outer_cell = answer_page.cell(rr, cc)

            if slot < len(block):
                title, grid, positions = block[slot]

                top_space = outer_cell.paragraphs[0]
                top_space.paragraph_format.space_before = Pt(0)
                top_space.paragraph_format.space_after = Pt(10)

                title_p = outer_cell.add_paragraph()
                title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                title_p.paragraph_format.space_before = Pt(0)
                title_p.paragraph_format.space_after = Pt(8)

                title_run = title_p.add_run(f'Answer Key: "{title}"')
                title_run.bold = True
                title_run.font.size = Pt(ANSWER_TITLE_SIZE)
                title_run.font.name = "Arial"

                mini = outer_cell.add_table(rows=ROWS, cols=COLS)
                mini.alignment = WD_TABLE_ALIGNMENT.CENTER
                remove_table_borders(mini)

                for row in mini.rows:
                    row.height = Pt(ANSWER_CELL_SIZE)
                    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                    for mini_cell in row.cells:
                        mini_cell.width = Pt(ANSWER_CELL_SIZE)

                for r in range(ROWS):
                    for c in range(COLS):
                        mini_cell = mini.cell(r, c)
                        para = mini_cell.paragraphs[0]
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)

                        run = para.add_run(grid[r][c])
                        run.font.size = Pt(ANSWER_GRID_FONT_SIZE)
                        run.font.name = "Arial"

                        if (r, c) in positions:
                            run.font.color.rgb = RGBColor(255, 0, 0)

            slot += 1

base_name = "wordsearch_booklet"
file_path = os.path.join(folder, f"{base_name}.docx")

counter = 1
while True:
    try:
        doc.save(file_path)
        break
    except PermissionError:
        file_path = os.path.join(folder, f"{base_name}_{counter}.docx")
        counter += 1

print("Saved booklet to:")
print(file_path)