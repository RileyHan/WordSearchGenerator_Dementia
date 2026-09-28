import os
import random
import string
import subprocess
from copy import deepcopy

from openpyxl import load_workbook  # type: ignore
from docx import Document  # type: ignore
from docx.enum.section import WD_ORIENT, WD_SECTION_START  # type: ignore
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT  # type: ignore
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER  # type: ignore
from docx.oxml import OxmlElement  # type: ignore
from docx.oxml.ns import qn  # type: ignore
from docx.shared import Inches, Pt, RGBColor  # type: ignore

ROWS = 8
COLS = 12

BASE_FOLDER = r"C:\Users\riley\OneDrive\Desktop\wordsearch_booklet"
EXCEL_PATH = os.path.join(BASE_FOLDER, "Book2.xlsx")
OUTPUT_FOLDER = BASE_FOLDER

TITLE_FONT_SIZE = 26
GRID_FONT_SIZE = 36
WORD_FONT_SIZE = 18
GRID_CELL_WIDTH = 58
GRID_ROW_HEIGHT = 42
WORD_BANK_ROW_HEIGHT = 24
WORD_BANK_COL_WIDTH = 155

LANDSCAPE_TOP_MARGIN = 0.25
LANDSCAPE_BOTTOM_MARGIN = 0.25
LANDSCAPE_LEFT_MARGIN = 0.25
LANDSCAPE_RIGHT_MARGIN = 0.25

ANSWER_TITLE_SIZE = 16
ANSWER_GRID_FONT_SIZE = 14
ANSWER_CELL_SIZE = 24
ANSWER_BLOCK_HEIGHT = 255

TOC_TITLE_SIZE = 20
TOC_HEADING_SIZE = 14
TOC_ENTRY_SIZE = 12
TOC_ENTRIES_PER_PAGE = 28

SCRIPT_FOLDER = os.path.dirname(os.path.abspath(__file__))
COVER_TEMPLATE_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Cover.docx")
COPYRIGHT_TEMPLATE_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Copyright.docx")
INSTRUCTIONS_TEMPLATE_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Instructions.docx")
BACK_COVER_TEMPLATE_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Back Cover.docx")
COVER_SUBTITLE_TEXT = "Simple, calming, and frustration-free word searches"


def verify_required_paths():
    required_files = [
        EXCEL_PATH,
        COVER_TEMPLATE_PATH,
        COPYRIGHT_TEMPLATE_PATH,
        INSTRUCTIONS_TEMPLATE_PATH,
        BACK_COVER_TEMPLATE_PATH,
    ]
    for file_path in required_files:
        print(f"  {file_path}")
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"Missing file: {file_path}")


def load_puzzles_from_excel(file_path):
    workbook = load_workbook(file_path, data_only=True)
    worksheet = workbook.worksheets[0]
    titles, puzzles = [], []

    for row_number, row in enumerate(
        worksheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        if not row or row[0] is None:
            continue

        title = str(row[0]).strip()
        words = []
        for value in row[1:]:
            if value is None:
                continue
            word = str(value).strip().upper()
            if not word:
                continue
            if len(word) > COLS:
                print(f"Skipping word '{word}' in row {row_number}: too long for {COLS} columns")
                continue
            words.append((word, 1, 1))

        if len(words) > ROWS:
            print(f"Row {row_number} has more than {ROWS} words. Extra words will be ignored.")
            words = words[:ROWS]

        if title and words:
            titles.append(title)
            puzzles.append(words)

    return titles, puzzles


def build_grid(words):
    grid = [["" for _ in range(COLS)] for _ in range(ROWS)]
    available_rows = list(range(ROWS))
    random.shuffle(available_rows)
    placed_words = []

    for index, (word, _, _) in enumerate(words):
        row = available_rows[index]
        column = random.randint(0, COLS - len(word))
        for offset, character in enumerate(word):
            grid[row][column + offset] = character
        placed_words.append((word, row + 1, column + 1))

    for row in range(ROWS):
        for column in range(COLS):
            if not grid[row][column]:
                grid[row][column] = random.choice(string.ascii_uppercase)

    return grid, placed_words


def get_answer_positions(placed_words):
    positions = set()
    for word, row, column in placed_words:
        row -= 1
        column -= 1
        for offset in range(len(word)):
            if 0 <= row < ROWS and 0 <= column + offset < COLS:
                positions.add((row, column + offset))
    return positions


def build_puzzle_layouts(titles, puzzles):
    layouts = []
    for title, puzzle in zip(titles, puzzles):
        grid, placed_words = build_grid(puzzle)
        layouts.append(
            {
                "title": title,
                "words": puzzle,
                "grid": grid,
                "answer_positions": get_answer_positions(placed_words),
            }
        )
    return layouts


def remove_table_borders(table):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        borders.append(element)
    table._tbl.tblPr.append(borders)


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


def add_spacer(doc, points_after):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(points_after)


def add_title(doc, title):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(12)
    run = paragraph.add_run(f'"{title}"')
    run.bold = True
    run.font.size = Pt(TITLE_FONT_SIZE)
    run.font.name = "Arial Rounded MT Bold"


def clear_footer(paragraph):
    element = paragraph._element
    for child in list(element):
        element.remove(child)


def unlink_footer(section):
    section.footer.is_linked_to_previous = False


def ensure_page_number_settings(section):
    element = section._sectPr.find(qn("w:pgNumType"))
    if element is None:
        element = OxmlElement("w:pgNumType")
        section._sectPr.append(element)
    return element


def set_page_number_start(section, number):
    ensure_page_number_settings(section).set(qn("w:start"), str(number))


def continue_page_numbering(section):
    element = section._sectPr.find(qn("w:pgNumType"))
    if element is not None:
        section._sectPr.remove(element)


def setup_footer(section, left_text, right_text):
    unlink_footer(section)
    paragraph = section.footer.paragraphs[0]
    clear_footer(paragraph)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.left_indent = Inches(0.35)
    paragraph.paragraph_format.right_indent = Inches(0.35)
    stops = paragraph.paragraph_format.tab_stops
    stops.add_tab_stop(Inches(5.35), WD_TAB_ALIGNMENT.CENTER, WD_TAB_LEADER.SPACES)
    stops.add_tab_stop(Inches(9.95), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)

    left = paragraph.add_run(left_text)
    left.font.name = "Arial"
    left.font.size = Pt(10)
    paragraph.add_run("\t")

    page = paragraph.add_run()
    page.font.name = "Arial"
    page.font.size = Pt(10)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = "PAGE"
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    page._r.extend([begin, instruction, separate, end])
    paragraph.add_run("\t")

    right = paragraph.add_run(right_text)
    right.font.name = "Arial"
    right.font.size = Pt(10)


def set_landscape_with_footer(section, left_text, right_text):
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)
    section.left_margin = Inches(LANDSCAPE_LEFT_MARGIN)
    section.right_margin = Inches(LANDSCAPE_RIGHT_MARGIN)
    section.top_margin = Inches(LANDSCAPE_TOP_MARGIN)
    section.bottom_margin = Inches(LANDSCAPE_BOTTOM_MARGIN)
    section.footer_distance = Inches(LANDSCAPE_BOTTOM_MARGIN)
    setup_footer(section, left_text, right_text)


def set_layout_from_template(section, template_path):
    template_section = Document(template_path).sections[0]
    section.orientation = template_section.orientation
    section.page_width = template_section.page_width
    section.page_height = template_section.page_height
    section.left_margin = template_section.left_margin
    section.right_margin = template_section.right_margin
    section.top_margin = template_section.top_margin
    section.bottom_margin = template_section.bottom_margin
    section.header_distance = template_section.header_distance
    section.footer_distance = template_section.footer_distance


def clear_footer_for_section(section):
    unlink_footer(section)
    clear_footer(section.footer.paragraphs[0])


def append_document_body(destination, source_path):
    source = Document(source_path)
    destination_body = destination._element.body
    section_properties = destination_body.sectPr
    if section_properties is not None:
        destination_body.remove(section_properties)

    for element in source._element.body.iterchildren():
        if element.tag != qn("w:sectPr"):
            destination_body.append(deepcopy(element))

    if section_properties is not None:
        destination_body.append(section_properties)


def center_cover_subtitle(doc):
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == COVER_SUBTITLE_TEXT:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            break


def add_toc_heading(doc, text):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(3)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt(TOC_HEADING_SIZE)
    run.font.name = "Arial"


def add_toc_entry(doc, text, page_number):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.left_indent = Inches(0.25)
    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
    )
    title = paragraph.add_run(text)
    title.font.size = Pt(TOC_ENTRY_SIZE)
    title.font.name = "Arial"
    paragraph.add_run("\t")
    page = paragraph.add_run(str(page_number))
    page.font.size = Pt(TOC_ENTRY_SIZE)
    page.font.name = "Arial"


def add_table_of_contents(doc, layouts):
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(10)
    title_run = title.add_run("Table of Contents")
    title_run.bold = True
    title_run.font.size = Pt(TOC_TITLE_SIZE)
    title_run.font.name = "Arial Rounded MT Bold"

    entries = [("heading", "Puzzles", 1)]
    entries.extend(("entry", layout["title"], index) for index, layout in enumerate(layouts, 1))
    answer_start = len(layouts) + 1
    entries.append(("heading", "Answer Keys", answer_start))
    entries.extend(
        (
            "entry",
            f'Answer Key: "{layout["title"]}"',
            answer_start + index // 4,
        )
        for index, layout in enumerate(layouts)
    )

    for index, (kind, text, page_number) in enumerate(entries):
        if index and index % TOC_ENTRIES_PER_PAGE == 0:
            doc.add_page_break()
        if kind == "heading":
            add_toc_heading(doc, text)
        else:
            add_toc_entry(doc, text, page_number)


def add_puzzle_pages(doc, layouts):
    for index, layout in enumerate(layouts):
        if index:
            doc.add_page_break()

        add_spacer(doc, 6)
        add_title(doc, layout["title"])
        add_spacer(doc, 8)

        grid_table = doc.add_table(rows=ROWS, cols=COLS)
        grid_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(grid_table)
        for row in grid_table.rows:
            row.height = Pt(GRID_ROW_HEIGHT)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            for cell in row.cells:
                cell.width = Pt(GRID_CELL_WIDTH)

        for row in range(ROWS):
            for column in range(COLS):
                cell = grid_table.cell(row, column)
                cell.text = layout["grid"][row][column]
                style_grid_cell(cell)

        add_spacer(doc, 10)
        word_table = doc.add_table(rows=2, cols=4)
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(word_table)
        for row in word_table.rows:
            row.height = Pt(WORD_BANK_ROW_HEIGHT)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            for cell in row.cells:
                cell.width = Pt(WORD_BANK_COL_WIDTH)

        words = [word for word, _, _ in layout["words"]]
        position = 0
        for row in range(2):
            for column in range(4):
                if position < len(words):
                    cell = word_table.cell(row, column)
                    cell.text = words[position]
                    style_word_cell(cell)
                position += 1


def add_answer_pages(doc, layouts):
    for start in range(0, len(layouts), 4):
        if start:
            doc.add_page_break()

        answer_table = doc.add_table(rows=2, cols=2)
        answer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(answer_table)
        for row in answer_table.rows:
            row.height = Pt(ANSWER_BLOCK_HEIGHT)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        block = layouts[start:start + 4]
        for slot, layout in enumerate(block):
            outer_cell = answer_table.cell(slot // 2, slot % 2)
            outer_cell.paragraphs[0].paragraph_format.space_after = Pt(8)
            title_paragraph = outer_cell.add_paragraph()
            title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_paragraph.paragraph_format.space_before = Pt(0)
            title_paragraph.paragraph_format.space_after = Pt(8)
            title_run = title_paragraph.add_run(f'Answer Key: "{layout["title"]}"')
            title_run.bold = True
            title_run.font.size = Pt(ANSWER_TITLE_SIZE)
            title_run.font.name = "Arial"

            mini = outer_cell.add_table(rows=ROWS, cols=COLS)
            mini.alignment = WD_TABLE_ALIGNMENT.CENTER
            remove_table_borders(mini)
            for row in mini.rows:
                row.height = Pt(ANSWER_CELL_SIZE)
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                for cell in row.cells:
                    cell.width = Pt(ANSWER_CELL_SIZE)

            for row in range(ROWS):
                for column in range(COLS):
                    cell = mini.cell(row, column)
                    paragraph = cell.paragraphs[0]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    run = paragraph.add_run(layout["grid"][row][column])
                    run.font.size = Pt(ANSWER_GRID_FONT_SIZE)
                    run.font.name = "Arial"
                    if (row, column) in layout["answer_positions"]:
                        run.font.color.rgb = RGBColor(255, 0, 0)


def build_booklet_doc(layouts):
    doc = Document(COVER_TEMPLATE_PATH)
    center_cover_subtitle(doc)

    doc.add_page_break()
    append_document_body(doc, COPYRIGHT_TEMPLATE_PATH)
    doc.add_page_break()
    append_document_body(doc, INSTRUCTIONS_TEMPLATE_PATH)
    doc.add_page_break()
    add_table_of_contents(doc, layouts)

    puzzle_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    set_landscape_with_footer(puzzle_section, "Book One Puzzles", "Puzzles by Riley")
    set_page_number_start(puzzle_section, 1)
    add_puzzle_pages(doc, layouts)

    answer_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    set_landscape_with_footer(answer_section, "Book One Answer Key", "Puzzles by Riley")
    continue_page_numbering(answer_section)
    add_answer_pages(doc, layouts)

    back_cover_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    set_layout_from_template(back_cover_section, BACK_COVER_TEMPLATE_PATH)
    continue_page_numbering(back_cover_section)
    clear_footer_for_section(back_cover_section)
    append_document_body(doc, BACK_COVER_TEMPLATE_PATH)
    return doc


def save_document_safely(doc, path):
    base, extension = os.path.splitext(path)
    counter = 1
    while True:
        try:
            doc.save(path)
            return path
        except PermissionError:
            path = f"{base}_{counter}{extension}"
            counter += 1


def get_available_output_path(path):
    base, extension = os.path.splitext(path)
    candidate = path
    counter = 1
    while os.path.exists(candidate):
        candidate = f"{base}_{counter}{extension}"
        counter += 1
    return candidate


def open_folder(path):
    try:
        os.startfile(path)
    except Exception:
        try:
            subprocess.Popen(["explorer", path])
        except Exception:
            pass


def main():
    try:
        print("Starting wordsearch generator...")
        print(f"Base folder: {BASE_FOLDER}")
        print(f"Excel file: {EXCEL_PATH}")
        print(f"Output folder: {OUTPUT_FOLDER}")
        print()

        verify_required_paths()
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        titles, puzzles = load_puzzles_from_excel(EXCEL_PATH)
        if not puzzles:
            print("No puzzles were found in the Excel file.")
            input("Press Enter to close...")
            return

        layouts = build_puzzle_layouts(titles, puzzles)
        booklet = build_booklet_doc(layouts)
        output_path = get_available_output_path(
            os.path.join(OUTPUT_FOLDER, "wordsearch_booklet.docx")
        )
        save_document_safely(booklet, output_path)

        print()
        print("Saved complete booklet to:")
        print(output_path)
        print()
        open_folder(OUTPUT_FOLDER)

    except Exception as error:
        print()
        print("An error occurred:")
        print(str(error))
        print()
        print("Also make sure required packages are installed:")
        print("pip install openpyxl python-docx")

    input("Press Enter to close...")


if __name__ == "__main__":
    main()
