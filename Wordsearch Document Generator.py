import random
import string
import os
import tempfile
import subprocess
from openpyxl import load_workbook  # type: ignore
from docx import Document  # type: ignore
from docx.shared import Pt, Inches, RGBColor  # type: ignore
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER  # type: ignore
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE  # type: ignore
from docx.enum.section import WD_ORIENT  # type: ignore
from docx.oxml import OxmlElement  # type: ignore
from docx.oxml.ns import qn  # type: ignore
from docxcompose.composer import Composer  # type: ignore

ROWS = 8
COLS = 12

desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
if not os.path.exists(desktop):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")

EXCEL_PATH = os.path.join(desktop, "wordsearch_booklet", "Book2.xlsx")

PUZZLES_FOLDER = os.path.join(desktop, "Puzzles")
ANSWER_KEY_FOLDER = os.path.join(desktop, "Answer Key")
OUTPUT_FOLDER = desktop

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
    ws = wb.worksheets[0]

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


def setup_footer(doc):
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


def setup_document(doc):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(LEFT_MARGIN)
    section.right_margin = Inches(RIGHT_MARGIN)
    section.top_margin = Inches(TOP_MARGIN)
    section.bottom_margin = Inches(BOTTOM_MARGIN)
    setup_footer(doc)


def generate_puzzles_and_answers(puzzle_titles, puzzles):
    puzzles_doc = Document()
    setup_document(puzzles_doc)

    answers_doc = Document()
    setup_document(answers_doc)

    answer_keys = []

    for i, puzzle in enumerate(puzzles, start=1):
        title = puzzle_titles[i - 1]

        grid, placed_words = build_grid(puzzle)
        answer_keys.append((title, grid, get_answer_positions(placed_words)))

        add_spacer(puzzles_doc, 2)
        add_title(puzzles_doc, title)
        add_spacer(puzzles_doc, 3)

        table = puzzles_doc.add_table(rows=ROWS, cols=COLS)
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

        add_spacer(puzzles_doc, 3)

        words_only = [word for word, _, _ in puzzle]
        word_bank = puzzles_doc.add_table(rows=2, cols=4)
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

        if i < len(puzzles):
            puzzles_doc.add_page_break()

    for start in range(0, len(answer_keys), 4):
        if start > 0:
            answers_doc.add_page_break()

        block = answer_keys[start:start + 4]

        answer_page = answers_doc.add_table(rows=2, cols=2)
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

    return puzzles_doc, answers_doc


def find_docx_file(folder, base_name):
    exact = os.path.join(folder, f"{base_name}.docx")
    if os.path.exists(exact):
        return exact

    for name in os.listdir(folder):
        if name.lower() == f"{base_name.lower()}.docx":
            return os.path.join(folder, name)

    raise FileNotFoundError(f"Could not find: {base_name}.docx in {folder}")


def merge_word_files_in_order(file_paths, output_path):
    master = Document(file_paths[0])
    composer = Composer(master)

    for path in file_paths[1:]:
        composer.append(Document(path))

    composer.save(output_path)


def save_document_safely(doc, file_path):
    base, ext = os.path.splitext(file_path)
    counter = 1

    while True:
        try:
            doc.save(file_path)
            return file_path
        except PermissionError:
            file_path = f"{base}_{counter}{ext}"
            counter += 1


def get_available_output_path(file_path):
    base, ext = os.path.splitext(file_path)
    candidate = file_path
    counter = 1

    while os.path.exists(candidate):
        candidate = f"{base}_{counter}{ext}"
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
        print("Starting wordsearch booklet generator...")
        print(f"Using desktop folder: {desktop}")
        print(f"Excel file: {EXCEL_PATH}")
        print(f"Puzzles folder: {PUZZLES_FOLDER}")
        print(f"Answer key folder: {ANSWER_KEY_FOLDER}")
        print()

        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        PUZZLE_TITLES, PUZZLES = load_puzzles_from_excel(EXCEL_PATH)

        if not PUZZLES:
            print("No puzzles were found in the Excel file.")
            input("Press Enter to close...")
            return

        generated_puzzles_doc, generated_answers_doc = generate_puzzles_and_answers(PUZZLE_TITLES, PUZZLES)

        temp_dir = tempfile.mkdtemp()

        generated_puzzles_path = save_document_safely(
            generated_puzzles_doc,
            os.path.join(temp_dir, "generated_puzzles.docx")
        )

        generated_answers_path = save_document_safely(
            generated_answers_doc,
            os.path.join(temp_dir, "generated_answer_keys.docx")
        )

        puzzles_merge_order = [
            find_docx_file(PUZZLES_FOLDER, "Puzzles Cover"),
            find_docx_file(PUZZLES_FOLDER, "Puzzles Copyright"),
            find_docx_file(PUZZLES_FOLDER, "Puzzles Instructions"),
            generated_puzzles_path,
            find_docx_file(PUZZLES_FOLDER, "Puzzles Back Cover"),
        ]

        answers_merge_order = [
            find_docx_file(ANSWER_KEY_FOLDER, "Answer Key Cover"),
            find_docx_file(ANSWER_KEY_FOLDER, "Answer Key Copyright"),
            generated_answers_path,
            find_docx_file(ANSWER_KEY_FOLDER, "Answer Key Back Cover"),
        ]

        final_puzzles_path = get_available_output_path(
            os.path.join(OUTPUT_FOLDER, "wordsearch_puzzles_booklet.docx")
        )

        final_answers_path = get_available_output_path(
            os.path.join(OUTPUT_FOLDER, "wordsearch_answer_key_booklet.docx")
        )

        merge_word_files_in_order(puzzles_merge_order, final_puzzles_path)
        merge_word_files_in_order(answers_merge_order, final_answers_path)

        print("Saved puzzles booklet to:")
        print(final_puzzles_path)
        print()
        print("Saved answer key booklet to:")
        print(final_answers_path)
        print()
        print("Opening output folder...")
        open_folder(OUTPUT_FOLDER)

    except Exception as e:
        print("An error occurred:")
        print(str(e))

    input("Press Enter to close...")


if __name__ == "__main__":
    main()