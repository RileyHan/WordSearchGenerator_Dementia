import random
import string
import os
import subprocess
from copy import deepcopy
from openpyxl import load_workbook  # type: ignore
from docx import Document  # type: ignore
from docx.shared import Pt, Inches, RGBColor  # type: ignore
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER  # type: ignore
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE  # type: ignore
from docx.enum.section import WD_ORIENT, WD_SECTION  # type: ignore
from docx.oxml import OxmlElement  # type: ignore
from docx.oxml.ns import qn  # type: ignore

# Puzzle size
ROWS = 8
COLS = 12

# Main folder locations
BASE_FOLDER = r"C:\Users\riley\OneDrive\Desktop\wordsearch_booklet"
EXCEL_PATH = os.path.join(BASE_FOLDER, "Book2.xlsx")
OUTPUT_FOLDER = BASE_FOLDER
SCRIPT_FOLDER = os.path.dirname(os.path.abspath(__file__))

# Front/back matter files
COVER_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Cover.docx")
COPYRIGHT_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Copyright.docx")
INSTRUCTIONS_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Instructions.docx")
BACK_COVER_PATH = os.path.join(SCRIPT_FOLDER, "Puzzles Back Cover.docx")

# Font sizes for puzzle pages
TITLE_FONT_SIZE = 26
GRID_FONT_SIZE = 36
WORD_FONT_SIZE = 18

# Puzzle page table sizing
GRID_CELL_WIDTH = 58
GRID_ROW_HEIGHT = 42
WORD_BANK_ROW_HEIGHT = 24
WORD_BANK_COL_WIDTH = 155

# Landscape page margins
LANDSCAPE_TOP_MARGIN = 0.25
LANDSCAPE_BOTTOM_MARGIN = 0.25
LANDSCAPE_LEFT_MARGIN = 0.25
LANDSCAPE_RIGHT_MARGIN = 0.25

# Answer key formatting
ANSWER_TITLE_SIZE = 16
ANSWER_GRID_FONT_SIZE = 14
ANSWER_CELL_SIZE = 24
ANSWER_BLOCK_HEIGHT = 255


def verify_required_paths():
    """
    Make sure the Excel file exists before trying to build anything.
    """
    required_files = [
        EXCEL_PATH,
        COVER_PATH,
        COPYRIGHT_PATH,
        INSTRUCTIONS_PATH,
        BACK_COVER_PATH,
    ]

    print("Checking required files...")
    for file_path in required_files:
        print(f"  {file_path}")
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"Missing file: {file_path}")


def load_puzzles_from_excel(file_path):
    """
    Load puzzle titles and words from the Excel workbook.

    Expected format:
    - Column A = puzzle title
    - Remaining columns = words for that puzzle

    Rules:
    - blank rows are skipped
    - blank cells are skipped
    - words longer than COLS are skipped
    - puzzles are limited to ROWS words
    """
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

            # Store word with placeholder row/column values
            words.append((word, 1, 1))

        if len(words) > ROWS:
            print(f"Row {excel_row_num} has more than {ROWS} words. Extra words will be ignored.")
            words = words[:ROWS]

        if words:
            puzzle_titles.append(title)
            puzzles.append(words)

    return puzzle_titles, puzzles


def build_grid(words):
    """
    Create a puzzle grid:
    - each word is placed horizontally on a random row
    - remaining blanks are filled with random letters

    Returns:
    - completed grid
    - placed word positions
    """
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

        # Store as 1-based row/column positions
        placed_words.append((word, r + 1, c + 1))

    # Fill remaining blank spots with random uppercase letters
    for r in range(ROWS):
        for c in range(COLS):
            if grid[r][c] == "":
                grid[r][c] = random.choice(string.ascii_uppercase)

    return grid, placed_words


def get_answer_positions(words):
    """
    Convert placed word data into a set of coordinates
    so answer-key letters can be colored red.
    """
    positions = set()
    for word, r, c in words:
        r -= 1
        c -= 1
        for i in range(len(word)):
            if 0 <= r < ROWS and 0 <= c + i < COLS:
                positions.add((r, c + i))
    return positions


def remove_table_borders(table):
    """
    Remove all visible borders from a table.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr

    borders = OxmlElement("w:tblBorders")
    for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        elem = OxmlElement(f"w:{edge}")
        elem.set(qn("w:val"), "nil")
        borders.append(elem)

    tblPr.append(borders)


def style_grid_cell(cell):
    """
    Apply formatting to puzzle grid cells.
    """
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.size = Pt(GRID_FONT_SIZE)
            run.font.name = "Arial"


def style_word_cell(cell):
    """
    Apply formatting to word-bank cells.
    """
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.size = Pt(WORD_FONT_SIZE)
            run.font.name = "Arial"


def add_title(doc, title_text):
    """
    Add a centered puzzle title.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(12)

    run = p.add_run(f'"{title_text}"')
    run.bold = True
    run.font.size = Pt(TITLE_FONT_SIZE)
    run.font.name = "Arial Rounded MT Bold"


def add_spacer(doc, points_after):
    """
    Add an empty paragraph used for vertical spacing.
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(points_after)


def remove_paragraph(paragraph):
    """
    Remove a paragraph from the document.
    """
    p = paragraph._element
    p.getparent().remove(p)


def set_portrait_layout(section):
    """
    Set a section to portrait layout with standard margins.
    """
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)


def clear_section_footer(section):
    """
    Remove inherited footer content from a section.
    """
    unlink_footer_from_previous(section)
    footer = section.footer
    for paragraph in footer.paragraphs:
        clear_footer(paragraph)


def append_document_contents(target_doc, source_path):
    """
    Append all body content from a Word document into the target document.
    """
    source_doc = Document(source_path)

    if len(target_doc.paragraphs) == 1 and target_doc.paragraphs[0].text == "" and len(target_doc.tables) == 0:
        remove_paragraph(target_doc.paragraphs[0])

    target_body = target_doc.element.body
    insert_at = len(target_body)
    if len(target_body) > 0 and target_body[-1].tag == qn("w:sectPr"):
        insert_at -= 1

    for element in source_doc.element.body:
        if element.tag == qn("w:sectPr"):
            continue
        target_body.insert(insert_at, deepcopy(element))
        insert_at += 1


def add_table_of_contents(doc, puzzle_titles):
    """
    Add a manual table of contents for the puzzle pages.

    Puzzle pages always begin in the next section and restart page
    numbering at 1, so the page references here stay stable even if
    the TOC itself grows to multiple pages.
    """
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(8)
    title_p.paragraph_format.space_after = Pt(12)

    title_run = title_p.add_run("Table of Contents")
    title_run.bold = True
    title_run.font.size = Pt(TITLE_FONT_SIZE)
    title_run.font.name = "Arial Rounded MT Bold"

    if not puzzle_titles:
        empty_p = doc.add_paragraph()
        empty_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        empty_p.paragraph_format.space_before = Pt(0)
        empty_p.paragraph_format.space_after = Pt(8)

        empty_run = empty_p.add_run("No puzzles available.")
        empty_run.font.size = Pt(16)
        empty_run.font.name = "Arial"
        return

    for page_number, title in enumerate(puzzle_titles, start=1):
        entry_p = doc.add_paragraph()
        entry_p.paragraph_format.space_before = Pt(0)
        entry_p.paragraph_format.space_after = Pt(6)
        entry_p.paragraph_format.tab_stops.add_tab_stop(
            Inches(6.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )

        title_run = entry_p.add_run(title)
        title_run.font.size = Pt(16)
        title_run.font.name = "Arial"

        page_run = entry_p.add_run(f"\t{page_number}")
        page_run.font.size = Pt(16)
        page_run.font.name = "Arial"


def prepare_puzzle_data(puzzle_titles, puzzles):
    """
    Build puzzle grids once so puzzle pages and answer keys stay aligned.
    """
    prepared = []
    for i, puzzle in enumerate(puzzles):
        title = puzzle_titles[i]
        grid, placed_words = build_grid(puzzle)
        prepared.append(
            {
                "title": title,
                "grid": grid,
                "placed_words": placed_words,
                "positions": get_answer_positions(placed_words),
            }
        )
    return prepared


def clear_footer(paragraph):
    """
    Remove everything currently inside a footer paragraph.
    """
    p = paragraph._element
    for child in list(p):
        p.remove(child)


def unlink_footer_from_previous(section):
    """
    Ensure this section has its own footer instead of inheriting
    from a previous section.
    """
    section.footer.is_linked_to_previous = False


def ensure_pgNumType(sectPr):
    """
    Make sure the section has a page-number settings element.
    """
    pg_num_type = sectPr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        pg_num_type = OxmlElement("w:pgNumType")
        sectPr.append(pg_num_type)
    return pg_num_type


def set_page_number_start(section, start_value):
    """
    Set the starting page number for a section.
    """
    sectPr = section._sectPr
    pg_num_type = ensure_pgNumType(sectPr)
    pg_num_type.set(qn("w:start"), str(start_value))


def setup_footer_for_section(section, left_text, right_text):
    """
    Create a footer with:
    - left text
    - centered page number
    - right text
    """
    unlink_footer_from_previous(section)
    footer = section.footer
    paragraph = footer.paragraphs[0]

    clear_footer(paragraph)

    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)

    # Extra spacing from left and right edges
    paragraph.paragraph_format.left_indent = Inches(0.35)
    paragraph.paragraph_format.right_indent = Inches(0.35)

    tab_stops = paragraph.paragraph_format.tab_stops

    # Center page number
    tab_stops.add_tab_stop(Inches(5.35), WD_TAB_ALIGNMENT.CENTER, WD_TAB_LEADER.SPACES)

    # Pull right-side footer text inward a bit
    tab_stops.add_tab_stop(Inches(9.95), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)

    # Left footer text
    left_run = paragraph.add_run(left_text)
    left_run.font.name = "Arial"
    left_run.font.size = Pt(10)

    paragraph.add_run("\t")

    # Page number field
    page_run = paragraph.add_run()
    page_run.font.name = "Arial"
    page_run.font.size = Pt(10)

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"

    fld_char_sep = OxmlElement("w:fldChar")
    fld_char_sep.set(qn("w:fldCharType"), "separate")

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")

    page_run._r.append(fld_char_begin)
    page_run._r.append(instr_text)
    page_run._r.append(fld_char_sep)
    page_run._r.append(fld_char_end)

    paragraph.add_run("\t")

    # Right footer text
    right_run = paragraph.add_run(right_text)
    right_run.font.name = "Arial"
    right_run.font.size = Pt(10)


def set_landscape_with_footer(section, left_text, right_text):
    """
    Set a section to landscape layout and apply footer formatting.
    """
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)
    section.left_margin = Inches(LANDSCAPE_LEFT_MARGIN)
    section.right_margin = Inches(LANDSCAPE_RIGHT_MARGIN)
    section.top_margin = Inches(LANDSCAPE_TOP_MARGIN)
    section.bottom_margin = Inches(LANDSCAPE_BOTTOM_MARGIN)
    setup_footer_for_section(section, left_text, right_text)


def add_puzzles_content(doc, prepared_puzzles):
    """
    Add puzzle pages to the current document.
    """
    for i, puzzle_data in enumerate(prepared_puzzles, start=1):
        # Start each puzzle on a new page except the first one
        if i > 1:
            doc.add_page_break()

        title = puzzle_data["title"]
        grid = puzzle_data["grid"]
        placed_words = puzzle_data["placed_words"]

        add_spacer(doc, 6)
        add_title(doc, title)
        add_spacer(doc, 8)

        # Puzzle letter grid
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

        add_spacer(doc, 10)

        # Word bank
        words_only = [word for word, _, _ in placed_words]
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

def add_answers_content(doc, prepared_puzzles):
    """
    Add answer key pages to the current document.
    Each answer page contains up to 4 answer grids.
    """
    first_page = True
    for start in range(0, len(prepared_puzzles), 4):
        if not first_page:
            doc.add_page_break()
        first_page = False

        block = prepared_puzzles[start:start + 4]

        add_spacer(doc, 4)

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
                    puzzle_data = block[slot]
                    title = puzzle_data["title"]
                    grid = puzzle_data["grid"]
                    positions = puzzle_data["positions"]

                    top_space = outer_cell.paragraphs[0]
                    top_space.paragraph_format.space_before = Pt(0)
                    top_space.paragraph_format.space_after = Pt(8)

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

                            # Highlight answer letters in red
                            if (r, c) in positions:
                                run.font.color.rgb = RGBColor(255, 0, 0)

                slot += 1


def build_combined_book_doc(puzzle_titles, puzzles):
    """
    Build the full book in this order:
    cover, copyright, instruction, table of contents,
    puzzles, answer keys, back cover.
    """
    prepared_puzzles = prepare_puzzle_data(puzzle_titles, puzzles)

    doc = Document()
    set_portrait_layout(doc.sections[0])
    clear_section_footer(doc.sections[0])

    append_document_contents(doc, COVER_PATH)
    doc.add_page_break()
    append_document_contents(doc, COPYRIGHT_PATH)
    doc.add_page_break()
    append_document_contents(doc, INSTRUCTIONS_PATH)
    doc.add_page_break()
    add_table_of_contents(doc, [puzzle_data["title"] for puzzle_data in prepared_puzzles])

    if prepared_puzzles:
        puzzle_section = doc.add_section(WD_SECTION.NEW_PAGE)
        set_landscape_with_footer(puzzle_section, "Book One Puzzles", "Puzzles by Riley")
        set_page_number_start(puzzle_section, 1)
        add_puzzles_content(doc, prepared_puzzles)

        answer_section = doc.add_section(WD_SECTION.NEW_PAGE)
        set_landscape_with_footer(answer_section, "Book One Answer Key", "Puzzles by Riley")
        set_page_number_start(answer_section, 1)
        add_answers_content(doc, prepared_puzzles)

    back_cover_section = doc.add_section(WD_SECTION.NEW_PAGE)
    set_portrait_layout(back_cover_section)
    clear_section_footer(back_cover_section)
    append_document_contents(doc, BACK_COVER_PATH)

    return doc


def save_document_safely(doc, file_path):
    """
    Save the document. If the file is open/locked,
    save with a numbered suffix instead.
    """
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
    """
    Find a filename that doesn't already exist.
    """
    base, ext = os.path.splitext(file_path)
    candidate = file_path
    counter = 1

    while os.path.exists(candidate):
        candidate = f"{base}_{counter}{ext}"
        counter += 1

    return candidate


def open_folder(path):
    """
    Open the output folder in Windows Explorer.
    """
    try:
        os.startfile(path)
    except Exception:
        try:
            subprocess.Popen(["explorer", path])
        except Exception:
            pass


def main():
    """
    Main program flow:
    - verify Excel exists
    - load puzzle data
    - build the combined book
    - save it
    - open output folder
    """
    try:
        print("Starting wordsearch generator...")
        print(f"Base folder: {BASE_FOLDER}")
        print(f"Excel file: {EXCEL_PATH}")
        print(f"Output folder: {OUTPUT_FOLDER}")
        print()

        verify_required_paths()
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        puzzle_titles, puzzles = load_puzzles_from_excel(EXCEL_PATH)

        if not puzzles:
            print("No puzzles were found in the Excel file. Creating book with front and back matter only.")

        combined_doc = build_combined_book_doc(puzzle_titles, puzzles)

        final_book_path = get_available_output_path(
            os.path.join(OUTPUT_FOLDER, "wordsearch_book_complete.docx")
        )

        save_document_safely(combined_doc, final_book_path)

        print()
        print("Saved combined book to:")
        print(final_book_path)
        print()
        print("Opening output folder...")
        open_folder(OUTPUT_FOLDER)

    except Exception as e:
        print()
        print("An error occurred:")
        print(str(e))
        print()
        print("Expected file:")
        print(EXCEL_PATH)
        print()
        print("Also make sure required packages are installed:")
        print("pip install openpyxl python-docx")

    input("Press Enter to close...")


if __name__ == "__main__":
    main()