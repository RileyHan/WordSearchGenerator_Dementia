import random
import string
import os
from docx import Document # type: ignore
from docx.shared import Pt # type: ignore
from docx.enum.text import WD_ALIGN_PARAGRAPH # type: ignore
from docx.oxml import OxmlElement # type: ignore
from docx.oxml.ns import qn # type: ignore

ROWS = 8
COLS = 12


# -----------------------------
# YOUR PUZZLES GO HERE
# -----------------------------
PUZZLES = [
    # 1 Flowers
[
("ROSE", 1, 2),
("LILLY", 2, 5),
("TULIP", 3, 1),
("DAISY", 4, 4),
("BLOOM", 5, 2),
("SMELL", 6, 6),
("BUD", 7, 3),
("PETAL", 8, 5),
],

# 2 Garden Tools
[
("SHOVEL", 1, 2),
("RAKE", 2, 6),
("HOE", 3, 1),
("GLOVES", 4, 3),
("PAIL", 5, 7),
("SHEARS", 6, 2),
("STAKE", 7, 4),
("SPADE", 8, 6),
],

# 3 Vegetables
[
("CARROT", 1, 1),
("ONION", 2, 3),
("POTATO", 3, 5),
("CORN", 4, 2),
("BEANS", 5, 7),
("PEAS", 6, 4),
("TOMATO", 7, 1),
("GARLIC", 8, 5),
],

# 4 Fruits
[
("APPLE", 1, 2),
("PEAR", 2, 6),
("PLUM", 3, 1),
("PEACH", 4, 3),
("LEMON", 5, 5),
("BERRY", 6, 2),
("MELON", 7, 4),
("CHERRY", 8, 1),
],

# 5 Seeds
[
("SEED", 1, 3),
("PLANT", 2, 1),
("SOW", 3, 5),
("GROW", 4, 2),
("SPROUT", 5, 4),
("SMALL", 6, 1),
("SOIL", 7, 6),
("BURY", 8, 2),
],

# 6 Plants
[
("PLANT", 1, 2),
("WATER", 2, 5),
("STEM", 3, 1),
("ROOT", 4, 3),
("GROW", 5, 6),
("SOIL", 6, 2),
("LIGHT", 7, 4),
("CARE", 8, 1),
],

# 7 Trees
[
("TREE", 1, 3),
("TRUNK", 2, 1),
("BRANCH", 3, 2),
("WOOD", 4, 6),
("ROOT", 5, 4),
("BARK", 6, 2),
("SHADE", 7, 5),
("TALL", 8, 1),
],

# 8 Leaves
[
("LEAF", 1, 2),
("GREEN", 2, 3),
("BROWN", 3, 1),
("FALL", 4, 5),
("CRISP", 5, 2),
("FRESH", 6, 4),
("PLANT", 7, 6),
("PILE", 8, 1),
],

# 9 Soil
[
("SOIL", 1, 1),
("TILL", 2, 5),
("MUD", 3, 8),
("EARTH", 4, 2),
("GROUND", 5, 3),
("RICH", 6, 7),
("PLANT", 7, 4),
("DARK", 8, 6),
],

# 10 Watering
[
("WATER", 1, 2),
("HOSE", 2, 6),
("CAN", 3, 1),
("SPRAY", 4, 3),
("POUR", 5, 5),
("SOAK", 6, 7),
("GROW", 7, 2),
("DROP", 8, 4),
],

# 11 Sunshine
[
("SUN", 1, 3),
("WARM", 2, 6),
("LIGHT", 3, 2),
("BRIGHT", 4, 1),
("SKY", 5, 5),
("SUMMER", 6, 2),
("HEAT", 7, 4),
("GLOW", 8, 6),
],

# 12 Rain
[
("RAIN", 1, 2),
("CLOUD", 2, 3),
("WIND", 3, 1),
("STORM", 4, 5),
("DRIP", 5, 8),
("DROP", 6, 4),
("COOL", 7, 6),
("BREEZE", 8, 1),
],

# 13 Creatures
[
("BUNNY", 1, 2),
("CAT", 2, 6),
("DOG", 3, 1),
("MOUSE", 4, 3),
("FROG", 5, 5),
("BIRD", 6, 2),
("SNAIL", 7, 4),
("DEER", 8, 1),
],

# 14 Birds
[
("ROBIN", 1, 3),
("HAWK", 2, 1),
("CROW", 3, 2),
("DOVE", 4, 5),
("FINCH", 5, 4),
("JAY", 6, 7),
("BIRD", 7, 2),
("SONG", 8, 6),
],

# 15 Insects
[
("BEE", 1, 2),
("ANT", 2, 3),
("BUG", 3, 1),
("FLY", 4, 5),
("MOTH", 5, 2),
("WASP", 6, 6),
("BEETLE", 7, 1),
("STING", 8, 4),
],

# 16 Butterflies
[
("SPRING", 1, 2),
("WING", 2, 4),
("COLOR", 3, 1),
("FLY", 4, 5),
("BEAUTY", 5, 2),
("FLOWER", 6, 3),
("LIGHT", 7, 6),
("SOFT", 8, 1),
],

# 17 Spring
[
("SPRING", 1, 3),
("BLOOM", 2, 6),
("GREEN", 3, 1),
("RAIN", 4, 5),
("BUD", 5, 8),
("GROW", 6, 2),
("WARM", 7, 4),
("PLANT", 8, 7),
],

# 18 Summer
[
("SUMMER", 1, 2),
("SUN", 2, 5),
("HEAT", 3, 1),
("GROW", 4, 6),
("WARM", 5, 3),
("WATER", 6, 4),
("GREEN", 7, 2),
("BRIGHT", 8, 1),
],

# 19 Autumn
[
("FALL", 1, 3),
("LEAF", 2, 1),
("BROWN", 3, 2),
("COOL", 4, 5),
("BARE", 5, 4),
("WIND", 6, 6),
("AUTUMN", 7, 2),
("DROP", 8, 1),
],

# 20 Colors
[
("RED", 1, 2),
("BLUE", 2, 3),
("YELLOW", 3, 1),
("GREEN", 4, 5),
("PURPLE", 5, 2),
("WHITE", 6, 4),
("ORANGE", 7, 1),
("PINK", 8, 6),
],

]


# -----------------------------
# GRID BUILDER
# -----------------------------
def build_grid(words):
    grid = [["" for _ in range(COLS)] for _ in range(ROWS)]

    for word, r, c in words:
        r -= 1
        c -= 1

        for i, ch in enumerate(word):
            grid[r][c + i] = ch

    for r in range(ROWS):
        for c in range(COLS):
            if grid[r][c] == "":
                grid[r][c] = random.choice(string.ascii_uppercase)

    return grid


# -----------------------------
# STYLE CELL (CENTER + FONT)
# -----------------------------
def style_cell(cell):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.runs
        if run:
            for r in run:
                r.font.size = Pt(12)


# -----------------------------
# MAKE CELL SQUARE + BORDER
# -----------------------------
def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = OxmlElement('w:tcBorders')

    for edge in ['top', 'left', 'bottom', 'right']:
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), 'single')
        tag.set(qn('w:sz'), '8')   # thickness
        tag.set(qn('w:space'), '0')
        tag.set(qn('w:color'), '000000')
        tcBorders.append(tag)

    tcPr.append(tcBorders)


# -----------------------------
# BUILD WORD DOC
# -----------------------------
desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
folder = os.path.join(desktop, "wordsearch_booklet")
os.makedirs(folder, exist_ok=True)

doc = Document()

for i, puzzle in enumerate(PUZZLES, start=1):

    doc.add_heading(f"Puzzle {i}", level=1)

    grid = build_grid(puzzle)

    table = doc.add_table(rows=ROWS, cols=COLS)
    table.style = 'Table Grid'

    # make cells uniform-ish
    for row in table.rows:
        for cell in row.cells:
            cell.width = Pt(25)

    for r in range(ROWS):
        for c in range(COLS):
            cell = table.cell(r, c)
            cell.text = grid[r][c]
            style_cell(cell)
            set_cell_border(cell)

    doc.add_page_break()


file_path = os.path.join(folder, "wordsearch_booklet.docx")
doc.save(file_path)

print("Saved booklet to:")
print(file_path)