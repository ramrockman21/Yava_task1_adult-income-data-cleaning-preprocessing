"""Generate the Week 1 internship report as an editable DOCX file."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "outputs"
REPORT_PATH = BASE_DIR / "Week1_Adult_Data_Preprocessing_Report.docx"
CONSOLE_IMAGE = OUTPUT_DIR / "successful_run.png"

BLUE = "2783DE"
DARK = "2C2C2B"
GRAY = "7D7A75"
LIGHT = "F0EFED"
SOFT_BLUE = "E5F2FC"
SOFT_GREEN = "E8F1EC"
WHITE = "FFFFFF"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, bold: bool = False, color: str = DARK) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_table(doc: Document, headers: list[str], rows: list[list[object]], widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for index, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[index], header, bold=True, color=WHITE)
        set_cell_shading(table.rows[0].cells[index], BLUE)
    header_pr = table.rows[0]._tr.get_or_add_trPr()
    repeat_header = OxmlElement("w:tblHeader")
    repeat_header.set(qn("w:val"), "true")
    header_pr.append(repeat_header)
    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        row_pr = table.rows[-1]._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        row_pr.append(cant_split)
        for column_index, value in enumerate(row):
            set_cell_text(cells[column_index], value)
            if row_index % 2 == 1:
                set_cell_shading(cells[column_index], "F9F8F7")
    if widths:
        for row in table.rows:
            for index, width in enumerate(widths):
                row.cells[index].width = Inches(width)
    doc.add_paragraph()
    return table


def add_code_block(doc: Document, code: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(8)
    paragraph.paragraph_format.left_indent = Inches(0.18)
    paragraph.paragraph_format.right_indent = Inches(0.18)
    set_cell_like_shading(paragraph, "F0EFED")
    run = paragraph.add_run(code.strip())
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor.from_string(DARK)


def set_cell_like_shading(paragraph, fill: str) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    p_pr.append(shd)


def add_callout(doc: Document, title: str, body: str, color: str = SOFT_BLUE) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    row_pr = table.rows[0]._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    row_pr.append(cant_split)
    set_cell_shading(cell, color)
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(4)
    title_run = paragraph.add_run(title + "\n")
    title_run.bold = True
    title_run.font.name = "Arial"
    title_run.font.size = Pt(10.5)
    title_run.font.color.rgb = RGBColor.from_string(DARK)
    body_run = paragraph.add_run(body)
    body_run.font.name = "Arial"
    body_run.font.size = Pt(10)
    body_run.font.color.rgb = RGBColor.from_string(DARK)
    doc.add_paragraph()


def add_bullets(doc: Document, items: list[str]) -> None:
    for item in items:
        paragraph = doc.add_paragraph(style="List Bullet")
        paragraph.add_run(item)


def add_numbered(doc: Document, items: list[str]) -> None:
    for number, item in enumerate(items, start=1):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.left_indent = Inches(0.24)
        paragraph.paragraph_format.first_line_indent = Inches(-0.24)
        paragraph.add_run(f"{number}.  {item}")


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Page ")
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string(GRAY)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def create_console_image(text_path: Path, image_path: Path) -> None:
    text = text_path.read_text(encoding="utf-8")
    font_path = "/usr/share/fonts/msttcore/cour.ttf"
    font = ImageFont.truetype(font_path, 22)
    lines = text.splitlines()
    width = 1500
    line_height = 31
    height = 80 + len(lines) * line_height
    image = Image.new("RGB", (width, height), "#202124")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((22, 18, width - 22, height - 18), 12, fill="#111315")
    y = 42
    for line in lines:
        draw.text((52, y), line, font=font, fill="#E8EAED")
        y += line_height
    image.save(image_path)


def configure_styles(doc: Document) -> None:
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(DARK)
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing = 1.25

    for name, size, color in [
        ("Title", 28, DARK),
        ("Heading 1", 19, DARK),
        ("Heading 2", 14, BLUE),
        ("Heading 3", 11.5, DARK),
    ]:
        style = styles[name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)


def main() -> None:
    metadata = json.loads((OUTPUT_DIR / "cleaning_metadata.json").read_text())
    outliers = pd.read_csv(OUTPUT_DIR / "outlier_summary.csv")
    create_console_image(OUTPUT_DIR / "run_output.txt", CONSOLE_IMAGE)

    doc = Document()
    configure_styles(doc)
    section = doc.sections[0]
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.68)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)

    # Header and footer
    header = section.header.paragraphs[0]
    header.text = "WEEK 1  /  DATA ACQUISITION, CLEANING & PREPROCESSING"
    header.runs[0].font.name = "Arial"
    header.runs[0].font.size = Pt(8.5)
    header.runs[0].font.bold = True
    header.runs[0].font.color.rgb = RGBColor.from_string(BLUE)
    add_page_number(section.footer.paragraphs[0])

    # Cover
    doc.add_paragraph("INTERNSHIP PROJECT", style="Subtitle")
    title = doc.add_paragraph(style="Title")
    title.add_run("Adult Census Income\n")
    accent = title.add_run("Data Cleaning & Preprocessing")
    accent.font.color.rgb = RGBColor.from_string(BLUE)
    subtitle = doc.add_paragraph()
    subtitle.paragraph_format.space_before = Pt(16)
    subtitle.paragraph_format.space_after = Pt(22)
    subtitle_run = subtitle.add_run(
        "A reproducible Python workflow for acquiring, auditing, cleaning, "
        "validating, and preparing public census data for analysis."
    )
    subtitle_run.font.size = Pt(14)
    subtitle_run.font.color.rgb = RGBColor.from_string(GRAY)

    add_table(
        doc,
        ["Project detail", "Value"],
        [
            ["Student", "Ramanan S"],
            ["Task", "Week 1: Data Acquisition, Cleaning, and Preprocessing"],
            ["Dataset", "Adult / Census Income"],
            ["Tools", "Python, pandas, NumPy, VS Code"],
            ["Report date", "28 September 2026"],
        ],
        widths=[1.7, 4.8],
    )
    add_callout(
        doc,
        "Outcome",
        "The workflow loaded 32,561 records, exposed 4,262 hidden missing-value "
        "cells, removed 24 exact duplicates, retained 32,537 valid records, and "
        "produced both a readable cleaned dataset and a 93-column model-ready dataset.",
        SOFT_GREEN,
    )
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # Executive summary
    doc.add_heading("Executive summary", level=1)
    doc.add_paragraph(
        "This report documents the complete preparation of the Adult Census Income "
        "dataset for subsequent exploratory analysis or machine-learning work. The "
        "raw CSV contains demographic and employment attributes for 32,561 records "
        "and a binary income label. Although pandas initially reported no nulls, "
        "three categorical fields used the character “?” as a hidden missing-value "
        "placeholder."
    )
    add_bullets(
        doc,
        [
            "Text fields were trimmed and “?” placeholders were converted to true missing values.",
            "Missing categories were labelled “Unknown” instead of deleting 2,399 affected rows.",
            "Twenty-four exact duplicates were removed; all numeric domain checks passed.",
            "Education and education_num were internally consistent for all categories.",
            "Outliers were flagged rather than automatically deleted; selected variables were capped at the 1st/99th percentiles and skewed financial variables were log-transformed.",
            "A separate model-ready table was created using robust scaling, one-hot encoding, and a binary target.",
        ],
    )
    add_callout(
        doc,
        "Central judgement",
        "Statistical rarity is not the same as data error. Extreme ages, work hours, "
        "capital gains, and losses can be genuine. Therefore the workflow preserves "
        "the original values in the cleaned file, adds transparent flags, and creates "
        "robust alternative features for modelling.",
    )

    doc.add_heading("1. Objectives and scope", level=1)
    doc.add_paragraph(
        "The objective was to demonstrate a traceable process for public-data "
        "acquisition, initial exploration, correction of quality issues, treatment "
        "of missing values and outliers, feature preparation, and documentation."
    )
    add_numbered(
        doc,
        [
            "Acquire and describe a reliable public dataset.",
            "Profile its structure, types, distributions, duplicates, and hidden quality problems.",
            "Apply defensible cleaning rules without silently discarding valid observations.",
            "Create analysis-ready and model-ready outputs.",
            "Explain the rationale, trade-offs, challenges, and likely analytical impact of every major decision.",
        ],
    )
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # Acquisition and dictionary
    doc.add_heading("2. Data acquisition and provenance", level=1)
    doc.add_paragraph(
        "The supplied file, adultfin.csv, matches the Adult dataset published by the "
        "UCI Machine Learning Repository. The dataset is also known as Census Income "
        "and is commonly used to study whether annual income exceeds $50,000. The "
        "repository attributes the dataset to Barry Becker and Ronny Kohavi and "
        "provides DOI 10.24432/C5XW20."
    )
    add_table(
        doc,
        ["Item", "Description"],
        [
            ["Local source", "adultfin.csv supplied with the internship task"],
            ["Public source", "UCI Machine Learning Repository — Adult dataset"],
            ["Repository URL", "https://archive.ics.uci.edu/dataset/2/adult"],
            ["Citation", "Becker, B. & Kohavi, R. (1996). Adult [Dataset]."],
            ["Unit of observation", "One census record / person"],
            ["Raw dimensions", "32,561 rows × 15 columns"],
            ["Target", "income: <=50K or >50K"],
        ],
    )
    doc.add_heading("2.1 Reproducible local acquisition", level=2)
    doc.add_paragraph(
        "The program reads the supplied CSV from the same folder as the script. This "
        "avoids dependence on internet access during assessment and ensures that the "
        "exact submitted dataset is analysed."
    )
    add_code_block(
        doc,
        """
BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "adultfin.csv"
raw = pd.read_csv(INPUT_FILE)
raw.columns = snake_case_columns(raw.columns.tolist())
        """,
    )

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_heading("3. Dataset structure", level=1)
    dictionary_rows = [
        ["age", "Integer", "Age in years"],
        ["workclass", "Category", "Employment class"],
        ["fnlwgt", "Integer", "Census final sampling weight"],
        ["education", "Category", "Highest education label"],
        ["education_num", "Integer", "Ordinal education code (1–16)"],
        ["marital_status", "Category", "Marital status"],
        ["occupation", "Category", "Occupation group"],
        ["relationship", "Category", "Household relationship"],
        ["race", "Category", "Race category"],
        ["sex", "Category", "Recorded sex category"],
        ["capital_gain", "Integer", "Capital gain amount"],
        ["capital_loss", "Integer", "Capital loss amount"],
        ["hours_per_week", "Integer", "Usual weekly work hours"],
        ["native_country", "Category", "Country of origin"],
        ["income", "Category", "Income class / target"],
    ]
    add_table(doc, ["Column", "Type", "Meaning"], dictionary_rows)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # Exploration
    doc.add_heading("4. Initial data exploration", level=1)
    doc.add_paragraph(
        "Profiling was performed before altering the data. The review included shape, "
        "data types, unique values, frequency distributions, numeric summaries, null "
        "counts, placeholder counts, duplicate detection, and range validation."
    )
    add_code_block(
        doc,
        """
print(raw.shape)
print(raw.info())
print(raw.describe(include="all").T)
print(raw.duplicated().sum())

for column in categorical_columns:
    values = raw[column].astype("string").str.strip()
    print(column, (values == "?").sum(), values.value_counts().head())
        """,
    )
    add_table(
        doc,
        ["Finding", "Result", "Interpretation"],
        [
            ["Shape", "32,561 × 15", "Expected UCI training-file dimensions"],
            ["Exact duplicates", "24 rows", "Safe to remove because every field is identical"],
            ["Native pandas nulls", "0", "Misleading until placeholders are checked"],
            ["“?” placeholder cells", "4,262", "Hidden missingness in three categorical fields"],
            ["Rows containing any “?”", "2,399 (7.37%)", "Dropping them would reduce coverage"],
            ["High-income class", "7,841 (24.08%)", "Target is imbalanced, though not extremely"],
            ["Numeric range failures", "0", "All values passed stated domain rules"],
        ],
    )
    doc.add_heading("4.1 Hidden missing values", level=2)
    add_table(
        doc,
        ["Column", "Placeholder count", "Percent of raw rows"],
        [
            ["workclass", "1,836", "5.64%"],
            ["occupation", "1,843", "5.66%"],
            ["native_country", "583", "1.79%"],
            ["Total cells", "4,262", "—"],
        ],
    )
    add_callout(
        doc,
        "Why this matters",
        "A simple df.isna().sum() would incorrectly conclude that the file is complete. "
        "The placeholder must first be standardized to pandas.NA.",
    )

    # Cleaning
    doc.add_heading("5. Cleaning procedures", level=1)
    doc.add_heading("5.1 Standardizing names and text", level=2)
    doc.add_paragraph(
        "Column labels were standardized to snake_case. Leading and trailing spaces "
        "were removed from every categorical value to prevent duplicate-looking "
        "categories such as “ Private” and “Private”. Empty strings and “?” were "
        "converted to explicit missing values."
    )
    add_code_block(
        doc,
        """
for column in CATEGORICAL_COLUMNS:
    cleaned[column] = cleaned[column].astype("string").str.strip()

cleaned[CATEGORICAL_COLUMNS] = cleaned[CATEGORICAL_COLUMNS].replace(
    {"?": pd.NA, "": pd.NA}
)
        """,
    )
    doc.add_heading("5.2 Removing exact duplicates", level=2)
    doc.add_paragraph(
        "The 24 exact duplicate rows were removed with drop_duplicates(). Because no "
        "single unique person identifier is present, only full-row duplicates were "
        "removed; records that merely share demographic values were retained."
    )
    add_code_block(
        doc,
        """
duplicate_rows = int(cleaned.duplicated().sum())
cleaned = cleaned.drop_duplicates().reset_index(drop=True)
        """,
    )
    doc.add_heading("5.3 Missing-value treatment", level=2)
    doc.add_paragraph(
        "The missing workclass, occupation, and native_country entries were assigned "
        "the explicit category “Unknown”. Mode imputation was rejected because it "
        "would manufacture common categories and understate uncertainty. Row deletion "
        "was also rejected because 2,399 rows have at least one placeholder and their "
        "missingness may not be random."
    )
    add_code_block(
        doc,
        """
for column in ["workclass", "occupation", "native_country"]:
    cleaned[column] = cleaned[column].fillna("Unknown")
        """,
    )
    doc.add_heading("5.4 Erroneous values and consistency checks", level=2)
    add_table(
        doc,
        ["Field", "Validation rule", "Failures"],
        [
            ["age", "17 to 100 years", "0"],
            ["fnlwgt", "Greater than 0", "0"],
            ["education_num", "1 to 16", "0"],
            ["capital_gain", "Non-negative", "0"],
            ["capital_loss", "Non-negative", "0"],
            ["hours_per_week", "1 to 99", "0"],
            ["income", "<=50K or >50K", "0"],
            ["education mapping", "One education_num per education label", "0 categories inconsistent"],
        ],
    )
    # Outliers
    doc.add_heading("6. Outlier identification and treatment", level=1)
    doc.add_paragraph(
        "Tukey’s 1.5 × IQR rule was used as an audit flag, not as an automatic deletion "
        "rule. This distinction is important for the Adult data: capital gains and "
        "losses are zero-inflated, and weekly hours are concentrated around 40. Their "
        "IQR fences therefore label many legitimate non-zero or non-standard values."
    )
    rows = []
    for _, row in outliers.iterrows():
        rows.append(
            [
                row["column"],
                f'{row["lower_fence"]:,.2f}',
                f'{row["upper_fence"]:,.2f}',
                f'{int(row["outlier_count"]):,}',
                f'{row["outlier_percent"]:.2f}%',
            ]
        )
    add_table(
        doc,
        ["Variable", "Lower fence", "Upper fence", "Flagged", "Percent"],
        rows,
    )
    add_code_block(
        doc,
        """
q1 = cleaned[column].quantile(0.25)
q3 = cleaned[column].quantile(0.75)
iqr = q3 - q1
lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
cleaned[f"{column}_outlier_iqr"] = (
    (cleaned[column] < lower) | (cleaned[column] > upper)
).astype("int8")
        """,
    )
    doc.add_heading("6.1 Treatment strategy", level=2)
    add_bullets(
        doc,
        [
            "No plausible record was deleted only because it was statistically unusual.",
            "Original values remain available in adult_cleaned.csv for auditability.",
            "Binary IQR flags support sensitivity checks and subgroup analysis.",
            "Age, fnlwgt, and hours_per_week received 1st/99th percentile capped alternatives.",
            "fnlwgt, capital_gain, and capital_loss received log1p transformations to reduce right skew while preserving zeros.",
            "Robust median/IQR scaling was used in the model-ready output; zero-IQR features fall back to a divisor of 1 after log transformation.",
        ],
    )
    add_table(
        doc,
        ["Variable", "1st percentile", "99th percentile", "Model treatment"],
        [
            ["age", "17", "74", "Cap alternative + robust scale"],
            ["fnlwgt", "27,185.08", "510,072", "Cap, log1p, robust scale"],
            ["hours_per_week", "8", "80", "Cap alternative + robust scale"],
            ["capital_gain", "0", "15,024 (raw profile)", "Keep raw + log1p"],
            ["capital_loss", "0", "1,980 (raw profile)", "Keep raw + log1p"],
        ],
    )

    # Preprocessing
    doc.add_heading("7. Preprocessing for further analysis", level=1)
    doc.add_heading("7.1 Target encoding", level=2)
    doc.add_paragraph(
        "The text target was mapped to income_high, where <=50K = 0 and >50K = 1. "
        "After duplicate removal, the cleaned target contains 24,698 lower-income and "
        "7,839 higher-income records."
    )
    add_code_block(
        doc,
        'cleaned["income_high"] = cleaned["income"].map({"<=50K": 0, ">50K": 1})',
    )
    doc.add_heading("7.2 Redundancy management", level=2)
    doc.add_paragraph(
        "education and education_num represent the same ordered concept. Both remain "
        "in the readable cleaned file, but only education_num is used in the model-ready "
        "table to avoid duplicate representations."
    )
    doc.add_heading("7.3 Categorical encoding", level=2)
    doc.add_paragraph(
        "Seven nominal predictors were one-hot encoded. Unknown remains a visible "
        "category. No reference category was dropped because the output is intended "
        "as a general model-ready matrix; a later model can choose its own encoding "
        "constraint."
    )
    add_code_block(
        doc,
        """
encoded_categories = pd.get_dummies(
    cleaned[model_categories],
    prefix=model_categories,
    drop_first=False,
    dtype="int8",
)
        """,
    )
    doc.add_heading("7.4 Robust numeric scaling", level=2)
    doc.add_paragraph(
        "Numeric features were centered by their median and divided by their IQR. "
        "Compared with mean/standard-deviation scaling, this approach is less affected "
        "by long tails. A divisor of 1 is used if the IQR is zero."
    )
    add_code_block(
        doc,
        """
median = model_numeric[column].median()
iqr = model_numeric[column].quantile(0.75) - model_numeric[column].quantile(0.25)
divisor = iqr if iqr != 0 else 1.0
scaled_numeric[column] = (model_numeric[column] - median) / divisor
        """,
    )
    add_callout(
        doc,
        "Leakage warning for future modelling",
        "This Week 1 output demonstrates preprocessing on the supplied dataset. In a "
        "predictive project, split the data first and estimate caps, medians, IQRs, "
        "and category vocabularies on the training set only; then apply them unchanged "
        "to validation and test data.",
        "FBEBDE",
    )
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # Results
    doc.add_heading("8. Results and output validation", level=1)
    add_table(
        doc,
        ["Metric", "Before", "After"],
        [
            ["Rows", "32,561", "32,537"],
            ["Columns", "15", "27 cleaned / 93 model-ready"],
            ["Exact duplicate rows", "24", "0"],
            ["Placeholder or null cells", "4,262", "0"],
            ["Numeric domain failures", "0", "0"],
            ["Education mapping conflicts", "0", "0"],
        ],
    )
    doc.add_paragraph(
        "The cleaned file is designed for human inspection and retains original values "
        "alongside engineered alternatives. The model-ready file contains six numeric "
        "features, one-hot encoded categories, and the income_high target."
    )
    doc.add_picture(str(CONSOLE_IMAGE), width=Inches(6.7))
    caption = doc.add_paragraph("Figure 1. Successful execution captured from the supplied Python program.")
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.runs[0].italic = True
    caption.runs[0].font.size = Pt(9)
    caption.runs[0].font.color.rgb = RGBColor.from_string(GRAY)

    doc.add_heading("8.1 Independent reconciliation", level=2)
    doc.add_paragraph(
        "The row count reconciles independently: 32,561 loaded rows − 24 exact "
        "duplicates = 32,537 cleaned rows. The cleaned income counts also reconcile: "
        "24,698 + 7,839 = 32,537. These checks confirm that no unexplained row loss "
        "occurred."
    )

    # Challenges
    doc.add_heading("9. Challenges and how they were overcome", level=1)
    add_table(
        doc,
        ["Challenge", "Risk", "Resolution"],
        [
            [
                "Hidden “?” values",
                "Null audit falsely reports complete data",
                "Strip text, replace placeholders with pandas.NA, then recount",
            ],
            [
                "No unique identifier",
                "Incorrectly deleting legitimate similar people",
                "Remove exact full-row duplicates only",
            ],
            [
                "Zero-inflated financial fields",
                "IQR method marks all positive values as outliers",
                "Use IQR as a flag; keep raw values and add log1p features",
            ],
            [
                "Hours concentrated at 40",
                "27.67% flagged by a mechanical rule",
                "Preserve observations; create capped and robust-scaled alternatives",
            ],
            [
                "Duplicate education representations",
                "Redundant model features",
                "Validate consistency and retain only education_num in model-ready data",
            ],
            [
                "Class imbalance",
                "Accuracy may hide weak >50K detection",
                "Preserve distribution and recommend stratified evaluation",
            ],
        ],
    )

    # Impact and ethics
    doc.add_heading("10. Impact on subsequent analysis", level=1)
    add_bullets(
        doc,
        [
            "Explicit Unknown categories preserve sample size and allow missingness itself to be analysed.",
            "Duplicate removal prevents repeated records from slightly distorting counts and model fitting.",
            "Capping and robust scaling reduce the leverage of extreme numeric values in distance- or gradient-based methods.",
            "Log transforms make highly skewed financial variables easier to model, but coefficients must be interpreted on the transformed scale.",
            "One-hot encoding expands dimensionality and can increase memory use; regularization may be useful.",
            "Keeping outlier flags enables sensitivity analysis, but the flags should not be interpreted as proof of error.",
            "The 24.09% high-income class share means accuracy alone is insufficient; precision, recall, F1, ROC-AUC, and a confusion matrix should be considered.",
        ],
    )
    doc.add_heading("10.1 Limitations and ethical considerations", level=2)
    doc.add_paragraph(
        "The dataset is historical and reflects social and labour-market conditions "
        "of its collection period. Variables such as race and sex are sensitive and "
        "may encode structural inequities. A high-performing model can still be unfair. "
        "Any future predictive use should include subgroup performance checks, fairness "
        "assessment, clear purpose limitation, and human oversight. The dataset should "
        "not be treated as a current or complete representation of income."
    )
    doc.add_paragraph(
        "The “Unknown” treatment is transparent but not neutral: if missingness is "
        "associated with the target, models may learn that association. Alternative "
        "complete-case or imputation analyses can be run as sensitivity tests."
    )

    # Conclusion
    doc.add_heading("11. Conclusion", level=1)
    doc.add_paragraph(
        "The project produced a reproducible and auditable data-preparation workflow. "
        "It corrected formatting inconsistencies, exposed hidden missing values, removed "
        "only defensible duplicates, validated ranges and category relationships, and "
        "handled outliers without conflating rarity with error. The resulting cleaned "
        "and model-ready datasets are suitable foundations for further exploration, "
        "visualization, or predictive modelling, subject to the leakage and fairness "
        "cautions documented above."
    )
    add_callout(
        doc,
        "Final deliverables",
        "adult_data_cleaning.py, adult_cleaned.csv, adult_model_ready.csv, "
        "quality_summary.csv, outlier_summary.csv, cleaning_metadata.json, "
        "README.md, requirements.txt, and this editable DOCX report.",
        SOFT_GREEN,
    )

    # Reproducibility
    doc.add_heading("12. Reproduction steps in VS Code", level=1)
    add_numbered(
        doc,
        [
            "Install Python 3.10 or newer and VS Code with the Python extension.",
            "Open the project folder and confirm adultfin.csv is beside adult_data_cleaning.py.",
            "Create and activate a virtual environment if desired.",
            "Run pip install -r requirements.txt in the VS Code terminal.",
            "Run python adult_data_cleaning.py.",
            "Review the outputs folder and compare the console summary with Figure 1.",
        ],
    )
    add_code_block(
        doc,
        """
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python adult_data_cleaning.py
        """,
    )

    # References
    doc.add_heading("References", level=1)
    references = [
        "Becker, B., & Kohavi, R. (1996). Adult [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5XW20",
        "UCI Machine Learning Repository. Adult dataset page. https://archive.ics.uci.edu/dataset/2/adult",
        "McKinney, W. pandas documentation. https://pandas.pydata.org/docs/",
        "Tukey, J. W. (1977). Exploratory Data Analysis. Addison-Wesley.",
    ]
    for number, reference in enumerate(references, start=1):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.left_indent = Inches(0.24)
        paragraph.paragraph_format.first_line_indent = Inches(-0.24)
        paragraph.add_run(f"{number}.  {reference}")

    doc.add_heading("Appendix A — Output inventory", level=1)
    add_table(
        doc,
        ["File", "Purpose"],
        [
            ["adult_cleaned.csv", "Readable, audited data with original and engineered features"],
            ["adult_model_ready.csv", "Numeric model matrix with encoded target"],
            ["quality_summary.csv", "Before/after data-quality reconciliation"],
            ["outlier_summary.csv", "IQR fences, counts, and percentages"],
            ["cleaning_metadata.json", "Thresholds, scaling parameters, and run metadata"],
            ["run_output.txt", "Captured terminal output"],
        ],
    )

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_heading("Appendix B — Submission checklist", level=1)
    checklist = [
        "Public dataset and source documented",
        "Data structure explored",
        "Missing values identified and handled",
        "Duplicates and erroneous entries checked",
        "Outliers identified and addressed",
        "Preprocessing steps coded and explained",
        "Challenges and solutions discussed",
        "Impact on subsequent analysis evaluated",
        "Reproducible VS Code instructions included",
        "Editable DOCX report supplied",
    ]
    for item in checklist:
        paragraph = doc.add_paragraph()
        run = paragraph.add_run("☑  " + item)
        run.font.name = "Arial"
        run.font.size = Pt(10.5)

    doc.save(REPORT_PATH)
    print(f"Created {REPORT_PATH}")


if __name__ == "__main__":
    main()