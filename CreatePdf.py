from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors

doc = SimpleDocTemplate(
    "colored_auto_pages.pdf",
    pagesize=A4,
    leftMargin=20*mm, rightMargin=20*mm,
    topMargin=20*mm, bottomMargin=20*mm
)

styles = getSampleStyleSheet()

# Base style for most text
base = ParagraphStyle(
    'base',
    parent=styles['Normal'],
    fontName='Helvetica',
    fontSize=14,
    leading=18
)

# Whole-paragraph color via style
green_paragraph = ParagraphStyle(
    'green_paragraph',
    parent=base,
    textColor=colors.green
)

story = []

# Whole paragraph colored by style
story.append(Paragraph("This entire paragraph is green.", green_paragraph))
story.append(Spacer(1, 8))

# Color using inline HTML-like tags
story.append(Paragraph('<font color="#d32f2f">This line is red (hex).</font>', base))
story.append(Spacer(1, 8))
story.append(Paragraph('<font color="blue">This line is blue (named color).</font>', base))
story.append(Spacer(1, 8))

# Mixed colors on a single line
story.append(Paragraph(
    'Mixed colors in one line: '
    '<font color="#2e7d32">green</font>, '
    '<font color="#ef6c00">orange</font>, and '
    '<font color="#1565c0">blue</font>.',
    base
))
story.append(Spacer(1, 16))

# Demonstrate automatic page breaks with many lines
for i in range(1, 120):
    story.append(Paragraph(
        f'Line {i}: flows across pages with '
        '<font color="#455A64">automatic page breaks</font>.',
        base
    ))

doc.build(story)
