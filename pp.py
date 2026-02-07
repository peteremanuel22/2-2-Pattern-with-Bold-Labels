# app.py
import io
import os
import streamlit as st
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Cm, Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import black

# ---------------------------
# Configuration
# ---------------------------
MAX_FILES = 150

# A4 landscape (cm)
A4_WIDTH_CM = 29.7
A4_HEIGHT_CM = 21.0

# For PDF (points)
A4_LANDSCAPE = landscape(A4)  # (width, height) in points: ~ (842, 595)

# Spacing between pictures (inches on PowerPoint/PDF ruler)
SPACING_INCH = 0.25  # fixed horizontal and vertical gap between the 4 pictures

# Label (text) band height as a fraction of each cell height
TEXT_FRACTION = 0.12  # 12% of each cell height reserved for the label

# Font settings for the label above each picture
LABEL_FONT_SIZE_PT = 24
LABEL_FONT_BOLD = True

# JPEG compression quality (aim ~50%)
JPEG_QUALITY = 50

# ---------------------------
# Utilities
# ---------------------------
def get_blank_layout(prs: Presentation):
    """Return the 'Blank' layout if available; fallback to the first layout."""
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return prs.slide_layouts[0]

def set_white_background(slide):
    """Ensure the slide background is solid white."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)

def pil_open_transposed(file_bytes: bytes) -> Image.Image:
    """Open image from bytes and fix EXIF orientation."""
    img = Image.open(io.BytesIO(file_bytes))
    img = ImageOps.exif_transpose(img)
    return img

def rotate_if_portrait(img: Image.Image) -> Image.Image:
    """Rotate 90° clockwise if image is portrait."""
    if img.height > img.width:
        return img.rotate(270, expand=True)
    return img

def to_jpeg_bytes(img: Image.Image, quality: int = 50) -> bytes:
    """
    Convert image to JPEG with specified quality.
    If image has alpha, composite over white first.
    """
    if img.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        alpha = img.split()[-1]
        bg.paste(img.convert("RGB"), mask=alpha)
        img_rgb = bg
    else:
        img_rgb = img.convert("RGB")

    buf = io.BytesIO()
    img_rgb.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()

def sanitize_title_from_name(name: str) -> str:
    base = os.path.basename(name)
    stem = os.path.splitext(base)[0]
    return stem.strip() or "Slide"

# ---------------------------
# PPTX Pattern Placement (2×2 with fixed gaps)
# ---------------------------
def add_2x2_pattern_with_labels_blank_slide_pptx(
    slide,
    prs,
    jpeg_bytes: bytes,
    img_w: int,
    img_h: int,
    label_text: str,
    spacing_inch: float = SPACING_INCH,
    text_fraction: float = TEXT_FRACTION,
    label_font_pt: int = LABEL_FONT_SIZE_PT,
    label_bold: bool = LABEL_FONT_BOLD,
):
    """
    Place the same image 4 times in a 2×2 grid with a centered label above each image.
    The entire pattern is maximized to fill the slide while preserving aspect ratio.
    Gaps between the pictures are fixed to `spacing_inch` (horizontal & vertical).
    """

    # Slide dimensions in EMUs
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    # Fixed gaps (in EMUs)
    gh = Inches(spacing_inch)
    gv = Inches(spacing_inch)

    # Image aspect ratio
    r_img = img_w / img_h if img_h else 1.0

    # Relationship between cell width and height to allow the image to fit under its label without distortion:
    # cell_width = r_img * (1 - text_fraction) * cell_height
    r_cell = r_img * (1.0 - text_fraction)
    if r_cell <= 0:
        r_cell = 1e-6  # avoid division by zero

    # Maximize cell height ch subject to width/height constraints
    ch_from_width = (slide_w - gh) / (2.0 * r_cell)
    ch_from_height = (slide_h - gv) / 2.0
    ch = min(ch_from_width, ch_from_height)
    if ch <= 0:
        gh = gv = 0
        ch_from_width = slide_w / (2.0 * r_cell)
        ch_from_height = slide_h / 2.0
        ch = max(1, min(ch_from_width, ch_from_height))

    cw = r_cell * ch
    text_h = ch * text_fraction
    img_area_h = ch - text_h

    # Total pattern dimensions
    pattern_w = 2 * cw + gh
    pattern_h = 2 * ch + gv

    # Center the pattern on the slide
    pattern_left = (slide_w - pattern_w) / 2.0
    pattern_top = (slide_h - pattern_h) / 2.0

    # Compute cell origins
    def cell_origin(r: int, c: int):
        left = pattern_left + (cw + gh) * c
        top = pattern_top + (ch + gv) * r
        return left, top

    # Helper: add one labeled image in a cell
    def add_labeled_image_at(left, top):
        # Text box (centered)
        tbox = slide.shapes.add_textbox(int(left), int(top), int(cw), int(text_h))
        tf = tbox.text_frame
        tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # vertically center within the text band

        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label_text
        run.font.size = Pt(label_font_pt)
        run.font.bold = label_bold

        # Image area (below the text band) — width-limited exact fit
        img_left = left
        img_top = top + text_h

        buf = io.BytesIO(jpeg_bytes)
        pic = slide.shapes.add_picture(buf, int(img_left), int(img_top), width=int(cw))
        # Center inside designated image area to counter tiny rounding differences
        pic.left = int(img_left + (cw - pic.width) / 2.0)
        pic.top = int(img_top + (img_area_h - pic.height) / 2.0)

    # Add 4 cells: (row 0..1, col 0..1)
    for r in range(2):
        for c in range(2):
            left, top = cell_origin(r, c)
            add_labeled_image_at(left, top)

# ---------------------------
# PDF Pattern Placement (2×2 with fixed gaps)
# ---------------------------
def add_2x2_pattern_with_labels_pdf(
    c: canvas.Canvas,
    page_w: float,
    page_h: float,
    jpeg_bytes: bytes,
    img_w: int,
    img_h: int,
    label_text: str,
    spacing_inch: float = SPACING_INCH,
    text_fraction: float = TEXT_FRACTION,
    label_font_pt: int = LABEL_FONT_SIZE_PT,
):
    """
    Draw the same image 4 times in a 2×2 grid with a centered bold label above each image
    on a reportlab canvas page. Pattern fills the page while preserving aspect ratio.
    """

    # Fixed gaps (points)
    gap_h = 72.0 * spacing_inch
    gap_v = 72.0 * spacing_inch

    # Image aspect ratio
    r_img = img_w / img_h if img_h else 1.0
    r_cell = r_img * (1.0 - text_fraction)
    if r_cell <= 0:
        r_cell = 1e-6

    # Maximize cell height ch subject to constraints in points
    ch_from_width = (page_w - gap_h) / (2.0 * r_cell)
    ch_from_height = (page_h - gap_v) / 2.0
    ch = min(ch_from_width, ch_from_height)
    if ch <= 0:
        gap_h = gap_v = 0
        ch_from_width = page_w / (2.0 * r_cell)
        ch_from_height = page_h / 2.0
        ch = max(1, min(ch_from_width, ch_from_height))

    cw = r_cell * ch
    text_h = ch * text_fraction
    img_area_h = ch - text_h

    pattern_w = 2 * cw + gap_h
    pattern_h = 2 * ch + gap_v

    # Center the pattern on the page (origin is bottom-left in reportlab)
    pattern_left = (page_w - pattern_w) / 2.0
    pattern_bottom = (page_h - pattern_h) / 2.0

    # Prepare image reader from JPEG bytes
    img_reader = ImageReader(io.BytesIO(jpeg_bytes))

    # Label font (bold)
    c.setFont("Helvetica-Bold", label_font_pt)
    c.setFillColor(black)

    # Draw 4 cells
    for r in range(2):
        for col in range(2):
            x = pattern_left + (cw + gap_h) * col
            y = pattern_bottom + (ch + gap_v) * r  # bottom of the cell

            # Text band is at the top part of the cell
            # Center text baseline roughly at middle of text band
            text_center_y = y + ch - (text_h / 2.0)
            text_center_x = x + (cw / 2.0)
            c.drawCentredString(text_center_x, text_center_y, label_text)

            # Image area bottom
            img_bottom = y
            # Draw image exactly filling width=cw and height=img_area_h
            c.drawImage(
                img_reader,
                x,
                img_bottom,
                width=cw,
                height=img_area_h,
                preserveAspectRatio=True,  # harmless; exact fit already computed
                anchor='sw',               # anchored at (x, y) bottom-left
            )

# ---------------------------
# Streamlit App
# ---------------------------
st.set_page_config(page_title="Images → A4 Landscape (2×2 Pattern, Blank) PPTX/PDF", page_icon="📄", layout="centered")
st.title("📄 A4 Landscape — 2×2 Pattern with Bold Labels (Output: PPTX or PDF)")
st.caption(
    "Office theme, white background. Each image becomes one blank slide/page showing the same picture 4 times "
    "in a 2×2 grid with labels (filename without extension) above each copy, bold and size 24. "
    "Portrait images are auto-rotated 90°; images are compressed to ~50% JPEG quality. "
    "The pattern fills the slide/page while keeping aspect ratio, with 0.25\" gaps between pictures."
)

uploaded_files = st.file_uploader(
    "Upload 1 to 150 images",
    type=["jpg", "jpeg", "png", "bmp", "tif", "tiff", "webp"],
    accept_multiple_files=True,
    help="Each uploaded image will produce one slide/page with a 2×2 pattern of that image."
)

output_format = st.radio("Output format", ["PPTX", "PDF"], horizontal=True)

default_name_pptx = "images_to_a4_landscape_2x2_pattern_blank_bold24.pptx"
default_name_pdf = "images_to_a4_landscape_2x2_pattern_blank_bold24.pdf"
default_output_name = default_name_pptx if output_format == "PPTX" else default_name_pdf
output_name = st.text_input("Output filename", value=default_output_name)

generate = st.button("Generate")

if generate:
    if not uploaded_files:
        st.error("Please upload at least 1 image.")
    elif len(uploaded_files) > MAX_FILES:
        st.error(f"Please upload at most {MAX_FILES} images.")
    else:
        try:
            progress = st.progress(0)
            status = st.empty()

            if output_format == "PPTX":
                # Create PPTX (Office theme, blank, white)
                prs = Presentation()
                prs.slide_width = Cm(A4_WIDTH_CM)
                prs.slide_height = Cm(A4_HEIGHT_CM)
                blank_layout = get_blank_layout(prs)

                for idx, up in enumerate(uploaded_files, start=1):
                    file_bytes = up.read()
                    img = pil_open_transposed(file_bytes)
                    img = rotate_if_portrait(img)

                    # Compress to ~50% quality JPEG
                    jpeg_bytes = to_jpeg_bytes(img, quality=JPEG_QUALITY)
                    label_text = sanitize_title_from_name(up.name)

                    slide = prs.slides.add_slide(blank_layout)
                    set_white_background(slide)

                    add_2x2_pattern_with_labels_blank_slide_pptx(
                        slide=slide,
                        prs=prs,
                        jpeg_bytes=jpeg_bytes,
                        img_w=img.width,
                        img_h=img.height,
                        label_text=label_text,
                        spacing_inch=SPACING_INCH,
                        text_fraction=TEXT_FRACTION,
                        label_font_pt=LABEL_FONT_SIZE_PT,
                        label_bold=LABEL_FONT_BOLD,
                    )

                    progress.progress(idx / len(uploaded_files))
                    status.text(f"[PPTX] Processing image {idx}/{len(uploaded_files)}: {up.name}")

                # Save to buffer and provide download
                out_buf = io.BytesIO()
                prs.save(out_buf)
                out_buf.seek(0)
                st.success("PPTX generated successfully!")
                st.download_button(
                    label="⬇️ Download PPTX",
                    data=out_buf,
                    file_name=(output_name.strip() or default_name_pptx),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

            else:  # PDF
                # Create PDF with reportlab
                page_w, page_h = A4_LANDSCAPE
                out_buf = io.BytesIO()
                cpdf = canvas.Canvas(out_buf, pagesize=A4_LANDSCAPE)

                for idx, up in enumerate(uploaded_files, start=1):
                    file_bytes = up.read()
                    img = pil_open_transposed(file_bytes)
                    img = rotate_if_portrait(img)

                    # Compress to ~50% quality JPEG
                    jpeg_bytes = to_jpeg_bytes(img, quality=JPEG_QUALITY)
                    label_text = sanitize_title_from_name(up.name)

                    # Optional: ensure white background (page is white by default)
                    cpdf.setFillColorRGB(1, 1, 1)
                    cpdf.rect(0, 0, page_w, page_h, fill=1, stroke=0)

                    # Draw the 2x2 pattern for this page
                    add_2x2_pattern_with_labels_pdf(
                        c=cpdf,
                        page_w=page_w,
                        page_h=page_h,
                        jpeg_bytes=jpeg_bytes,
                        img_w=img.width,
                        img_h=img.height,
                        label_text=label_text,
                        spacing_inch=SPACING_INCH,
                        text_fraction=TEXT_FRACTION,
                        label_font_pt=LABEL_FONT_SIZE_PT,
                    )

                    cpdf.showPage()
                    progress.progress(idx / len(uploaded_files))
                    status.text(f"[PDF] Processing image {idx}/{len(uploaded_files)}: {up.name}")

                cpdf.save()
                out_buf.seek(0)
                st.success("PDF generated successfully!")
                st.download_button(
                    label="⬇️ Download PDF",
                    data=out_buf,
                    file_name=(output_name.strip() or default_name_pdf),
                    mime="application/pdf",
                )

        except Exception as e:
            st.error(f"An error occurred: {e}")
