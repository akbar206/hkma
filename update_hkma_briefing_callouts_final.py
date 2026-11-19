from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor

INPUT = '/mnt/data/HKMA_Threat_Landscape_One_Slide_Revised.pptx'
OUTPUT = '/mnt/data/HKMA_Threat_Landscape_One_Slide_Callouts_Final.pptx'

TOP_TEXT = (
    "The cyber threat landscape continues to intensify, with critical vulnerabilities increasingly exploited within hours of disclosure and major cyber incidents causing significant financial and operational disruption. "
    "AI is further accelerating attacker capabilities, while the following high-profile cases illustrate the growing speed, scale and cross-sector impact of cyber threats."
)

BOTTOM_TEXT = (
    "HKMA has assessed emerging threats using intelligence from multiple sources and, where needed, taken timely mitigation measures, including emergency patching and temporary service suspension. "
    "We will continue to monitor the evolving threat landscape and strengthen response readiness, with further vulnerability management enhancements covered under Section 8 (VME)."
)

TITLE_TEXT = '2. Latest Cyber Threat Intelligence – The New Vulnerability Norm'
PURPLE = RGBColor(102, 45, 145)
BLUE_FILL = RGBColor(91, 155, 213)
BLUE_LINE = RGBColor(55, 96, 146)
WHITE = RGBColor(255, 255, 255)

prs = Presentation(INPUT)
slide = prs.slides[0]

# In the revised deck, shapes 2-55 correspond to the legend + timeline region.
TIMELINE_IDS = list(range(2, 56))


def remove_shape(shape):
    el = shape._element
    el.getparent().remove(el)


def scale_shapes(shape_ids, sx=0.96, sy=0.88):
    shapes = [slide.shapes[i] for i in shape_ids]
    left = min(s.left for s in shapes)
    top = min(s.top for s in shapes)
    right = max(s.left + s.width for s in shapes)
    bottom = max(s.top + s.height for s in shapes)
    cx = (left + right) / 2
    cy = (top + bottom) / 2
    for s in shapes:
        old_left, old_top, old_w, old_h = s.left, s.top, s.width, s.height
        new_w = int(old_w * sx)
        new_h = int(old_h * sy)
        old_cx = old_left + old_w / 2
        old_cy = old_top + old_h / 2
        new_cx = cx + (old_cx - cx) * sx
        new_cy = cy + (old_cy - cy) * sy
        s.left = int(new_cx - new_w / 2)
        s.top = int(new_cy - new_h / 2)
        s.width = new_w
        s.height = new_h


def set_title():
    for sh in slide.shapes:
        if hasattr(sh, 'text') and 'Latest Cyber Threat Intelligence' in (sh.text or ''):
            tf = sh.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = TITLE_TEXT
            run.font.name = 'Aptos Display'
            run.font.size = Pt(24)
            run.font.bold = True
            run.font.color.rgb = PURPLE
            sh.left = Inches(0.80)
            sh.top = Inches(0.08)
            sh.width = Inches(10.60)
            sh.height = Inches(0.58)
            return


def set_legend_positions():
    # Move legend to bottom-right, below the bottom callout.
    for sh in slide.shapes:
        if not hasattr(sh, 'text'):
            continue
        txt = (sh.text or '').strip()
        if txt == 'Software vulnerability':
            sh.left = Inches(8.35)
            sh.top = Inches(6.78)
            sh.width = Inches(1.20)
            sh.height = Inches(0.18)
        elif txt == 'Platform or provider compromise':
            sh.left = Inches(9.95)
            sh.top = Inches(6.78)
            sh.width = Inches(1.95)
            sh.height = Inches(0.18)

    # The legend markers are the two smallest empty rounded rectangles.
    empties = [sh for sh in slide.shapes if hasattr(sh, 'text') and (sh.text or '').strip() == '' and sh.width < 250000 and sh.height < 160000]
    empties = sorted(empties, key=lambda s: s.left)
    if len(empties) >= 2:
        empties[-2].left = Inches(8.13)
        empties[-2].top = Inches(6.78)
        empties[-2].width = Inches(0.18)
        empties[-2].height = Inches(0.14)
        empties[-1].left = Inches(9.73)
        empties[-1].top = Inches(6.78)
        empties[-1].width = Inches(0.18)
        empties[-1].height = Inches(0.14)


def add_callout(left, top, width, height, text, font_size):
    shp = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = BLUE_FILL
    shp.line.color.rgb = BLUE_LINE
    shp.line.width = Pt(1.4)

    tf = shp.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    tf.margin_left = Pt(7)
    tf.margin_right = Pt(7)
    tf.margin_top = Pt(4)
    tf.margin_bottom = Pt(4)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = 'Aptos'
    run.font.size = Pt(font_size)
    run.font.color.rgb = WHITE
    return shp


# Apply layout updates.
scale_shapes(TIMELINE_IDS, sx=0.96, sy=0.88)
set_title()
set_legend_positions()

# Remove old plain top/bottom note text boxes.
for sh in list(slide.shapes):
    if not hasattr(sh, 'text'):
        continue
    txt = (sh.text or '').strip()
    if txt.startswith('Recent cyber incidents show') or txt.startswith('The cyber threat landscape continues') or txt.startswith('HKMA has assessed emerging threats'):
        remove_shape(sh)

# Add rounded note callouts.
add_callout(Inches(0.95), Inches(1.08), Inches(9.55), Inches(0.52), TOP_TEXT, 11.3)
add_callout(Inches(0.95), Inches(5.88), Inches(9.55), Inches(0.54), BOTTOM_TEXT, 10.9)

prs.save(OUTPUT)
print(f'Saved: {OUTPUT}')
