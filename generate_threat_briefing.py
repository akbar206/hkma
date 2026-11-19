"""Generate a two-slide HKMA cyber threat briefing using the supplied PPTX template.

Install: python3 -m pip install python-pptx
Run:     python3 generate_threat_briefing.py

Keep Threat_Landscape_Timeline_v3.pptx beside this script (or edit TEMPLATE).
"""
from copy import deepcopy
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / 'Threat_Landscape_Timeline_v3.pptx'
OUTPUT = HERE / 'HKMA_Cyber_Threat_Briefing_Nov2026.pptx'
PURPLE = RGBColor(68, 43, 108)
TEAL = RGBColor(16, 124, 140)
INK = RGBColor(42, 38, 50)
GRAY = RGBColor(98, 95, 104)
LAVENDER = RGBColor(243, 239, 249)
PALE_TEAL = RGBColor(235, 246, 247)
WHITE = RGBColor(255, 255, 255)

# Kept deliberately short to fit the source template's precise card geometry.
# Each tuple: date, category, label, leading fact, qualifier, impact, response.
EVENTS = [
 ('25 Sep','vuln','SHAREPOINT SERVER','SharePoint','Exploited flaw',
  'CISA listed active exploitation', 'Urgent exposure review and patching'),
 ('27 Sep','vuln','CITRIX NETSCALER','Citrix NetScaler','2 zero-days',
  'Remote-access products targeted', 'Citrix patches; CISA warning'),
 ('20 Oct','incident','CLOUD HOSTING','GMO Internet','26,739 servers',
  'Hosting environment exposed', 'Service restrictions and response'),
 ('21 Oct','vuln','WEB PLATFORM','WordPress','< 48 h',
  'Rapid scanning after disclosure', 'Emergency updates advised'),
 ('21 Oct','vuln','WEB RUNTIME','PHP2Shell','Active attacks',
  'Web and security products affected', 'Emergency mitigation and patches'),
 ('29 Oct','incident','CRYPTO EXCHANGE','Binance','US$6.36B',
  'Major digital-asset theft reported', 'Withdrawals and incident response'),
 ('12 Nov','vuln','CITRIX NETSCALER','Citrix NetScaler','28 flaws',
  'Multiple Gateway / ADC weaknesses', 'Vendor security fixes released'),
 ('16 Nov','incident','SOFTWARE SUPPLIER','Optimizely','178 orgs',
  'Malicious scripts via trusted vendor', 'Vendor removed injected code'),
]

SOURCES = [
 'SharePoint active exploitation: https://www.cyber.gc.ca/en/alerts-advisories/al26-023-vulnerability-impacting-microsoft-sharepoint-server-cve-2026-65660',
 'Citrix September: https://content.govdelivery.com/accounts/USDHSCISA/bulletins/42cc465',
 'Citrix November: https://cert.europa.eu/publications/security-advisories/2026-020',
 'PHP2Shell: https://www.csa.gov.sg/alerts-and-advisories/alerts/al-2026-157/',
 'Microsoft 365 customer impact: https://support.akumina.com/hc/en-us/articles/50349245338509-Service-degradation-on-Microsoft-365-Business-or-Enterprise-11-12-2026',
 'Microsoft 365 technical cause: https://www.imss.caltech.edu/news/m365-email-issue-12nov',
 'Optimizely official: https://status.optimizely.com/history',
 'Optimizely 178 orgs (research count): https://www.bleepingcomputer.com/news/security/optimizely-supply-chain-attack-delivers-malware-to-cnn-fox-news-bbc-salesforce-sites/',
]


def replace_paragraph(shape, i, value):
    """Update text without losing the template's original run formatting."""
    p = shape.text_frame.paragraphs[i]
    if p.runs:
        p.runs[0].text = value
        for r in p.runs[1:]: r.text = ''
    else:
        p.add_run().text = value


def update_card(slide, indices, event, node_index):
    date, kind, label, name, fact, impact, response = event
    header, stat, info = (slide.shapes[k] for k in indices)
    replace_paragraph(header,0,label)
    replace_paragraph(header,1,name)
    replace_paragraph(stat,0,fact)
    replace_paragraph(stat,1,'Sep–Nov 2026 case')
    replace_paragraph(info,0,'Impact: ' + impact)
    replace_paragraph(info,1,'Response: ' + response)
    replace_paragraph(slide.shapes[node_index],0,date)
    # Match the original teal-vulnerability / purple-incident visual system.
    color = TEAL if kind == 'vuln' else PURPLE
    header.fill.solid(); header.fill.fore_color.rgb = color
    node = slide.shapes[node_index]
    node.line.color.rgb = color
    # The source's second card text is small: prevent unexpected spill.
    for sh in (header, stat, info):
        sh.text_frame.word_wrap = True


def txt(slide, x,y,w,h, text, size=12, color=INK, bold=False, align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(.035);tf.margin_right = Inches(.035)
    tf.margin_top = Inches(.025);tf.margin_bottom = Inches(.015)
    p = tf.paragraphs[0]
    p.text=text
    if align: p.alignment=align
    for run in p.runs:
        run.font.name='Arial';run.font.size=Pt(size)
        run.font.bold=bold;run.font.color.rgb=color
    return box


def rect(slide,x,y,w,h,fill, border=None, rounding=False):
    r=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounding else MSO_SHAPE.RECTANGLE,
         Inches(x),Inches(y),Inches(w),Inches(h))
    r.fill.solid();r.fill.fore_color.rgb=fill
    if border:r.line.color.rgb=border
    else:r.line.fill.background()
    return r


def section(slide,x,y,w,h,title,body,color_bg,accent):
    rect(slide,x,y,w,h,color_bg,rounding=True)
    rect(slide,x,y,.055,h,accent)
    txt(slide,x+.17,y+.10,w-.35,.31,title,15,accent,True)
    txt(slide,x+.17,y+.51,w-.35,h-.58,body,11,INK)


def add_sources_notes(slide, text):
    try:
        slide.notes_slide.notes_text_frame.text = text
    except Exception:
        pass


def make_second_slide(prs):
    # A new blank-layout slide inherits the template's theme and slide size.
    s = prs.slides.add_slide(prs.slide_layouts[6])
    # Copy the exact purple gradient-like left rail from the original slide.
    source = prs.slides[0]
    for name in ['Rectangle 3']:
        for sh in source.shapes:
            if sh.name == name:
                s.shapes._spTree.insert_element_before(deepcopy(sh.element), 'p:extLst')
                break
    txt(s,.91,.29,11.7,.54,'3. Latest Cyber Threat Intelligence – Third-Party Dependencies',25,PURPLE,True)
    txt(s,1.42,.88,11.15,.37,
        'Two supplier incidents illustrate how external technology dependencies can affect operational resilience and cybersecurity.',11,GRAY)

    # Microsoft: top 60%, detailed security-tool impact.
    rect(s,1.38,1.40,10.85,.38,PURPLE,rounding=True)
    txt(s,1.55,1.44,10.4,.30,'MICROSOFT 365  |  GLOBAL SERVICE OUTAGE  |  12–13 NOV 2026',14,WHITE,True)
    txt(s,1.48,1.88,10.8,.48,
        'Non-cyber failure: a latent network-gateway software defect, traffic increase and capacity constraints caused cascading disruption.',12,INK)
    rect(s,1.49,2.43,10.61,1.47,LAVENDER,rounding=True)
    txt(s,1.68,2.52,3.15,.33,'SECURITY FUNCTION',12,PURPLE,True)
    txt(s,5.02,2.52,6.75,.33,'REPORTED IMPACT',12,PURPLE,True)
    rect(s,1.65,2.93,10.1,.012,RGBColor(210,200,225))
    txt(s,1.68,3.00,3.2,.86,'Defender for Office 365 / XDR\nPurview security & compliance',11,INK,True)
    txt(s,5.02,3.00,6.75,.83,
        'Safe Links verification and certain portal functions could fail; DLP, labelling, classification and compliance services were disrupted.',11,INK)
    txt(s,1.51,4.00,10.65,.32,
        'External example: Microsoft-linked tools and email workflows were disrupted at organisations using the affected services.',10.5,GRAY)

    # Optimizely: concise supporting case in lower third.
    rect(s,1.38,4.44,10.85,.38,TEAL,rounding=True)
    txt(s,1.55,4.48,10.4,.30,'OPTIMIZELY  |  THIRD-PARTY SCRIPT COMPROMISE  |  16 NOV 2026',14,WHITE,True)
    rect(s,1.49,4.93,10.61,.92,PALE_TEAL,rounding=True)
    txt(s,1.69,5.03,10.1,.70,
        'Attackers exploited a vendor account-signup flaw to inject malicious JavaScript into customer website projects. Researchers identified affected scripts associated with 178 organisations; exposure does not establish that all were breached.',11,INK)

    rect(s,1.38,6.02,10.85,.68,RGBColor(239,237,244),rounding=True)
    txt(s,1.55,6.08,10.52,.53,
        'HKMA LESSON  |  HKMA was not impacted by the Microsoft outage. Future design and procurement should consider critical security-function availability, supplier concentration and fallback arrangements.',11,PURPLE,True)
    txt(s,1.46,6.98,10.8,.19,
        'Source links and qualification notes are included in PowerPoint speaker notes.  |  As of 19 November 2026',8.5,GRAY)
    add_sources_notes(s,'\n'.join(SOURCES)+'\n\nHKMA non-impact and cloud-independence statements come from internal manager notes, not externally verified sources.\nSecurity-tool impact reflects reported service-level disruptions, not proof that all protection stopped.')
    return s


def main():
    if not TEMPLATE.exists():
        raise FileNotFoundError(f'Put the original template beside this script: {TEMPLATE}')
    prs=Presentation(TEMPLATE)
    if len(prs.slides)!=1 or len(prs.slides[0].shapes)<57:
        raise ValueError('This script expects the original single-slide Threat_Landscape_Timeline_v3.pptx')
    slide=prs.slides[0]
    replace_paragraph(slide.shapes[56],0,'2. Latest Cyber Threat Intelligence – The New Vulnerability Norm')
    replace_paragraph(slide.shapes[1],0,
        'Sep–19 Nov 2026: exploited vulnerabilities and supplier-related incidents demonstrate rapidly evolving, cross-sector technology risks.')
    slots=[((10,11,12),48),((15,16,17),49),((20,21,22),50),((25,26,27),51),
           ((30,31,32),52),((35,36,37),53),((40,41,42),54),((45,46,47),55)]
    for event,(indices,node_idx) in zip(EVENTS,slots):
        update_card(slide,indices,event,node_idx)
    add_sources_notes(slide,'Timeline uses 25 Sep–16 Nov examples, within the requested Sep–19 Nov window. Dates are disclosure/report dates.\n'+'\n'.join(SOURCES)+'\n\nGMO, WordPress and Binance headline facts originated in the supplied source template; independently revalidate any numerical figures before management distribution.')
    make_second_slide(prs)
    prs.save(OUTPUT)
    print('Created:',OUTPUT)

if __name__=='__main__': main()
