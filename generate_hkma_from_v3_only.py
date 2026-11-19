"""Generate one-slide HKMA threat landscape briefing using ONLY the original v3 template.

Place this .py file and Threat_Landscape_Timeline_v3.pptx in the SAME folder.
Windows PowerShell:
    py -m pip install python-pptx
    py "FULL_PATH_TO_THIS_SCRIPT.py"

The output PowerPoint is created beside the script.
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

FOLDER = Path(__file__).resolve().parent
TEMPLATE = FOLDER / 'Threat_Landscape_Timeline_v3.pptx'
OUTPUT = FOLDER / 'HKMA_Threat_Landscape_Final.pptx'

TEAL = RGBColor(16, 124, 140)
PURPLE = RGBColor(68, 43, 108)
TITLE_PURPLE = RGBColor(102, 45, 145)
WHITE = RGBColor(255,255,255)
DARK = RGBColor(30,35,44)
BLUE = RGBColor(91,155,213)
BLUE_BORDER = RGBColor(55,96,146)

TOP = (
    'Cyber threats are escalating: critical vulnerabilities can be exploited within hours, while major breaches and supplier failures disrupt services across sectors. '
    'AI-assisted techniques may further accelerate attacks, increasing the importance of timely risk assessment and response.'
)
BOTTOM = (
    'HKMA has assessed emerging threats using intelligence from multiple sources and taken timely mitigation measures where necessary, including emergency patching and temporary service suspension. '
    'We will continue monitoring, with vulnerability management enhancements covered under Section 8 (VME).'
)

NOTES = '''Q&A | 19 NOVEMBER 2026

OPTIMIZELY (16 NOV) — SUPPLY-CHAIN CYBERATTACK
• Attackers exploited an Optimizely Web Experimentation signup-flow flaw to impersonate customer administrators; stolen customer passwords were not required.
• Modified project JavaScript was distributed through trusted vendor scripts and ran in visitors’ browsers; reported payloads included card skimming and cryptocurrency-payment redirection.
• Researchers identified 188 affected scripts associated with 178 organisations — NOT 178 confirmed data breaches. Optimizely removed the code and corrected the signup flaw.
• Lesson: Review third-party JavaScript on sensitive webpages and prepare rapid supplier-script disablement.
Source: Optimizely security incident report (17 Nov 2026):
https://pages.optimizely.com/rs/947-RKE-048/images/Security%20Incident%20Report%20-%20Optimizely%20Web%20Experimentation%20-%20Nov%2017%202026.pdf

MICROSOFT 365 (12–13 NOV) — NON-CYBER GLOBAL OUTAGE
• Cause: Software defect in network gateways, compounded by traffic failover and reduced capacity, caused cascading service failures; this was NOT a cyberattack.
• Security functions: Some Defender for Office 365, Purview, and security-management portal functions were disrupted. This does NOT mean all endpoint defences stopped.
• Wider impact: Microsoft 365-dependent integrations and communications at external organisations were interrupted; these were service impacts, not confirmed breaches.
• HKMA context (per manager): HKMA was not impacted; its security tools are not cloud-dependent. Procurement lesson: consider supplier concentration, fallback and security-tool availability.
Source: Microsoft preliminary post-incident report (15 Nov 2026):
https://www.365fanatic.com/wp-content/uploads/2026/11/PostIncidentReport.pdf

TIMELINE CAUTION
Potential exposure is not confirmed compromise. Issued patches are not proof of exploitation. Incident information is based on public reporting and should be checked before management circulation.'''

# date, category, header category, name, headline, one- to two-sentence management explanation
EVENTS = [
    ('25 Sep','vuln','SHAREPOINT SERVER','SharePoint','Exploited flaw',
     'Attackers exploited a SharePoint code-injection flaw, putting enterprise collaboration and sensitive documents at risk; CISA urged remediation.'),
    ('27 Sep','vuln','CITRIX NETSCALER','Citrix NetScaler','2 zero-days',
     'Attackers exploited unpatched NetScaler gateway flaws, exposing critical remote-access infrastructure to potential compromise.'),
    ('20 Oct','incident','CLOUD HOSTING','GMO Internet','26,739 servers',
     'Malware was found on GMO-hosted customer servers; service suspension disrupted customers, including financial-trading users.'),
    ('21 Oct','vuln','SHARED LIBRARY','Brotli / Node2Shell','Shared flaw',
     'Crafted compressed data exploited a common library, triggering emergency fixes across software including Node.js and .NET.'),
    ('21 Oct','vuln','WEB RUNTIME','PHP2Shell','Critical PHP flaws',
     'Unauthenticated code execution threatened exposed PHP web services, forcing urgent risk review and emergency patching.'),
    ('29 Oct','incident','CRYPTO EXCHANGE','Binance','~US$6.34B',
     'Attackers abused compromised access to manipulate exchange wallets, causing major digital-asset losses and emergency containment.'),
    ('12 Nov','vuln','CITRIX NETSCALER','Citrix NetScaler','28 flaws patched',
     'Citrix fixed 28 memory-related flaws in gateway products; organisations needed to assess and update exposed infrastructure.'),
    ('16 Nov','incident','SOFTWARE SUPPLIER','Optimizely','178 organisations',
     'Attackers altered trusted Optimizely scripts through administrator impersonation, exposing website visitors to potential card-skimming attacks.'),
]


def set_text(shape, text, fontsize, color=DARK, bold=False, align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(.015)
    tf.margin_right = Inches(.015)
    tf.margin_top = Inches(.01)
    tf.margin_bottom = Inches(.005)
    p=tf.paragraphs[0]
    p.alignment=align
    run=p.add_run()
    run.text=text
    run.font.name='Arial'
    run.font.size=Pt(fontsize)
    run.font.bold=bold
    run.font.color.rgb=color
    return shape


def heading(shape, category, name, color):
    tf=shape.text_frame
    tf.clear()
    tf.word_wrap=True
    tf.margin_left=Inches(.11)
    tf.margin_right=Inches(.04)
    tf.margin_top=Inches(.045)
    tf.margin_bottom=Inches(.015)
    for i,(text,pt,bold) in enumerate([(category,8,False),(name,10.5,True)]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.space_after=Pt(0)
        p.space_before=Pt(0)
        r=p.add_run();r.text=text;r.font.name='Arial';r.font.size=Pt(pt);r.font.bold=bold;r.font.color.rgb=WHITE
    shape.fill.solid();shape.fill.fore_color.rgb=color


def rounded_note(slide,x,y,w,h,text,pt):
    shape=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h))
    shape.fill.solid();shape.fill.fore_color.rgb=BLUE
    shape.line.color.rgb=BLUE_BORDER
    shape.line.width=Pt(1.1)
    set_text(shape,text,pt,WHITE,align=PP_ALIGN.CENTER)
    shape.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
    shape.text_frame.margin_left=Inches(.15)
    shape.text_frame.margin_right=Inches(.15)
    shape.text_frame.margin_top=Inches(.055)
    shape.text_frame.margin_bottom=Inches(.04)
    return shape


def resize_timeline(shapes):
    # Slight horizontal + vertical reduction, preserving the existing v3 geometry.
    left=min(s.left for s in shapes);right=max(s.left+s.width for s in shapes)
    top=min(s.top for s in shapes);bottom=max(s.top+s.height for s in shapes)
    midx=(left+right)/2;midy=(top+bottom)/2
    sx,sy=.96,.88
    for s in shapes:
        oldx=s.left+s.width/2;oldy=s.top+s.height/2
        nw=int(s.width*sx);nh=int(s.height*sy)
        s.left=int(midx+(oldx-midx)*sx-nw/2)
        s.top=int(midy+(oldy-midy)*sy-nh/2)
        s.width=nw;s.height=nh


def main():
    if not TEMPLATE.is_file():
        raise SystemExit('ERROR: Put Threat_Landscape_Timeline_v3.pptx in the same folder as this .py script. Looking for: '+str(TEMPLATE))
    prs=Presentation(TEMPLATE)
    if len(prs.slides)!=1 or len(prs.slides[0].shapes)<57:
        raise SystemExit('ERROR: The template must be the original, one-slide Threat_Landscape_Timeline_v3.pptx.')
    slide=prs.slides[0]
    # Edit original eight card groups, retaining outlines, connectors, timeline, theme.
    slots=[((10,11,12),48),((15,16,17),49),((20,21,22),50),((25,26,27),51),
           ((30,31,32),52),((35,36,37),53),((40,41,42),54),((45,46,47),55)]
    for (date,kind,category,name,fact,description),(indices,node_idx) in zip(EVENTS,slots):
        color=TEAL if kind=='vuln' else PURPLE
        header,stat,info=(slide.shapes[i] for i in indices)
        heading(header,category,name,color)
        set_text(stat,fact,15.8,DARK,True)
        set_text(info,description,8.7,DARK)
        set_text(slide.shapes[node_idx],date,7.4,color,True,PP_ALIGN.CENTER)
        slide.shapes[node_idx].line.color.rgb=color
        # Theme date node fill retained.
    # Move old subtitle out of the way; its contents become top callout.
    old=slide.shapes[1];old.text='';old.width=Inches(.01);old.height=Inches(.01)
    # Scale just timeline shapes: connectors/cards/date ovals (6-55), not header or legend.
    resize_timeline([slide.shapes[i] for i in range(6,56)])
    # Shift the central cards a few pixels down if needed to balance top/bottom spaces.
    for i in range(6,56):
        sh=slide.shapes[i]
        sh.top+=Inches(.03)
    # Move legend out of the main body into the empty right-hand footer.
    for i,x in [(2,8.13),(3,8.35),(4,9.73),(5,9.95)]:
        slide.shapes[i].left=Inches(x)
        slide.shapes[i].top=Inches(6.78)
    for i in (2,4):
        slide.shapes[i].width=Inches(.18);slide.shapes[i].height=Inches(.14)
    for i,w in [(3,1.2),(5,1.95)]:
        slide.shapes[i].width=Inches(w)
        slide.shapes[i].height=Inches(.18)
        set_text(slide.shapes[i], 'Software vulnerability' if i==3 else 'Platform or provider compromise', 7.7,RGBColor(91,91,96))
    # Single-line heading, using template purple.
    title=slide.shapes[56]
    set_text(title,'2. Latest Cyber Threat Intelligence – The New Vulnerability Norm',23.2,TITLE_PURPLE,True)
    title.left=Inches(.80);title.top=Inches(.08);title.width=Inches(11.9);title.height=Inches(.58)
    # Two visible oval/rounded note containers.
    rounded_note(slide,.95,1.04,11.3,.53,TOP,10.9)
    rounded_note(slide,.95,5.88,11.3,.53,BOTTOM,10.5)
    try:
        slide.notes_slide.notes_text_frame.text=NOTES
    except Exception as e:
        print('WARNING: Could not set notes:',e)
    prs.save(OUTPUT)
    print('SUCCESS: Created',OUTPUT)

if __name__=='__main__':
    main()
