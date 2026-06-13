from docx import Document

doc = Document(r"c:\Users\sindhu.sundara\Downloads\Sindhu Sundaramoorthy_Resume.docx")
body = doc.element.body

W_NS   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS  = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WP_NS  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
A_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"

FALLBACK_TAG = "{%s}Fallback" % MC_NS

def has_fallback_ancestor(el):
    for anc in el.iterancestors():
        if anc.tag == FALLBACK_TAG:
            return True
    return False

txbx_all = body.findall(".//{%s}txbxContent" % W_NS)
primary = [t for t in txbx_all if not has_fallback_ancestor(t)]

def get_wsp_pos(txbx_el):
    # txbxContent -> wps:txbx -> wps:wsp
    txbx_parent = txbx_el.getparent()   # wps:txbx
    wsp = txbx_parent.getparent() if txbx_parent is not None else None  # wps:wsp
    if wsp is None or not wsp.tag.endswith("}wsp"):
        return (0, 0)
    # wsp -> wps:spPr -> a:xfrm -> a:off
    spPr = wsp.find("{%s}spPr" % WPS_NS)
    if spPr is None:
        return (0, 0)
    xfrm = spPr.find("{%s}xfrm" % A_NS)
    if xfrm is None:
        return (0, 0)
    off = xfrm.find("{%s}off" % A_NS)
    if off is None:
        return (0, 0)
    return (int(off.get("y", 0)), int(off.get("x", 0)))

def box_text(txbx):
    paras = txbx.findall("{%s}p" % W_NS)
    lines = []
    for p in paras:
        t = "".join(r.text or "" for r in p.findall(".//{%s}t" % W_NS))
        if t.strip():
            lines.append(t.strip())
    return "\n".join(lines)

boxes = []
for txbx in primary:
    pos = get_wsp_pos(txbx)
    txt = box_text(txbx)
    boxes.append((pos, txt))

boxes.sort()
print("=== BOXES SORTED BY GROUP POSITION (y, x) ===")
for (y, x), txt in boxes:
    print("y=%8d  x=%8d  |  %r" % (y, x, txt[:90]))
