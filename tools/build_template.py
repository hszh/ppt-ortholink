"""
Generate build/gvml/OL_<topology>.gvml, one per connector topology.

Each file holds one "custom geometry" shape for a connector topology
(sequence of horizontal / vertical segments), packaged in the clipboard
format Office uses to move shapes between programs ("Art::GVML ClipFormat").
build.cmd embeds the files in the add-in; the add-in puts the right one on
the clipboard, pastes it into the user's slide, then only touches
position / size / flip / adjust values, never the geometry itself.

Adjust values (visible in the object model as Shape.Adjustments):
  adj1          corner radius in EMU (12700 EMU = 1 pt); no handle
  adj2 .. adjN  position of each *interior* segment as a fraction of the
                shape width (vertical segments) or height (horizontal
                segments); each has a yellow drag handle in PowerPoint

Local coordinate space of every shape: the path runs from (0,0) to (w,h).
PowerPoint's flipH / flipV take care of the other three quadrants.
Every corner is a quarter circle whose radius shrinks automatically when a
segment is too short.
"""
import io, os, sys, zipfile

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
LC = "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"
EMU_PER_PT = 12700
DEFAULT_RADIUS_EMU = 12 * EMU_PER_PT
SIZE_EMU = (100 * EMU_PER_PT, 80 * EMU_PER_PT)   # any size will do, the add-in resizes the shape
LINE_EMU = 19050     # 1.5 pt
LINE_RGB = "404040"
Q = 5400000          # 90 degrees in 60000ths of a degree
TOPOLOGIES = ["HV", "VH", "HVH", "VHV", "HVHV", "VHVH", "HVHVH", "VHVHV"]


def gd(name, fmla):
    return f'<a:gd name="{name}" fmla="{fmla}"/>'


def build(topo):
    """Return (guides, handles, path) for a segment sequence like 'HVHV'."""
    segs = list(topo)
    n = len(segs)
    guides = [gd("rad", "*/ adj1 1 1")]
    handles = []

    # position of each segment along its perpendicular axis (H: y, V: x)
    pos = []
    free = 0
    for k, s in enumerate(segs):
        if k == 0:
            pos.append("0")
        elif k == n - 1:
            pos.append("h" if s == "H" else "w")
        else:
            free += 1
            name = f"p{k}"
            guides.append(gd(name, f"*/ {'h' if s == 'H' else 'w'} adj{free + 1} 100000"))
            pos.append(name)

    # corner points C1..C(n-1): corner i sits between seg i-1 and seg i
    cx, cy = ["0"], ["0"]
    for i in range(1, n):
        a, b = segs[i - 1], segs[i]
        if a == "H":     # H then V: x from the V segment, y from the H segment
            cx.append(pos[i]); cy.append(pos[i - 1])
        else:
            cx.append(pos[i - 1]); cy.append(pos[i])
    cx.append("w"); cy.append("h")

    # segment deltas, directions, lengths
    lens = []
    dirs = []
    for k, s in enumerate(segs):
        if s == "H":
            guides.append(gd(f"d{k}", f"+- {cx[k + 1]} 0 {cx[k]}"))
        else:
            guides.append(gd(f"d{k}", f"+- {cy[k + 1]} 0 {cy[k]}"))
        guides.append(gd(f"len{k}", f"abs d{k}"))
        guides.append(gd(f"dir{k}", f"?: d{k} 1 -1"))
        lens.append(f"len{k}"); dirs.append(f"dir{k}")

    # radius: min(rad, terminal segment lengths, half of interior segment lengths)
    prev = "rad"
    for k in range(n):
        if 0 < k < n - 1:
            guides.append(gd(f"half{k}", f"*/ {lens[k]} 1 2"))
            lim = f"half{k}"
        else:
            lim = lens[k]
        guides.append(gd(f"rm{k}", f"min {prev} {lim}"))
        prev = f"rm{k}"
    guides.append(gd("rEff", f"max {prev} 1"))

    # path
    path = ['<a:moveTo><a:pt x="l" y="t"/></a:moveTo>']
    for i in range(1, n):
        din, dout = dirs[i - 1], dirs[i]
        guides.append(gd(f"off{i}", f"*/ rEff {din} 1"))
        if segs[i - 1] == "H":           # H -> V corner
            guides.append(gd(f"tx{i}", f"+- {cx[i]} 0 off{i}"))
            path.append(f'<a:lnTo><a:pt x="tx{i}" y="{cy[i]}"/></a:lnTo>')
            guides.append(gd(f"st{i}", f"?: {dout} 16200000 {Q}"))
            guides.append(gd(f"pr{i}", f"*/ {din} {dout} 1"))
            guides.append(gd(f"sw{i}", f"*/ {Q} pr{i} 1"))
        else:                            # V -> H corner
            guides.append(gd(f"ty{i}", f"+- {cy[i]} 0 off{i}"))
            path.append(f'<a:lnTo><a:pt x="{cx[i]}" y="ty{i}"/></a:lnTo>')
            guides.append(gd(f"st{i}", f"?: {dout} 10800000 0"))
            guides.append(gd(f"pr{i}", f"*/ {din} {dout} 1"))
            guides.append(gd(f"sw{i}", f"*/ -{Q} pr{i} 1"))
        path.append(f'<a:arcTo wR="rEff" hR="rEff" stAng="st{i}" swAng="sw{i}"/>')
    path.append('<a:lnTo><a:pt x="r" y="b"/></a:lnTo>')

    # handles: one per interior segment, sitting at the middle of that segment
    free = 0
    for k in range(1, n - 1):
        free += 1
        if segs[k] == "V":
            guides.append(gd(f"hm{k}", f"+/ {cy[k]} {cy[k + 1]} 2"))
            handles.append(f'<a:ahXY gdRefX="adj{free + 1}" minX="-2147483647" maxX="2147483647">'
                           f'<a:pos x="{pos[k]}" y="hm{k}"/></a:ahXY>')
        else:
            guides.append(gd(f"hm{k}", f"+/ {cx[k]} {cx[k + 1]} 2"))
            handles.append(f'<a:ahXY gdRefY="adj{free + 1}" minY="-2147483647" maxY="2147483647">'
                           f'<a:pos x="hm{k}" y="{pos[k]}"/></a:ahXY>')
    return guides, handles, "".join(path), free


def custgeom_xml(topo):
    guides, handles, path, free = build(topo)
    av = [gd("adj1", f"val {DEFAULT_RADIUS_EMU}")] + [gd(f"adj{i + 2}", "val 50000") for i in range(free)]
    return (
        '<a:custGeom>'
        f'<a:avLst>{"".join(av)}</a:avLst>'
        f'<a:gdLst>{"".join(guides)}</a:gdLst>'
        f'<a:ahLst>{"".join(handles)}</a:ahLst>'
        '<a:cxnLst/>'
        '<a:rect l="l" t="t" r="r" b="b"/>'
        f'<a:pathLst><a:path fill="none">{path}</a:path></a:pathLst>'
        '</a:custGeom>'
    )


def drawing_xml(topo):
    """The shape on a locked canvas, the way Office writes shapes to the clipboard."""
    cx, cy = SIZE_EMU
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{LC}"><lc:lockedCanvas xmlns:lc="{LC}">'
        '<a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr>'
        f'<a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></a:grpSpPr>'
        f'<a:sp><a:nvSpPr><a:cNvPr id="2" name="OL_{topo}"/><a:cNvSpPr/></a:nvSpPr>'
        f'<a:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>{custgeom_xml(topo)}'
        f'<a:noFill/><a:ln w="{LINE_EMU}"><a:solidFill><a:srgbClr val="{LINE_RGB}"/></a:solidFill><a:round/></a:ln>'
        '</a:spPr></a:sp>'
        '</lc:lockedCanvas></a:graphicData></a:graphic>'
    )


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/clipboard/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
    '</Types>'
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="clipboard/drawings/drawing1.xml"/>'
    '</Relationships>'
)


def gvml(topo):
    """Bytes of an "Art::GVML ClipFormat" package holding the connector shape for topo."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for name, text in (("[Content_Types].xml", CONTENT_TYPES), ("_rels/.rels", ROOT_RELS),
                           ("clipboard/drawings/drawing1.xml", drawing_xml(topo))):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))   # fixed time: same input, same bytes
            z.writestr(info, text.encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
    return buf.getvalue()


def main(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    for topo in TOPOLOGIES:
        with open(os.path.join(out_dir, "OL_%s.gvml" % topo), "wb") as f:
            f.write(gvml(topo))
    print("wrote %d connector templates to %s" % (len(TOPOLOGIES), out_dir))


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    default = os.path.join(here, "..", "build", "gvml")
    main(os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else default))
