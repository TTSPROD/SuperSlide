"""Extract the original Severstal icon library, without redrawing its artwork.

Original SVG media are copied byte for byte. Native DrawingML paths are converted
to SVG using their original coordinates, transforms and fills. Several slides
contain separate background and foreground shapes; intersecting shapes are kept
together as one icon. Wide groups containing multiple icons are unpacked first.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import math
import posixpath
import re
import io
from collections import Counter, defaultdict
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "SeverstalBrand/Template/Severstal_Icons.pptx"
OUT = ROOT / "assets/icons/severstal"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
SVG = "http://www.w3.org/2000/svg"
EMU = 914400
GROUP_IDS = ["metallurgy", "mining", "steel_service_centres", "construction",
             "mechanical_engineering", "energy", "products", "products",
             "product_properties", "services", "transport", "occupational_safety",
             "occupational_safety", "prohibitions", "logistics", "innovation",
             "information_technology", "project_teamwork", "marketing_sales", "finance",
             "staff_training", "teams", "internet", "information_security", "ecology",
             "law_order", "navigation", "navigation", "sport", "corporate_social_responsibility",
             "miscellaneous", "holidays", "emotions"]


def jwrite(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def svg_el(name, **attrs):
    return E.Element(f"{{{SVG}}}{name}", {k.replace('_', '-'): str(v) for k, v in attrs.items()})


def xfrm(node):
    return node.find("p:grpSpPr/a:xfrm", NS) if node.tag.endswith("grpSp") else node.find("p:spPr/a:xfrm", NS)


def box(node):
    t = xfrm(node)
    if t is None:
        return None
    o, ex = t.find("a:off", NS), t.find("a:ext", NS)
    return tuple(int(v) for v in (o.get("x"), o.get("y"), ex.get("cx"), ex.get("cy")))


def mult(a, b):
    # scale+translation only; library groups have no rotations or flips.
    sx, sy, tx, ty = a
    bx, by, ux, uy = b
    return sx*bx, sy*by, sx*ux+tx, sy*uy+ty


def group_matrix(node):
    t = xfrm(node)
    o, ex, co, ce = [t.find("a:" + s, NS) for s in ("off", "ext", "chOff", "chExt")]
    sx, sy = int(ex.get("cx"))/int(ce.get("cx")), int(ex.get("cy"))/int(ce.get("cy"))
    return sx, sy, int(o.get("x"))-sx*int(co.get("x")), int(o.get("y"))-sy*int(co.get("y"))


def global_box(b, matrix):
    x, y, w, h = b
    sx, sy, tx, ty = matrix
    return x*sx+tx, y*sy+ty, w*sx, h*sy


def candidates(tree, matrix=(1, 1, 0, 0)):
    for node in tree:
        if E.QName(node).localname not in {"sp", "grpSp", "pic"}:
            continue
        if node.xpath(".//a:t[string-length(normalize-space(.)) > 0]", namespaces=NS):
            continue
        b = box(node)
        if b is None:
            continue
        gb = global_box(b, matrix)
        if gb[1] > 7*EMU:  # empty footer placeholder
            continue
        if node.tag.endswith("grpSp") and (gb[2] > 1.2*EMU or gb[3] > 1.2*EMU):
            yield from candidates(node, mult(matrix, group_matrix(node)))
        elif gb[2] > 0 and gb[3] > 0:
            yield {"node": node, "box": gb, "matrix": matrix}


def intersects(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return min(ax+aw, bx+bw)-max(ax, bx) > min(aw,bw)*.1 and min(ay+ah,by+bh)-max(ay,by) > min(ah,bh)*.1


def cluster(items):
    parents = list(range(len(items)))
    def find(i):
        while parents[i] != i:
            parents[i] = parents[parents[i]]
            i = parents[i]
        return i
    for i, a in enumerate(items):
        for j in range(i):
            if intersects(a["box"], items[j]["box"]):
                parents[find(i)] = find(j)
    result = defaultdict(list)
    for i, item in enumerate(items):
        result[find(i)].append(item)
    def sort_key(c):
        b = bounds(c)
        return round((b[1]+b[3]/2)/EMU/.35), b[0]+b[2]/2
    return sorted(result.values(), key=sort_key)


def bounds(items):
    x = min(c["box"][0] for c in items)
    y = min(c["box"][1] for c in items)
    r = max(c["box"][0]+c["box"][2] for c in items)
    b = max(c["box"][1]+c["box"][3] for c in items)
    return x, y, r-x, b-y


def fill(props):
    if props.find("a:noFill", NS) is not None:
        return "none"
    c = props.find("a:solidFill", NS)
    if c is None:
        return "none"
    clr = c[0]
    v = clr.get("val")
    if E.QName(clr).localname == "schemeClr":
        v = {"bg1":"FFFFFF", "tx1":"000000", "accent1":"002F6C"}.get(v, v)
    if not re.fullmatch("[0-9a-fA-F]{6}", v or ""):
        raise ValueError(f"Unknown fill {v}")
    return "#" + v


def transform_attr(t, b):
    x, y, w, h = b
    rotation = int(t.get("rot", 0))/60000
    flipx, flipy = t.get("flipH") == "1", t.get("flipV") == "1"
    transform = ""
    if rotation or flipx or flipy:
        transform = f"translate({x+w/2} {y+h/2}) rotate({rotation}) scale({-1 if flipx else 1} {-1 if flipy else 1}) translate({-x-w/2} {-y-h/2})"
    return transform


def native_svg(node, media):
    b = box(node)
    if node.tag.endswith("grpSp"):
        sx, sy, tx, ty = group_matrix(node)
        g = svg_el("g", transform=f"translate({tx} {ty}) scale({sx} {sy})")
        for child in node:
            if E.QName(child).localname in {"sp", "grpSp", "pic"}:
                g.append(native_svg(child, media))
        return g
    props = node.find("p:spPr", NS)
    t = xfrm(node)
    outer = svg_el("g")
    if transform_attr(t,b):
        outer.set("transform", transform_attr(t,b))
    x,y,w,h = b
    if node.tag.endswith("pic"):
        embeds = node.xpath('.//*[local-name()="svgBlip"]/@r:embed', namespaces=NS)
        if not embeds:
            embeds = node.xpath('./p:blipFill/a:blip/@r:embed', namespaces=NS)
        content, ext = media[embeds[0]]
        mime = {".svg":"image/svg+xml", ".png":"image/png", ".jpeg":"image/jpeg", ".jpg":"image/jpeg"}[ext]
        outer.append(svg_el("image", x=x, y=y, width=w,height=h,
                            href=f"data:{mime};base64,"+base64.b64encode(content).decode(),preserveAspectRatio="none"))
        return outer
    geom = props.find("a:custGeom", NS)
    color = fill(props)
    if geom is None:
        preset = props.find("a:prstGeom", NS)
        if preset is not None and preset.get("prst") == "rect":
            outer.append(svg_el("rect",x=x,y=y,width=w,height=h,fill=color))
            return outer
        raise ValueError("Unsupported native icon geometry")
    for path in geom.findall("a:pathLst/a:path",NS):
        coords=[]
        for cmd in path:
            kind=E.QName(cmd).localname
            points = [(float(pt.get("x")),float(pt.get("y"))) for pt in cmd]
            if kind == "close": coords.append("Z")
            elif kind in {"moveTo","lnTo","cubicBezTo"}:
                op={"moveTo":"M","lnTo":"L","cubicBezTo":"C"}[kind]
                coords.append(op + " ".join(f"{px:g},{py:g}" for px,py in points))
            else: raise ValueError(f"Unsupported path command {kind}")
        sx,sy=w/int(path.get("w")),h/int(path.get("h"))
        attr={"d":" ".join(coords), "fill": "none" if path.get("fill")=="none" else color,
              "transform":f"translate({x} {y}) scale({sx} {sy})"}
        # All visible native icon outlines are represented by filled paths.
        ln=props.find("a:ln",NS)
        if ln is not None and fill(ln) != "none" and path.get("stroke") != "0":
            attr.update(stroke=fill(ln), **{"stroke-width":str(int(ln.get("w",0))/max(sx,sy))})
        outer.append(svg_el("path",**attr))
    return outer


def slug(label):
    ru="абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
    en=["a","b","v","g","d","e","e","zh","z","i","i","k","l","m","n","o","p","r","s","t","u","f","kh","ts","ch","sh","shch","","y","","e","yu","ya"]
    t=label.lower().translate(str.maketrans(dict(zip(ru,en))))
    return re.sub('[^a-z0-9]+','_',t).strip('_')[:100]


def expand_names(entry):
    if isinstance(entry,list):
        return entry
    result=[name + " — " + style for style in entry["styles"] for name in entry["base"]]
    for index,label in entry.get("overrides",{}).items():
        result[int(index)-1]=label
    return result


def build(previews=None, allow_unnamed=False, render=False):
    names_path = ROOT / "scripts/severstal_icon_names.json"
    names=json.loads(names_path.read_text(encoding="utf-8")) if names_path.exists() else {}
    for part in sorted((ROOT/"scripts").glob("severstal_icon_names_part*.json")):
        names.update(json.loads(part.read_text(encoding="utf-8")))
    old_catalog=json.loads((OUT/"catalog.json").read_text(encoding="utf-8")) if (OUT/"catalog.json").exists() else {}
    used_paths=set()
    OUT.mkdir(parents=True,exist_ok=True)
    groups={}
    records=[]
    inventory=[]
    with ZipFile(SOURCE) as z:
        presentation=E.fromstring(z.read("ppt/presentation.xml"))
        size=presentation.find("p:sldSz",NS)
        slidew,slideh=int(size.get("cx")),int(size.get("cy"))
        for num,gid in enumerate(GROUP_IDS,1):
            root=E.fromstring(z.read(f"ppt/slides/slide{num}.xml"))
            title=next(t.strip() for t in root.xpath("//a:t/text()",namespaces=NS) if not t.strip().isnumeric())
            groups.setdefault(gid,{"id":gid,"name":title,"source_slides":[],"icon_ids":[]})["source_slides"].append(num)
            rels=E.fromstring(z.read(f"ppt/slides/_rels/slide{num}.xml.rels"))
            media={}
            for rel in rels:
                if rel.get("Type").endswith("/image"):
                    part=posixpath.normpath(posixpath.join("ppt/slides",rel.get("Target")))
                    media[rel.get("Id")]=z.read(part),Path(part).suffix
            cells=cluster(list(candidates(root.find("p:cSld/p:spTree",NS))))
            sheet_names=expand_names(names.get(str(num),[]))
            if len(sheet_names)!=len(cells) and not allow_unnamed:
                raise ValueError(f"Slide {num}: {len(cells)} icons but {len(sheet_names)} names")
            image_path=Path(previews)/f"slide-{num:02}.png" if previews else None
            source_image=Image.open(image_path).convert("RGB") if image_path and image_path.exists() else None
            thumbs=[]
            for index,items in enumerate(cells,1):
                label=sheet_names[index-1] if len(sheet_names)>=index else f"{title} {index:02}"
                if num==7 and index==35:
                    label="Красная пометка 7 из исходника"
                key=f"s{num:02}_{index:02}"
                fname=slug(label)
                aid=f"severstal_{gid}_{fname}"
                if any(r["id"]==aid for r in records):
                    fname += "_"+key
                    aid=f"severstal_{gid}_{fname}"
                folder=OUT/gid
                folder.mkdir(parents=True,exist_ok=True)
                asset=folder/(fname+".svg")
                bbox=bounds(items)
                # Original SVG files retain the complete, unmodified media bytes.
                node=items[0]["node"]
                svgids=node.xpath('.//*[local-name()="svgBlip"]/@r:embed',namespaces=NS)
                if len(items)==1 and node.tag.endswith("pic") and svgids:
                    content=media[svgids[0]][0]
                    mode="original_svg"
                else:
                    x,y,w,h=bbox
                    svg=E.Element(f"{{{SVG}}}svg",nsmap={None:SVG},width="256",height="256",viewBox=f"{x:g} {y:g} {w:g} {h:g}")
                    for item in items:
                        sx,sy,tx,ty=item["matrix"]
                        group=svg_el("g",transform=f"translate({tx} {ty}) scale({sx} {sy})")
                        group.append(native_svg(item["node"],media))
                        svg.append(group)
                    content=E.tostring(svg,xml_declaration=True,encoding="UTF-8")
                    mode="drawingml_to_svg"
                asset.write_bytes(content)
                used_paths.add(asset.resolve())
                r={"id":aid,"asset_id":aid,"group_id":gid,"name":label,"label":label,"description":label+". Группа: "+title+".",
                   "tags":[title,*[x for x in re.split(r'[ /(),]+',label.lower()) if x]],
                   "path":asset.relative_to(ROOT).as_posix(),"file":asset.relative_to(ROOT).as_posix(),
                   "svg_file":asset.relative_to(ROOT).as_posix(),"format":"svg","extraction":mode,
                   "selectable":not (num==7 and index in (35,72)),
                   "is_vector":b"data:image/png" not in content and b"data:image/jpeg" not in content,
                   "sha256":hashlib.sha256(content).hexdigest(),
                   "source":{"file":SOURCE.relative_to(ROOT).as_posix(),"slide":num,"index":index,
                             "shape_ids":[int(item["node"].xpath('./*/p:cNvPr/@id',namespaces=NS)[0]) for item in items],
                             "bbox_emu":[round(v) for v in bbox]}}
                if source_image is not None:
                    x,y,w,h=bbox
                    crop=source_image.crop((max(0,int(x/slidew*source_image.width)),max(0,int(y/slideh*source_image.height)),
                                            min(source_image.width,math.ceil((x+w)/slidew*source_image.width)),
                                            min(source_image.height,math.ceil((y+h)/slideh*source_image.height))))
                    preview=folder/(fname+".png")
                    crop.save(preview)
                    used_paths.add(preview.resolve())
                    r["preview_path"]=preview.relative_to(ROOT).as_posix()
                    thumbs.append((index,crop,label))
                if render:
                    import resvg_py
                    png=asset.with_suffix(".png")
                    png.write_bytes(resvg_py.svg_to_bytes(svg_path=str(asset),width=256,height=256,skip_system_fonts=True))
                    used_paths.add(png.resolve())
                    r["png_file"]=png.relative_to(ROOT).as_posix()
                    r["preview_path"]=png.relative_to(ROOT).as_posix()
                    crop=Image.open(io.BytesIO(resvg_py.svg_to_bytes(svg_path=str(asset),width=100,height=100,background="#8194a8",skip_system_fonts=True))).convert("RGB")
                    if source_image is not None:
                        thumbs[-1]=(index,crop,label)
                    else:
                        thumbs.append((index,crop,label))
                records.append(r)
                groups[gid]["icon_ids"].append(aid)
            inventory.append({"slide":num,"group_id":gid,"name":title,"count":len(cells)})
            if thumbs:
                font=ImageFont.truetype("C:/Windows/Fonts/arial.ttf",15)
                cols=10; tw,th=125,130
                sheet=Image.new("RGB",(cols*tw,math.ceil(len(thumbs)/cols)*th),"#e0e7ef")
                draw=ImageDraw.Draw(sheet)
                for i,(index,crop,label) in enumerate(thumbs):
                    crop.thumbnail((100,100))
                    xx,yy=(i%cols)*tw,(i//cols)*th
                    sheet.paste(crop,(xx+(tw-crop.width)//2,yy+4))
                    draw.text((xx+4,yy+106),str(index),fill="black",font=font)
                dest=OUT/"contact_sheets"/f"{num:02}_{gid}.jpg"
                dest.parent.mkdir(exist_ok=True)
                sheet.save(dest,quality=90)
    catalog={"version":"1.0","brand":"Северсталь","source":SOURCE.relative_to(ROOT).as_posix(),
             "source_sha256":hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
             "selection_instructions":"Выбирайте только записи с selectable:true. Сначала определите исходную тематическую группу, затем asset_id по name, description и tags. Цветовые исполнения и фон указаны в названии; сохраняйте пропорции. is_vector:false означает исходный растр внутри SVG; используйте png_file для вставки PNG.",
             "groups":list(groups.values()),"icons":records,
             "stats":{"groups":len(groups),"source_slides":33,"icons":len(records),
                      "selectable_icons":sum(x["selectable"] for x in records),
                      "source_annotations":sum(not x["selectable"] for x in records),
                      "vector_icons":sum(x["is_vector"] for x in records),
                      "with_embedded_raster":sum(not x["is_vector"] for x in records),
                      "extraction":dict(Counter(x["extraction"] for x in records))}}
    jwrite(OUT/"catalog.json",catalog)
    jwrite(OUT/"source_inventory.json",inventory)
    # Only remove files explicitly owned by the previous generated catalog.
    for record in old_catalog.get("icons",[]):
        for field in ("path","preview_path","png_file"):
            old=ROOT/record.get(field,"")
            if old.is_file() and old.resolve().is_relative_to(OUT.resolve()) and old.resolve() not in used_paths:
                old.unlink()
    print(json.dumps(catalog["stats"],ensure_ascii=False))
    return catalog


if __name__=="__main__":
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--previews",type=Path,help="PowerPoint PNG exports, slide-01.png through slide-33.png")
    parser.add_argument("--allow-unnamed",action="store_true",help="Only for preparing annotation contact sheets")
    parser.add_argument("--render",action="store_true",help="Also render transparent PNG files (requires resvg_py)")
    args=parser.parse_args()
    build(args.previews,args.allow_unnamed,args.render)
