#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptd_to_png.py — 纯本地 .pptd -> 页面预览图 渲染器（Pillow，无浏览器/无 LibreOffice/无外网）
用于 DSH 内的视觉 QA：按 .pptd 每页几何/配色/文字/字体画成一页 PNG，并合成一张 overview.jpg。

用法:
  python pptd_to_png.py <deck.pptd> [-o <outdir>] [--scale 2]
  输出: <outdir>/page_1.png ... /overview.jpg
"""
import os, sys, re, argparse, html
import yaml
from PIL import Image, ImageDraw, ImageFont

FONTDIR = r"C:\Windows\Fonts"

# 规范名 -> 实装名（与 reference/dsh-fonts.md 一致；本地导出按实装名找文件）
CANON_TO_INSTALLED = {
    "MiSans": "MiSans", "Noto Sans SC": "Noto Sans SC", "思源宋体": "思源宋体 CN",
    "Source Han Serif": "思源宋体 CN", "阿里妈妈刀隶体": "阿里妈妈刀隶体",
    "阿里妈妈东方大楷": "阿里妈妈东方大楷", "阿里妈妈数黑体": "阿里妈妈数黑体",
    "站酷文艺体": "站酷文艺体", "ZCOOL KuaiLe": "ZCOOL KuaiLe",
    "得意黑": "得意黑 斜体", "Smiley Sans": "得意黑 斜体", "飞波正点体": "飞波正点体",
    "霞鹜新致宋": "霞鹜新致宋＋", "LXGW Bright": "霞鹜文楷", "霞鹜文楷": "霞鹜文楷",
    "精品点阵体": "精品点阵体9×9 1.93 R", "Liter": "Liter",
    "HedvigLettersSans": "Hedvig Letters Sans", "Oranienbaum": "Oranienbaum",
    "QuattrocentoSans": "Quattrocento Sans", "Unna": "Unna", "Coda": "Coda",
    "Jersey15": "Jersey 15", "Jersey20Charted": "Jersey 20 Charted",
    "SortsMillGoudy": "Sorts Mill Goudy", "更纱黑体 SC": "更纱黑体 SC",
    "Sarasa Gothic": "更纱黑体 SC", "Microsoft YaHei": "Microsoft YaHei",
    "微软雅黑": "Microsoft YaHei", "SimHei": "SimHei", "黑体": "SimHei",
    "SimSun": "SimSun", "宋体": "SimSun", "FangSong": "FangSong", "仿宋": "FangSong",
    "KaiTi": "KaiTi", "楷体": "KaiTi", "思源黑体 CN": "思源黑体 CN",
    "Century Gothic": "Century Gothic", "Bahnschrift": "Bahnschrift",
    "Georgia": "Georgia", "Cambria": "Cambria", "Constantia": "Constantia",
}

_INSTALLED_CACHE = None

def _installed_font_list():
    """从注册表读本机字体：[(归一化显示名, 文件名)]，按族名起始匹配用。"""
    global _INSTALLED_CACHE
    if _INSTALLED_CACHE is not None:
        return _INSTALLED_CACHE
    entries = []
    try:
        import winreg
        k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                           r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        i = 0
        while True:
            try:
                name, val, _t = winreg.EnumValue(k, i)
            except OSError:
                break
            norm = re.sub(r"\s*\([^)]*\)\s*$", "", str(name)).strip().lower()
            fn = os.path.basename(str(val))
            if norm and fn and os.path.isfile(os.path.join(FONTDIR, fn)):
                entries.append((norm, fn))
            i += 1
        winreg.CloseKey(k)
    except Exception:
        pass
    _INSTALLED_CACHE = entries
    return entries

def font_file_for(name, bold):
    """按(规范名)解析本机字体文件；找不到回退微软雅黑。bold 优先 bold 字形。"""
    base_name = CANON_TO_INSTALLED.get((name or "").strip(), (name or "").strip())
    target = base_name.lower()
    entries = _installed_font_list()
    cands = [(norm, fn) for norm, fn in entries
             if norm == target or norm.startswith(target + " ") or norm.startswith(target + "&")]
    if not cands:
        cands = [(norm, fn) for norm, fn in entries if target in norm]
    if not cands:
        return os.path.join(FONTDIR, "msyhbd.ttc" if bold else "msyh.ttc")
    def score(c):
        norm = c[0]; s = 0
        if bold and any(k in norm for k in ("bold", "heavy", "black", "semibold", "demi")): s += 2
        if (not bold) and any(k in norm for k in ("regular", "light", "thin", "xlight")): s += 1
        if norm == target: s += 1
        if "italic" in norm: s -= 2
        if "cond" in norm or "condensed" in norm: s -= 1
        return s
    cands.sort(key=score, reverse=True)
    return os.path.join(FONTDIR, cands[0][1])

def resolve_color(c, theme_colors, fallback="000000"):
    if c is None:
        return "#" + fallback
    s = str(c).strip()
    if s.startswith("$"):
        key = s[1:]
        if key in theme_colors:
            return "#" + theme_colors[key].replace("#", "")
        return "#" + {"black": "111111", "white": "FFFFFF", "red": "E30613",
                      "yellow": "FFC300", "blue": "0066B3", "gray": "F4F4F2",
                      "text": "222222", "accent": "FFC300", "primary": "0066B3",
                      "secondary": "555555"}.get(key, fallback)
    return "#" + s.replace("#", "")

def plain_paragraphs(text):
    """把 content.text 转成纯文本段落列表，去掉 HTML 标签，保留换行。"""
    if not text:
        return []
    text = html.unescape(str(text))
    text = re.sub(r"</?p[^>]*>", "\n", text)
    text = re.sub(r"<br\s*/?>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    out = []
    for line in text.split("\n"):
        line = line.strip()
        if line:
            out.append(line)
    return out

def shape_pts(kind, x, y, w, h, border_ex, scale):
    """返回 PIL 多边形绘制所需的点（矩形/三角/菱形等）。"""
    if kind == "triangle":
        return [(x + w / 2, y), (x, y + h), (x + w, y + h)]
    if kind == "diamond":
        return [(x + w / 2, y), (x + w, y + h / 2), (x + w / 2, y + h), (x, y + h / 2)]
    if kind == "chevron":
        return [(x, y), (x + w * 0.7, y), (x + w, y + h / 2), (x + w * 0.7, y + h), (x, y + h), (x + w * 0.3, y + h / 2)]
    if kind == "star5":
        return None  # 复杂星形，退回矩形
    return None      # 退回矩形

def draw_text(draw, el, theme, scale, w, h, colors):
    c = el.get("content", {})
    x, y, bw, bh = [v * scale for v in el.get("bounds", [0, 0, 100, 40])]
    tc = theme.get("textStyles", {})
    base = {}
    st = c.get("style")
    if isinstance(st, str) and st.startswith("$"):
        base = tc.get(st[1:], {}) or {}
    font_size = (c.get("fontSize") or base.get("fontSize") or 18) * scale
    color = resolve_color(c.get("color") or base.get("color") or "$text", colors, "222222")
    bold = bool(c.get("bold", base.get("bold", False)))
    fname = c.get("fontFamily") or base.get("fontFamily") or "Microsoft YaHei"
    if isinstance(fname, dict):
        latin = fname.get("latin", "Arial"); ea = fname.get("ea", latin)
    else:
        latin = ea = fname
    lh = c.get("lineHeight", base.get("lineHeight", 1)) or 1
    align = c.get("align") or ["left", "top"]
    lines = plain_paragraphs(c.get("text", ""))
    # 有中文用 ea（须支持 CJK），纯西文用 latin；都按实装名解析
    joined = "".join(lines)
    has_cjk = any(ord(ch) > 0x2E80 for ch in joined)
    chosen = ea if has_cjk else latin
    font_path = font_file_for(chosen, bold)
    try:
        font = ImageFont.truetype(font_path, int(font_size))
    except Exception:
        try:
            font = ImageFont.truetype(os.path.join(FONTDIR, "msyhbd.ttc" if bold else "msyh.ttc"), int(font_size))
        except Exception:
            font = ImageFont.load_default()
    line_h = int(font_size * lh)
    # 总文本高
    total_h = line_h * max(1, len(lines))
    # 垂直定位
    if align[1] == "middle":
        ty = y + (bh - total_h) / 2
    elif align[1] == "bottom":
        ty = y + bh - total_h
    else:
        ty = y
    for i, line in enumerate(lines):
        lw = draw.textlength(line, font=font)
        if align[0] == "center":
            tx = x + (bw - lw) / 2
        elif align[0] == "right":
            tx = x + bw - lw
        else:
            tx = x
        draw.text((tx, ty + i * line_h), line, font=font, fill=color)
    return

def render_page(page, theme, colors, size, scale, page_no):
    w, h = size
    W, H = int(w * scale), int(h * scale)
    img = Image.new("RGB", (W, H), "#FFFFFF")
    drw = ImageDraw.Draw(img)
    # 背景
    bkg = page.get("background", {}) or {}
    if bkg.get("type", "solid") == "solid":
        bc = resolve_color(bkg.get("color"), colors, "FFFFFF")
        drw.rectangle([0, 0, W, H], fill=bc)
    elif bkg.get("type") == "gradient":
        stops = bkg.get("stops") or []
        if stops:
            gc = resolve_color(stops[0].get("color"), colors, "FFFFFF")
            drw.rectangle([0, 0, W, H], fill=gc)
    # 元素（按顺序 = 层序）
    for el in page.get("elements", []):
        et = el.get("elementType")
        x, y, bw, bh = [v * scale for v in el.get("bounds", [0, 0, 100, 40])]
        if et == "shape":
            sn = el.get("shapeName", "rect")
            fill = el.get("fill") or {}
            fc = resolve_color(fill.get("color"), colors)
            border = el.get("border") or {}
            bc = resolve_color(border.get("color"), colors, "000000")
            bwd = int(float(border.get("width", 0)) * scale)
            if sn in ("ellipse", "circle"):
                drw.ellipse([x, y, x + bw, y + bh], fill=fc,
                            outline=bc if border else None, width=bwd)
            elif sn == "roundRect":
                drw.rounded_rectangle([x, y, x + bw, y + bh], radius=int(min(bw, bh) * 0.1),
                                      fill=fc, outline=bc if border else None, width=bwd)
            elif sn in ("triangle", "diamond", "chevron"):
                pts = shape_pts(sn, x, y, bw, bh, border, scale)
                if pts:
                    drw.polygon(pts, fill=fc,
                                outline=bc if border else None)
                else:
                    drw.rectangle([x, y, x + bw, y + bh], fill=fc)
            else:
                drw.rectangle([x, y, x + bw, y + bh], fill=fc,
                              outline=bc if border else None, width=bwd)
        elif et == "text":
            draw_text(drw, el, theme, scale, w, h, colors)
        elif et == "line":
            pts = (el.get("points") or "0,0 1,1").split()
            p0 = [float(a) for a in pts[0].split(",")]
            p1 = [float(a) for a in pts[-1].split(",")]
            lc = resolve_color((el.get("border") or {}).get("color"), colors, "000000")
            drw.line([(x + p0[0] / 1, y + p0[1] / 1), (x + p1[0] / 1, y + p1[1] / 1)],
                     fill=lc, width=max(1, int(float((el.get("border") or {}).get("width", 1)) * scale)))
        elif et == "image":
            src = el.get("src")
            if src and os.path.isfile(os.path.join(os.path.dirname(manifest), src)):
                try:
                    im = Image.open(os.path.join(os.path.dirname(manifest), src)).convert("RGB")
                    im = im.resize((int(bw), int(bh)))
                    img.paste(im, (int(x), int(y)))
                except Exception:
                    pass
        # icon/table/chart：跳过（略）
    out = os.path.join(outdir, f"page_{page_no}.png")
    img.save(out)
    return out, img

def main():
    global manifest, outdir
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest")
    ap.add_argument("-o", "--output")
    ap.add_argument("--scale", type=int, default=2)
    args = ap.parse_args()
    manifest = os.path.abspath(args.manifest)
    base = os.path.dirname(manifest)
    outdir = args.output or os.path.join(base, ".preview")
    os.makedirs(outdir, exist_ok=True)
    m = yaml.safe_load(open(manifest, encoding="utf-8"))
    size = m.get("size", [960, 540])
    theme = m.get("theme", {}) or {}
    colors = theme.get("colors", {}) or {}
    imgs = []
    for i, rel in enumerate(m.get("pages", []), 1):
        page = yaml.safe_load(open(os.path.join(base, rel), encoding="utf-8"))
        out, img = render_page(page, theme, colors, size, args.scale, i)
        imgs.append((rel, img))
        print("rendered", out)
    # 合成 overview
    if imgs:
        gap = 20
        W = max(im.width for _, im in imgs)
        H = sum(im.height for _, im in imgs) + gap * (len(imgs) + 1)
        ov = Image.new("RGB", (W, H), "#DDDDDD")
        yoff = gap
        for _, im in imgs:
            ov.paste(im, (0, yoff)); yoff += im.height + gap
        ovp = os.path.join(outdir, "overview.jpg")
        ov.save(ovp, quality=92)
        print("overview", ovp)

if __name__ == "__main__":
    main()
