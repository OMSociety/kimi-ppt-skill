#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptd_to_png.py — 纯本地 .pptd -> 页面预览图 渲染器（Pillow，无浏览器/无 LibreOffice/无外网）
用于受限环境内的视觉 QA：按 .pptd 每页几何/配色/文字/字体画成一页 PNG，并合成一张 overview.jpg。

用法:
  python pptd_to_png.py <deck.pptd> [-o <outdir>] [--scale 2]
  输出: <outdir>/page_1.png ... /overview.jpg
"""
import os, sys, re, argparse, html
import yaml
from PIL import Image, ImageDraw, ImageFont

# 系统字体目录（Windows 固定；POSIX 走 fontconfig 常用路径，逐个探测）
FONT_DIRS = [
    r"C:\Windows\Fonts",
    "/usr/share/fonts",
    "/usr/local/share/fonts",
    os.path.expanduser("~/.fonts"),
    "/Library/Fonts",
    "/System/Library/Fonts",
    os.path.expanduser("~/Library/Fonts"),
]
FONT_EXTS = (".ttf", ".ttc", ".otf", ".otc")

_SYSTEM_FONTS = None

# 规范名 -> 实装名（与 reference/local-fonts.md 一致；本地导出按实装名找文件）
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


def _styles(path):
    """文件名里的字重/字形：(weight, italic)。weight 越大越粗。"""
    stem = os.path.splitext(os.path.basename(path))[0].lower()
    tokens = set(re.split(r"[^a-z0-9]+", stem))
    if tokens & {"thin", "hairline"}:
        w = 100
    elif tokens & {"extralight", "ultralight"}:
        w = 200
    elif tokens & {"light"}:
        w = 300
    elif tokens & {"medium"}:
        w = 500
    elif tokens & {"semibold", "demibold", "demi"}:
        w = 600
    elif tokens & {"extrabold", "ultrabold"}:
        w = 800
    elif tokens & {"black", "heavy"}:
        w = 900
    elif tokens & {"bold", "bd"}:
        w = 700
    else:
        w = 400
    italic = bool(tokens & {"italic", "oblique", "it"}) or bool(tokens & {"bi", "bdit"})
    return w, italic


def _scan_fonts():
    """目录扫描兜底：[(family, path, weight, italic)]，族名由文件名推断。"""
    entries = []
    for d in FONT_DIRS:
        if not os.path.isdir(d):
            continue
        for root, _dirs, files in os.walk(d, followlinks=True):
            for fn in files:
                if fn.lower().endswith(FONT_EXTS):
                    path = os.path.join(root, fn)
                    entries.append((os.path.splitext(fn)[0].strip().lower(), path, *_styles(path)))
    return entries


# 注册表族名里可辨认的字重/字形词；其余情况一律回退用文件名判断
_WEIGHT_WORDS = {
    "thin": 100, "hairline": 100, "extralight": 200, "ultralight": 200, "light": 300,
    "book": 350, "regular": 400, "normal": 400, "medium": 500, "demibold": 600,
    "semibold": 600, "demi": 600, "bold": 700, "extrabold": 800, "ultrabold": 800,
    "black": 900, "heavy": 900,
}
# 「X Bold & X UI Bold」这类合并条目：族名只取 & 前那段
_FAMILY_SPLIT = " & "
# 尾部字重/字形词不属于族名
_TAIL_STYLE = re.compile(
    r"\s+(thin|hairline|extralight|ultralight|light|book|regular|normal|medium|"
    r"demibold|semibold|demi|bold|extrabold|ultrabold|black|heavy|italic|oblique)$")


def _style_from_name(raw):
    """(family, weight, italic)：解析注册表条目名，识别不出字重返回 None 交给文件名兜底。"""
    name = raw.strip().lower()
    italic = bool(re.search(r"\b(italic|oblique)\b", name))
    if _FAMILY_SPLIT in name:          # 「X Bold & X UI Bold」→ 只取 & 前那段
        name = name.split(_FAMILY_SPLIT, 1)[0]
    weight = None
    fam = name
    m = _TAIL_STYLE.search(name)       # 尾部字重/字形词不属于族名
    if m:
        fam = name[:m.start()].strip()
        if m.group(1) in _WEIGHT_WORDS:
            weight = _WEIGHT_WORDS[m.group(1)]
        if m.group(1) in ("italic", "oblique"):
            italic = italic or True
    return fam.strip(), weight, italic


def system_fonts():
    """系统可用字体：[(family, path, weight, italic)]。

    Windows 先读字体注册表拿真实族名与字重（"微软雅黑 Bold & Microsoft YaHei UI Bold"
    → 族 `微软雅黑`、weight 700）；注册表读不到时退到目录扫描（族名由文件名推断）。
    同族同时保留常规与粗体，供调用方按需挑选。
    """
    global _SYSTEM_FONTS
    if _SYSTEM_FONTS is None:
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
                raw = re.sub(r"\s*\([^)]*\)\s*$", "", str(name)).strip().lower()
                fn = os.path.basename(str(val))
                path = next((os.path.join(d, fn) for d in FONT_DIRS
                             if fn and os.path.isfile(os.path.join(d, fn))), None)
                if raw and path:
                    fam, weight, italic = _style_from_name(raw)
                    file_weight, file_italic = _styles(fn)
                    if weight is None:
                        weight = file_weight
                    italic = italic or file_italic
                    entries.append((fam, path, weight, italic))
                i += 1
            winreg.CloseKey(k)
        except Exception:
            entries = []
        if not entries:
            entries = _scan_fonts()
        _SYSTEM_FONTS = _dedupe_fonts(entries)
    return _SYSTEM_FONTS


def _dedupe_fonts(entries):
    """同族同字重只留一条：ASCII 族名优先（避免 CJK/乱码变体抢位），文件名短者优先。"""
    best = {}
    for fam, path, weight, italic in entries:
        key = (fam, italic, weight)
        prev = best.get(key)
        if prev is None or (len(os.path.basename(path)), fam.isascii()) < \
                (len(os.path.basename(prev)), fam.isascii()):
            best[key] = path
    return sorted((fam, path, weight, italic) for (fam, italic, weight), path in best.items())


def _family_rank(fam, prefs):
    """族名偏好序（分组：匹配到的 pref 序号 → 是否 ASCII → 族名）；无匹配返回 (-1,) 表示不候选。"""
    for i, pref in enumerate(prefs):
        if pref in fam:
            return i, 0 if fam.isascii() else 1, fam
    return (-1,)


def default_font_file(bold=False):
    """兜底字体：优先常见 CJK 族（同族内按需挑粗/常规字形），再退到任意一个系统字体。

    偏好表刻意只用 ASCII 名（`yahei`/`simsun`…）：中文 Windows 的注册表常常同时给出
    ASCII 与本地化两套族名，ASCII 那套才能在跨区域机器上稳定命中。
    """
    entries = system_fonts()
    if not entries:
        return ""
    prefs = ("yahei", "noto sans cjk", "noto sans sc", "source han", "simhei",
             "simsun", "pingfang", "hiragino", "wenquanyi", "dejavu sans", "arial")
    fams = {f for f, _p, _w, _i in entries}
    ranked = sorted((r, f) for r, f in ((_family_rank(f, prefs), f) for f in fams)
                    if r[0] >= 0)
    fam = ranked[0][1] if ranked else sorted(fams)[0]
    cands = [(p, w, i) for f, p, w, i in entries if f == fam]
    plain = [c for c in cands if not c[2]] or cands     # 优先非斜体
    target = 700 if bold else 400
    plain.sort(key=lambda c: abs(c[1] - target))
    return plain[0][0]


def font_file_for(name, bold):
    """按(规范名)解析本机字体文件；找不到回退系统兜底字体。bold 优先 bold 字形。"""
    base_name = CANON_TO_INSTALLED.get((name or "").strip(), (name or "").strip())
    target = base_name.lower()
    entries = system_fonts()
    cands = [(fam, p, w, i) for fam, p, w, i in entries
             if fam == target or fam.startswith(target + " ") or fam.startswith(target + "&")]
    if not cands:
        cands = [(fam, p, w, i) for fam, p, w, i in entries if target and target in fam]
    if not cands:
        return default_font_file(bold)
    # 斜体最次、族名完全相等优先、字重按需（bold 取 700，否则取 400）
    want = 700 if bold else 400
    cands.sort(key=lambda c: (c[3], c[0] != target, abs(c[2] - want)))
    return cands[0][1]
    return cands[0][1]


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
            font = ImageFont.truetype(default_font_file(bold), int(font_size))
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
