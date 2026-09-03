#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_fonts.py — 字体可用性检查（DSH 本地导出用）

读 .pptd 声明的字体（或直接给字体名），按「规范名→实装名」映射翻译，
对照本机已装字体（注册表），输出：
  已装 / 缺失 / 本地可用替代。

用法:
  python check_fonts.py <deck.pptd>            # 从 .pptd 的 theme.textStyles 收集字体
  python check_fonts.py --fonts "MiSans" "Georgia"
  python check_fonts.py <deck.pptd> --list-missing   # 只列出缺失（供用户安装）
"""
import os, sys, re, argparse, json
import yaml

# 规范名 -> 实装名（与 reference/dsh-fonts.md 保持一致）
CANON_TO_INSTALLED = {
    "MiSans": "MiSans",
    "Noto Sans SC": "Noto Sans SC",
    "思源宋体": "思源宋体 CN",
    "Source Han Serif": "思源宋体 CN",
    "阿里妈妈刀隶体": "阿里妈妈刀隶体",
    "阿里妈妈东方大楷": "阿里妈妈东方大楷",
    "阿里妈妈数黑体": "阿里妈妈数黑体",
    "站酷文艺体": "站酷文艺体",
    "ZCOOL KuaiLe": "ZCOOL KuaiLe",
    "得意黑": "得意黑 斜体",
    "Smiley Sans": "得意黑 斜体",
    "飞波正点体": "飞波正点体",
    "霞鹜新致宋": "霞鹜新致宋＋",
    "LXGW Bright": "霞鹜文楷",
    "霞鹜文楷": "霞鹜文楷",
    "精品点阵体": "精品点阵体9×9 1.93 R",
    "Liter": "Liter",
    "HedvigLettersSans": "Hedvig Letters Sans",
    "Oranienbaum": "Oranienbaum",
    "QuattrocentoSans": "Quattrocento Sans",
    "Unna": "Unna",
    "Coda": "Coda",
    "Jersey15": "Jersey 15",
    "Jersey20Charted": "Jersey 20 Charted",
    "SortsMillGoudy": "Sorts Mill Goudy",
    "更纱黑体 SC": "更纱黑体 SC",
    "Sarasa Gothic": "更纱黑体 SC",
    "Microsoft YaHei": "Microsoft YaHei",
    "微软雅黑": "Microsoft YaHei",
    "SimHei": "SimHei",
    "黑体": "SimHei",
    "SimSun": "SimSun",
    "宋体": "SimSun",
    "FangSong": "FangSong",
    "仿宋": "FangSong",
    "KaiTi": "KaiTi",
    "楷体": "KaiTi",
}

# 缺失时的本地替代建议
ALTERNATIVES = {
    "SortsMillGoudy": ["Georgia", "Cambria"],
    "MiSans": ["更纱黑体 SC", "思源黑体 CN"],
    "Noto Sans SC": ["更纱黑体 SC", "思源黑体 CN"],
    "QuattrocentoSans": ["Century Gothic"],
    "HedvigLettersSans": ["Arial"],
    "Liter": ["Century Gothic", "Arial"],
    "Oranienbaum": ["Georgia", "Cambria"],
    "Unna": ["Cambria"],
    "Coda": ["Arial"],
    "Jersey15": ["Impact", "Franklin Gothic"],
    "Jersey20Charted": ["Impact", "Franklin Gothic"],
    "得意黑": ["思源黑体 CN", "更纱黑体 SC"],
    "精品点阵体": ["SimSun", "SimHei"],
}

REG_RE = r"^(\S[^\(]*?)\s*\("

def installed_families():
    """从注册表读本机已装字体族名集合（已归一化，去掉 (TrueType) 与后缀）。"""
    import winreg
    key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    fams = []
    try:
        k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key)
        i = 0
        while True:
            try:
                name, _val, _t = winreg.EnumValue(k, i)
            except OSError:
                break
            norm = re.sub(r"\s*\([^)]*\)\s*$", "", name).strip()
            if norm:
                fams.append(norm.lower())
            i += 1
        winreg.CloseKey(k)
    except Exception as e:
        print(f"[warn] 读取字体注册表失败: {e}", file=sys.stderr)
    return set(fams)

def normalize_req(name):
    # 归一化规范名，方便与注册表条目「包含/起始」匹配
    return re.sub(r"\s*\([^)]*\)\s*$", "", name).strip().lower()

def present(fam_set, req_installed):
    """判断实装名是否出现在本机。用「起始包含」容忍字重/变体后缀。"""
    r = normalize_req(req_installed)
    if not r:
        return False
    for f in fam_set:
        if f == r:
            return True
        if f.startswith(r + " ") or f.startswith(r + "\u3000"):
            return True
        if r in f and len(r) >= 2:
            # 对中文/特殊名更宽松：整词出现在条目开头
            if f.startswith(r):
                return True
    return False

def collect_fonts(manifest):
    m = yaml.safe_load(open(manifest, encoding="utf-8"))
    theme = m.get("theme", {}) or {}
    fonts = set()
    for _k, st in (theme.get("textStyles", {}) or {}).items():
        ff = st.get("fontFamily")
        if isinstance(ff, dict):
            fonts.add(ff.get("latin")); fonts.add(ff.get("ea"))
        elif ff:
            fonts.add(ff)
    # 也扫页面内联 fontFamily
    base = os.path.dirname(os.path.abspath(manifest))
    for rel in m.get("pages", []):
        try:
            page = yaml.safe_load(open(os.path.join(base, rel), encoding="utf-8"))
        except Exception:
            continue
        for el in page.get("elements", []):
            c = el.get("content", {})
            ff = c.get("fontFamily")
            if isinstance(ff, dict):
                fonts.add(ff.get("latin")); fonts.add(ff.get("ea"))
            elif ff and isinstance(ff, str):
                fonts.add(ff)
    return {f for f in fonts if f}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest", nargs="?")
    ap.add_argument("--fonts", nargs="*", default=[])
    ap.add_argument("--list-missing", action="store_true")
    args = ap.parse_args()

    req = set(args.fonts or [])
    if args.manifest:
        req |= collect_fonts(args.manifest)
    req = {f for f in req if f}

    fams = installed_families()
    installed, missing = [], []
    for c in sorted(req):
        inst = CANON_TO_INSTALLED.get(c, c)
        if present(fams, inst):
            installed.append((c, inst))
        else:
            missing.append(c)

    if args.list_missing:
        for c in missing:
            print(c)
        return 0

    print("=== 本机已装字体族数 ===", len(fams))
    print("\n[已装]")
    for c, inst in installed:
        print(f"  {c!r:28} -> {inst!r}")
    print("\n[缺失]")
    if missing:
        for c in missing:
            sug = ALTERNATIVES.get(c)
            print(f"  {c!r:28} 建议安装 或 用替代 {sug}")
    else:
        print("  (无)")
    print("\n[本地可用替代建议]")
    for c in missing:
        sug = ALTERNATIVES.get(c)
        if sug:
            print(f"  {c!r} -> {sug}")
    if not missing:
        print("  (所需字体本机均已具备)")

if __name__ == "__main__":
    main()
