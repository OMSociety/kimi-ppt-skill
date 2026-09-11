# 本地导出字体（以运行时实装为准）

> 本文件是**本地导出**（`pptd_to_pptx.py` / `pptd_to_png.py` / `check_fonts.py`）的字体基准。
> `fonts.md` 里的规范字体名是「设计意图/浏览器导出」用的；**实装名才是本地导出能真正命中的**，且**以当前机器的实装为准**。
> 规则：**规范名 ≠ 实装名时，一律以实装名为准**，生成 `.pptd` 时把 `fontFamily` 写成下表的「实际安装名」，或让导出器按下表映射。
> 若字体未装，按 `SKILL.md`「三、字体核对（防豆腐块）」处理（本文件只给映射与用途，处理动作不在此重复）。
>
> **下表是某台基准机的实测快照，不是通用事实**：换机器后先跑 `check_fonts.py` 拿到该机的「已装 / 缺失 / 可用替代」，**以脚本输出为准**；确需长期复用时再把该机结果回写本表。

## 规范名 → 实际安装名 映射（基准机快照）

### 中文
| 规范名（fonts.md/设计用） | 典型实装名 | 免费商用 | 说明 |
|---|---|---|---|
| MiSans | `MiSans` | ✅ 小米官方全球免费商用 | 默认中文首选；多字重建议安装 |
| Noto Sans SC | `Noto Sans SC` | ✅ | 别名：`思源黑体 CN`（同为 Source Han Sans） |
| Source Han Serif / 思源宋体 | `思源宋体 CN` | ✅ | 别名：`Noto Serif SC` |
| 阿里妈妈刀隶体 | `阿里妈妈刀隶体` | ✅ 阿里官方免费商用 | |
| 阿里妈妈东方大楷 | `阿里妈妈东方大楷` | ✅ | |
| 阿里妈妈数黑体 | `阿里妈妈数黑体` | ✅ | 常见为 Bold 字重 |
| 站酷文艺体 | `站酷文艺体` | ✅ ZCOOL | 同类备选 `站酷高端黑`、`站酷快乐体2016修订版` |
| 得意黑 (Smiley Sans) | `得意黑 斜体` | ✅ | 官方预设仅斜体字形（符合其设计） |
| 飞波正点体 | `飞波正点体` | ✅（免费商用） | |
| 霞鹜新致宋 | `霞鹜新致宋＋` | ✅ LXGW | 注：名字带「＋」 |
| LXGW Bright | `霞鹜文楷` | ✅ LXGW | 霞鹜文楷家族 |
| ZCOOL KuaiLe | `ZCOOL KuaiLe` | ✅ ZCOOL | |
| 精品点阵体 | `精品点阵体9×9…`（1.93 R/B、方格/光晕/渐层/立体/圆形版、港版） | ✅ | 用 `精品点阵体9×9 1.93 R` 作默认 |

### 英文
| 规范名 | 典型实装名 | 免费商用 | 说明 |
|---|---|---|---|
| Liter | `Liter` | ✅ OFL | |
| HedvigLettersSans | `Hedvig Letters Sans` | ✅ | 注意空格 |
| Oranienbaum | `Oranienbaum` | ✅ OFL | |
| QuattrocentoSans | `Quattrocento Sans` | ✅ OFL | 注意空格 |
| Unna | `Unna` | ✅ OFL | |
| Coda | `Coda` | ✅ OFL | |
| Jersey15 | `Jersey 15` | ✅ OFL | |
| Jersey20Charted | `Jersey 20 Charted` | ✅ OFL | |
| SortsMillGoudy | `Sorts Mill Goudy`（注意空格） | ✅ OFL | （SourcesMillGoudy-Regular/Italic.ttf） |

### 本地常见补充（系统预装 / 免费商用）
| 字体 | 用途建议 |
|---|---|
| `Century Gothic` / `Bahnschrift` | **几何感**，适合 Bauhaus/设计标题（比 Arial 更对味） |
| `Franklin Gothic` / `Impact` / `Rockwell` | 展示/强调标题 |
| `Georgia` / `Cambria` / `Constantia` | 学术衬线正文 |
| `Microsoft YaHei`（微软雅黑） | 默认中文回退 |
| `SimHei` `SimSun` `FangSong` `KaiTi` | 系统标准中文 |
| `更纱黑体 SC`（=Sarasa Gothic J/K） | MiSans 类似，备选中文黑体 |

## 检查命令（脚本 `scripts/check_fonts.py`）

按「规范名 → 实装名」映射翻译后，对照本机已装字体，输出 `已装 / 缺失 / 本地可用替代`。**检查动作与「缺失→改字体」的处理步骤见 `SKILL.md`「三、字体核对（防豆腐块）」，此处不重复。**

```bash
python scripts/check_fonts.py <deck.pptd>                  # 从 .pptd 的 theme.textStyles + 页面内联 fontFamily 收集字体
python scripts/check_fonts.py --fonts "MiSans" "Georgia"   # 直接给字体名检查
python scripts/check_fonts.py <deck.pptd> --list-missing   # 只列出缺失（供用户安装）
```
