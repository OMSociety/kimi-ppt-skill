# DSH 本地导出字体（以本机实装为准）

> 本文件是 **DSH 本地导出**（`pptd_to_pptx.py` / `pptd_to_png.py` / `check_fonts.py`）的字体基准。
> `fonts.md` 里的规范字体名是「设计意图/浏览器导出」用的；**本机实装名才是本地导出能真正命中的**。
> 规则：**规范名 ≠ 实装名时，一律以实装名为准**，生成 `.pptd` 时把 `fontFamily` 写成下表的「实际安装名」，或让导出器按下表映射。
> 若字体未装（见"未安装"节），按"字体不足"流程处理。

## 规范名 → 实际安装名 映射

### 中文
| 规范名（fonts.md/设计用） | 实际安装名（本机） | 免费商用 | 说明 |
|---|---|---|---|
| MiSans | `MiSans` | ✅ 小米官方全球免费商用 | 默认中文首选；多字重已装 |
| Noto Sans SC | `Noto Sans SC` | ✅ | 别名：`思源黑体 CN`（同为 Source Han Sans） |
| Source Han Serif / 思源宋体 | `思源宋体 CN` | ✅ | 别名：`Noto Serif SC` |
| 阿里妈妈刀隶体 | `阿里妈妈刀隶体` | ✅ 阿里官方免费商用 | |
| 阿里妈妈东方大楷 | `阿里妈妈东方大楷` | ✅ | |
| 阿里妈妈数黑体 | `阿里妈妈数黑体` | ✅ | 本机仅 Bold 字重 |
| 站酷文艺体 | `站酷文艺体` | ✅ ZCOOL | 亦已装 `站酷高端黑`、`站酷快乐体2016修订版` |
| 得意黑 (Smiley Sans) | `得意黑 斜体` | ✅ | 本机仅斜体字形（符合其设计） |
| 飞波正点体 | `飞波正点体` | ✅（用户确认免费商用） | |
| 霞鹜新致宋 | `霞鹜新致宋＋` | ✅ LXGW | 注：名字带「＋」 |
| LXGW Bright | `霞鹜文楷` | ✅ LXGW | 霞鹜文楷家族已装 |
| ZCOOL KuaiLe | `ZCOOL KuaiLe` | ✅ ZCOOL | |
| 精品点阵体 | `精品点阵体9×9…`（1.93 R/B、方格/光晕/渐层/立体/圆形版、港版） | ✅（用户确认） | 用 `精品点阵体9×9 1.93 R` 作默认 |

### 英文
| 规范名 | 实际安装名 | 免费商用 | 说明 |
|---|---|---|---|
| Liter | `Liter` | ✅ OFL | |
| HedvigLettersSans | `Hedvig Letters Sans` | ✅ | 注意空格 |
| Oranienbaum | `Oranienbaum` | ✅ OFL | |
| QuattrocentoSans | `Quattrocento Sans` | ✅ OFL | 注意空格 |
| Unna | `Unna` | ✅ OFL | |
| Coda | `Coda` | ✅ OFL | |
| Jersey15 | `Jersey 15` | ✅ OFL | |
| Jersey20Charted | `Jersey 20 Charted` | ✅ OFL | |
| SortsMillGoudy | `Sorts Mill Goudy`（注意空格） | ✅ OFL | 本机已装（SourcesMillGoudy-Regular/Italic.ttf） |

### 本地补充（已装、Windows/免费）
| 字体 | 用途建议 |
|---|---|
| `Century Gothic` / `Bahnschrift` | **几何感**，适合 Bauhaus/设计标题（比 Arial 更对味） |
| `Franklin Gothic` / `Impact` / `Rockwell` | 展示/强调标题 |
| `Georgia` / `Cambria` / `Constantia` | 学术衬线正文 |
| `Microsoft YaHei`（微软雅黑） | 默认中文回退 |
| `SimHei` `SimSun` `FangSong` `KaiTi` | 系统标准中文 |
| `更纱黑体 SC`（=Sarasa Gothic J/K） | MiSans 类似，备选中文黑体 |

## 字体不足 处理提示（给 Agent 的模板）
当 `.pptd` 所需字体经 `check_fonts.py` 判定为「缺失」时：
1. 明确告知：**"字体不足：`MissingList`"**（用实际安装名后的缺失列表）。
2. **询问是否列出缺失字体**（供用户安装）：列出缺失字体的规范名 + 免费商用来源 + 下载地址（如 Google Fonts / 官方声明）。
3. 给出**本地可用的替代**建议（如 `SortsMillGoudy`→`Georgia`；`MiSans`→`更纱黑体 SC`；`QuattrocentoSans`→`Century Gothic`），并询问是否改用替代。
4. 用户确认后，把 `.pptd` 里缺失字体改写为已装字体，再继续生成。
