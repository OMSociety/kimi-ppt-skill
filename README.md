# kimi-ppt-skill

> DeepSeek Harness (DSH) 演示文稿技能：创建 / 编辑 / 复刻 / 导出 PPT，默认产出「可编辑 PPTD 项目 + 本地生成的 .pptx」。
> 源自 MIT 开源的 [`open-kimi-ppt-skill`](https://github.com/binaryify/open-kimi-ppt-skill)（逆向 Kimi Slides），并针对 DSH 做了**本地导出增强**。

## 这是什么

基于 Moonshot 的 `.pptd`（YAML DSL）做演示文稿，支持：

- **创建 / 编辑 / 复刻 / 读取 / 导出** PPT、PPTX、PPTD。
- 内置**设计系统**（学术 / 咨询 / 金融 / 促销 / 工作等场景）与字体规范、海报/信息图。
- **DSH 内纯本地导出**：
  - `scripts/pptd_to_pptx.py` —— python-pptx 本地生成 `.pptx`（无浏览器 / 无外网）。
  - `scripts/pptd_to_png.py` —— Pillow 本地渲染每页预览图（视觉 QA）。
  - `scripts/check_fonts.py` —— 按本机实装字体验证，缺失时提示 + 给本地替代。

## 安装（在 DSH 里）

本仓库根即技能目录（`SKILL.md` + `reference/` + `scripts/` + `tests/`）。

```bash
# 克隆（Windows 下 git 需代理 + openssl，见下）
git clone https://github.com/OMSociety/kimi-ppt-skill.git
# 把仓库内容放到 DSH 用户级技能目录
# Windows:
#   xcopy /E /I kimi-ppt-skill "%USERPROFILE%\.dsh\skills\kimi-ppt"
# macOS/Linux:
#   cp -r kimi-ppt-skill/* ~/.dsh/skills/kimi-ppt/
```

之后 DSH 会自动发现技能（对话里直接说"帮我做一份 PPT"即触发）。

## DSH 本地导出 vs 浏览器导出

- **DSH 默认路径**：`pptd_to_pptx.py`（本地、python-pptx）+ `pptd_to_png.py`（本地、Pillow）。稳定、离线。
- **浏览器导出**：`export_pptx.py` / `export_images.py`（依赖 `agent-browser` + 连 Kimi 公共编辑器）—— 唯一能带**字体嵌入**与**淡入淡出翻页**的路径，但需要在**非 DSH 沙箱**的桌面环境，并可能因 socket 权限/网络超时失败。
- **字体**：按本机实装解析（`reference/dsh-fonts.md` 是 规范名→实装名 映射 + 免费商用标记）。生成前跑 `check_fonts.py` 检查；缺失则提示并给本地替代（如 MiSans→更纱黑体 SC，SortsMillGoudy→Georgia）。

## 结构

```
SKILL.md            # 技能入口（frontmatter name=kimi-ppt）
reference/          # 规范 / 设计系统 / 字体（pptd.md、slides_categories.md、dsh-fonts.md、design_system/…）
scripts/            # 导出 / 预览 / 字体检查脚本
tests/              # 测试
```

## 依赖（本地导出）

- 仅需 `python-pptx`（PPTX）与 `Pillow`（预览）；`PyYAML`。
- 浏览器导出还需：Node.js 18+、`agent-browser`、一个 Chromium 浏览器、可连 `www.kimi.com` / `statics.moonshot.cn`。

## License

MIT。**衍生自** [`binaryify/open-kimi-ppt-skill`](https://github.com/binaryify/open-kimi-ppt-skill)（MIT，Kimni/奇梦尔 Slides 逆向项目）。本仓库含其参考文档/脚本，并新增 DSH 本地导出/预览/字体检查等增强。
