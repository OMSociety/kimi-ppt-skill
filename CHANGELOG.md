# Changelog

本项目所有重要更改都会记录在此文件。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)；
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [1.0.2]

### ✨ 新增

- **字体解析跨平台**：`scripts/check_fonts.py` 与 `scripts/pptd_to_png.py` 不再写死 `C:\Windows\Fonts`——Windows 仍读字体注册表，Linux / macOS 改扫系统字体目录（fontconfig / Font Book 常用路径，去掉字重后缀还原族名），并新增系统兜底字体解析（按需选 bold/常规字形），预览渲染在非 Windows 机器上不再因找不到字体而失败。

### 🐛 修复

- `reference/local-fonts.md` 里的两处失效引用（指向已被中文版 `SKILL.md` 删除的旧小节名 `Font availability check`）改为指向现行「三、字体核对（防豆腐块）」。

### ⚙️ 变更

- `reference/local-fonts.md` 明确**映射表是某台基准机的快照、不是通用事实**：换机器先跑 `check_fonts.py`，以脚本输出为准；`SKILL.md` 相应把「写本机实装名」改为「写目标机的实装名」，并说明字体目录的平台差异。

## [1.0.1]

### ⚙️ 变更

- `skills/kimi-ppt/SKILL.md` 由英文叙述改为**中文精炼版**：按「需求确认 → 前置检查 → 字体核对 → 生成 → 校验 → 导出交付」六步重排，篇幅从 217 行压到 116 行，去除与 `reference/` 重复的英文长段说明。
- 补齐被原版分散在正文中的关键约束：需求确认四维度表、工程目录布局、复刻/编辑/套模板/风格迁移四类生成策略、配图四条纪律、导出后 CT_Slide 根级 `transition` 校验、动画与备注默认不加。
- 明确**本地轨道为默认**：`pptd_to_pptx.py`（纯本地、离线）优先，`export_pptx.py`（字体嵌入 + 淡入淡出）作为桌面增强轨道，沙箱受限时直接回退、不反复重试。
