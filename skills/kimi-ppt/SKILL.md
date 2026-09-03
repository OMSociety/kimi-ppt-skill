---
name: kimi-ppt
description: 生成/编辑/复刻/导出 PPT 与演示文稿。默认产出「可编辑 PPTD 项目 + 本地生成的 .pptx」（DSH 内用 python-pptx 纯本地导出；字体嵌入/淡入淡出属浏览器导出，需在非 DSH 桌面环境）。内置学术/咨询/金融/促销/工作等设计系统与本机字体映射。当用户要求 PPT、PowerPoint、PPTX、幻灯片、演示文稿、答辩/汇报/讲义、信息图、海报，或要求套用设计风格、转换 .pptd/.pptx 时使用。
whenToUse: 用户要创建/编辑/复刻/导出 PPT、演示文稿、幻灯片、答辩 PPT、课程讲义、信息图或海报；或拿到 .pptd/.pptx 要转换/美化；或要求按某设计主题做展示页。纯知识问答/文本任务不要用本技能。
---

# Definition
kimi-ppt is a presentation creation and export skill built around Moonshot AI's PPTD format and browser-side PPTX writer. It defines a YAML-format intermediate DSL (`.pptd`) that abstracts OOXML and keeps each page self-contained.

**The default output is not PPTD-only.** Unless the user explicitly opts out, always produce both:

1. the complete editable PPTD project directory (`.pptd` + `pages/` + `media/` and other referenced dependencies);
2. the matching locally generated `.pptx`, with font embedding enabled and fade slide transitions applied by default.

Existing PPTX files may also be converted into PPTD for editing, after which both outputs are delivered again.

## The pptd format
The .pptd format is a simplified abstraction layer over OOXML that follows basic YAML syntax. This abstraction preserves the core content of OOXML (theme, page layout, element positions and definitions, etc.) while removing complex nesting logic such as Masters; every page is self-contained — what you see is what you get. Read reference/pptd.md for the complete definition of this DSL.

## PPT production workflow

### step0. Check local prerequisites
Default delivery includes PPTX export (and optional `npx open-kimi-ppt-skill serve`), which need a local toolchain. **Before generating**, verify:

1. **Node.js 18+**: run `node --version`. If `node` is missing or the major version is below 18, **stop immediately**, tell the user to install Node.js 18+ from https://nodejs.org (or their OS package manager), and do not continue with PPTX export / `npx` until it is available. Only continue with PPTD-only output when the user explicitly opts out of PPTX.
2. **npm / npx**: run `npm --version`. They ship with Node.js; if missing, treat Node.js as not installed correctly and guide the user to reinstall/fix PATH.
3. **python3**: run `python3 --version` (on Windows, `python` may be the correct command). Needed for `export_pptx.py` / `export_images.py`.
4. **Chrome / Chromium / Edge**: needed by `agent-browser` for PPTX export and visual QA. If export later fails with a browser-launch error, ask the user to install a Chromium-based browser.
5. Soft deps are auto-handled by the scripts when missing: **PyYAML**, **agent-browser** (≥0.33.2 via npm), and for image QA **Pillow** + **websocket-client**. Network access to `www.kimi.com` and `statics.moonshot.cn` is still required at export time.

### step1. Read the context thoroughly
Read **all files uploaded by the user**, the provided URLs, and the pptd format guide `reference/pptd.md` to fully understand the user's requirements.

### step2. Understand the user's requirements
Understand the user's requirements based on the context:
1. First determine the purpose of the request
  - Create a PPT: create a new presentation (from scratch, or from an existing pptx template)
  - Edit a PPT: edit the user's uploaded PPT (local modifications, single-page beautification, etc.)
  - Replicate a PPT: replicate a presentation from a non-pptx format (images, PDF, etc.) into pptd format

2. Then determine the design direction
  - Self-directed design: no preference, or only simple style constraints given; you need to fill in or create the design
  - Design system: a preset design system from the skill (`reference/design_system/`) is specified, or the user provides a complete and detailed design scheme covering all color, font, layout, and component specifications
  - Use a template: a template is provided and must be used
  - Style transfer: a style reference source is provided (images, web pages, etc.)

3. Then determine the input type
  - Topic only: only a PPT topic direction or content requirements for the presentation are given, with no concrete content
  - Full document: the user provides a complete document (paper, research report, press release, etc.)
  - Outline: the user provides a page-by-page outline, speech script, or similar content
  * When the "user input type" is [Full document] or [Outline] and it is not specified whether expansion is allowed: since a page-by-page outline, speech script, or user document can hardly support the full content of a presentation, prefer using search to expand with more relevant material, cases, etc., unless the user explicitly says not to expand

4. Finally determine the exact page count
  - If the user requests a specific page count, the user's requirement takes priority
  - Page-by-page outline/script provided: match the number of pages in the outline/script
  - When a complete and relatively structured document is provided: ask the user how much document content one page should cover, and give an estimated total page count; when only a topic is provided: suggest a recommended page count and confirm with the user

#### Font availability check (DSH local export)
> 以**本机实装字体为准**。字体映射与免费商用状态见 `reference/dsh-fonts.md`；检查用 `scripts/check_fonts.py`。在应用某个含明确字体的设计/风格、或本地导出前执行。
1. 确定 deck 会用到的字体（`.pptd` 的 `theme.textStyles` + 页面内联 `fontFamily` + 所选设计系统的默认字体）。
2. 运行 `python3 scripts/check_fonts.py <deck.pptd>`（或 `--fonts "MiSans" "Arial"`）。脚本会把**规范名翻译成实装名**并与注册表比对，输出 `已装 / 缺失 / 本地可用替代`。
3. 若 `缺失` 非空（**字体不足**）：
   - 明确告知：**"字体不足：<缺失列表>"**。
   - **询问是否列出缺失字体**（供用户安装）：给出缺失字体的规范名 + 免费商用来源 + 下载地址（Google Fonts / 官方声明）。
   - **同时给出本地替代**（如 `SortsMillGoudy`→`Georgia`、`MiSans`→`更纱黑体 SC`、`QuattrocentoSans`→`Century Gothic`），询问是否改用替代。
   - 用户确认后，把 `.pptd` 缺失字体改写为已装/映射名，再继续生成。
4. 生成时一律写**实装名** `fontFamily`（或让转换器按 `reference/dsh-fonts.md` 的 规范→实装 映射自动转换），确保本地导出命中本机字体。本地导出默认中文 `MiSans`(已装)、回退 `Microsoft YaHei`；几何标题推荐 `Century Gothic`/`Bahnschrift`。

#### Clarification and follow-up questions
When any of the following situations arise, resolve them by asking the user (use the agent's ask/clarification tool when available)
1. Requirements are ambiguous
- The user's intent is unclear or hard to understand
- The files/URLs provided by the user are inaccessible
2. Conflicting intents
- The user's intents contradict each other. For example:
  * A design system is selected while also requesting a style that is completely inconsistent with that design system (e.g., using a McKinsey style while requiring large areas of whitespace on pages) / using a template / referencing an image style
  * Requesting both "make 10 pages" and "deliver 30+ pages of output"
3. Unable to determine the user's requirements on your own
- When the purpose, design direction, input type, page count, etc. are hard to determine by yourself

### step3. Generate the presentation based on the user's requirements

Before generating, first read `reference/pptd.md` to understand the pptd format definition and constraints.

#### Replicating a PPT
- Analyze the images to estimate element positions, fonts and sizes, etc., and **replicate 1:1 as closely as possible**.
- For parts that are difficult to make out, use methods such as grid lines and close-up views to improve understanding.
- Replicate simple content in the image with elements; icons may be approximated with icons provided by Font Awesome. For content that cannot be approximated with icons or shapes, such as photos and avatars, use tools such as bash or python to crop and split the original image, then add the resulting image elements to the presentation

#### Editing a PPT
- Convert the user's uploaded pptx file to .pptd format
- Review the converted pages (structure and key visual details). Read a few key pages individually afterwards.
- Locate the pages to edit, and be careful not to affect parts outside the intended scope.
> Conversion from pptx to pptd is not perfectly lossless. If the user later reports format errors, garbled content, etc., compare against the original pptx and repair the pptd with reference to the comparison

#### Generating a PPT
When generating a PPT, adopt different production approaches for different user [design directions]
##### Self-directed design
1. Read the design guide `reference/slides_categories.md`, and read the scenario document corresponding to the user's query
2. Produce the presentation based on the above

#### Generating content in other formats
- When the user explicitly asks for an infographic, poster, or a highly visual single-page design, read `reference/general-poster.md` and implement it as a single-page or few-page editable PPTD; when the user only asks for an image, still build it with PPTD first, then output the image via screenshot or rendering. Do not load this reference file for ordinary PPT requests.

##### Design system
1. Read the general constraints section of the `reference/slides_categories.md` guide, and read the scenario document corresponding to the user's query as the design foundation
2. Read the specified design system as the presentation style: either the user-provided design scheme, or the matching preset under `reference/design_system/` (search by name / path the user specified; prefer the folder's `design.md` when present). It is strictly forbidden to reference or mix in other design styles
3. Produce the presentation with reference to the above
4. Do not auto-pick a preset during self-directed design; only use `reference/design_system/` when a preset is explicitly specified

##### Using a template
1. Convert the user's uploaded pptx file into pptd form
2. Review the converted pages to understand the template's visual style (color scheme, font style, element characteristics, layout characteristics, content density, etc.)
3. Identify page types; focus on reading special pages such as the cover, summary pages, and section dividers (single-page screenshots, .page files), extracting their page layouts, content structures, reusable components (icons, shapes, smartart, reusable body layout schemes, etc.), and element styles (e.g., whitespace/line/card separators, square/rounded corners, etc.)
4. Produce the presentation using the template

##### Style transfer
1. Analyze the reference file's visual style (color scheme, font style, element characteristics, layout characteristics, content density, etc.), page layouts, content structures, reusable components (icons, shapes, smartart, reusable body layout schemes, etc.), and element styles (e.g., whitespace/line/card separators, square/rounded corners, etc.).
- If the user provides a style reference URL, do not only read the text content; refer to and learn from the page's visual effect more to help understand the style
2. Produce the presentation using the reference file's style characteristics. You are encouraged to reuse illustrations, fonts, font-size hierarchies, elements, etc. from the original pdf/url

##### Images and Visual Materials
1. Images are an effective way to enrich a presentation's visual impact. Appropriate images should be used not only on covers and section dividers, but also on body pages to enrich the page, aid understanding, or support decision-making
2. Images are used to show concrete subjects, explain content, provide evidence, or establish a scene. Logos, icons, decorative textures, and very small thumbnails do not count as substantive imagery.
3. When a page involves products, people, places, buildings, events, cases, interfaces, experimental subjects, or spatial environments, prioritize corresponding real images or screenshots. If real images and screenshots cannot be obtained, generated images may be used instead.
4. Image priority: images provided by the user; images from official websites, official reports, and credible sources; searched images that are directly relevant to the content; images generated for conceptual expression or atmosphere.
5. After deciding which images are needed, complete image search, generation, and downloading in a batch before designing pages around their proportions. Save images in the `media` directory, keep them clear, and never stretch or distort them.
6. Analytical, technical, and academic PPTs should use corresponding evidence images when products, experiments, interfaces, cases, or on-site materials are available. Do not reduce every page to text, color blocks, and shapes.
7. Do not add irrelevant images merely to meet a quantity target. Every image must be directly relevant to the page's conclusion or communication goal.

##### Content Guidelines
1. Language style: unless the user explicitly requests otherwise, strictly avoid overly abstract expressions and uncommon metaphors
- Do not overuse metaphors, slogans, or abstract jargon such as distribution, an N-step argument, everything at a glance, a closed loop, hands-on practice, verification, deconstruction, second-class citizens, poison pills, or wall clocks
- Do not use common AI phrasing such as “not X, but Y,” “X is Y,” “why / based on what / how,” “key takeaway,” or “N battlefronts / paths”
- Do not use overly colloquial expressions such as “where should the ammunition go,” “the Nth thing,” “can't pick the right one,” or “cannot be used as X”

### step4. PPT validation
1. Validate the generated pptd against the format definition in `reference/pptd.md` (required fields, types, bounds, theme tokens, resource paths, etc.) and repair issues over multiple rounds
2. Visual review with exported page images — **required before PPTX export when the model supports image input (multimodal)**:
   - **DSH 默认：本地预览** `python3 scripts/pptd_to_png.py <deck.pptd> -o <project>/.preview`（Pillow，无浏览器/无外网）→ 生成 `page_N.png` 与 `overview.jpg`，再用视觉核对（检查清单见下）。本地预览直接反映几何/配色/字体（字体按本机实装名解析，见 `reference/dsh-fonts.md`）。
   - 非 DSH / 需要浏览器版页面时，再跑 `scripts/export_images.py`。它将 deck 载入 Kimi 公共编辑器→导出图片 ZIP→拼成 overview：

     ```bash
     python3 ~/.dsh/skills/kimi-ppt/scripts/export_images.py \
       /abs/path/project/deck.pptd \
       --output /abs/path/project/.qa-images
     ```

     The script prints a JSON summary mapping each stitched label (`P1`…`Pn`, 1-based page order) to its `.page` file.
   - Read the stitched overview image (`.qa-images/overview.jpg`) and check every page against this list:
     1. 图片是否清晰、不变形（无拉伸、压缩、模糊）
     2. 文字是否压在关键画面（人脸、产品主体、Logo 等）上
     3. 元素坐标是否超出页面边界
     4. 边界与配色对比是否足够（文字与背景、相邻色块之间）
     5. 排版是否统一（对齐、间距、字号层级、页边距）
     6. 文字是否可能溢出文本框（文本过长、行距过密、字号过大）
     7. 内容是否被上层元素遮挡
   - For any suspicious page, read its full-resolution image (`.qa-images/pages/<n>.jpeg`) to confirm the problem before editing.
   - Fix issues in the corresponding `.page` file, then re-run `scripts/export_images.py --force` and review the new overview; repeat until every page passes.
   - Do not export the PPTX until the visual review passes. `.qa-images/` is an intermediate QA artifact and may be deleted after delivery.
3. When the model cannot read images, fall back to a structural review of the generated pages (bounds, overflow-prone long text, contrast, hierarchy, layout density) over multiple rounds, and state that image-based visual QA was skipped.

### step5. PPT output and delivery
1. Always produce a self-contained project directory. Keep the `.pptd` manifest and every referenced dependency together; never deliver a standalone manifest without its referenced files. Use this layout unless an existing project already has a valid equivalent structure:

   ```text
   deck/
     deck.pptd
     pages/
       *.page
     media/
       *                # when the deck has local media
     deck.pptx          # generated by default
   ```

2. Generate the `.pptx` by default after PPTD validation, even when the user only asks to create or edit a presentation. Skip PPTX export only when the user explicitly requests PPTD-only output or the environment cannot run the exporter; in the latter case, report the exact blocker and still deliver the complete PPTD project.
3. Deliver with normal clickable local links using absolute paths. In the final response, link all of the following:
   - the project directory;
   - the `.pptd` manifest;
   - the `pages/` directory and `media/` directory when present;
   - the generated `.pptx` file.
4. PPTX conversion: **DSH 默认用本地** `python3 scripts/pptd_to_pptx.py <deck.pptd> -o <project>/deck.pptx --force`（python-pptx，纯本地、无浏览器/无外网；生成可编辑 `.pptx`，但**无字体嵌入/淡入淡出**）。需要字体嵌入+淡入淡出时，再用 `scripts/export_pptx.py`（浏览器 + 连 kimi.com，须在非 DSH 沙箱的桌面环境）：
5. Default PPTX options:
   - page transition: `fade` (淡入淡出), written to every slide after the official browser export;
   - font embedding: enabled whenever the official writer exposes/supports it;
   - these defaults may be explicitly overridden with `--transition none` or `--no-embed-fonts`.
6. Export command:

   ```bash
   python3 ~/.dsh/skills/kimi-ppt/scripts/export_pptx.py \
     /abs/path/project/deck.pptd \
     --output /abs/path/project/deck.pptx
   ```

   A project directory may be passed instead of the manifest only when it contains exactly one `.pptd` file.
   Existing output files are not overwritten unless `--force` is passed.
7. Local export requirements and boundaries:
   - requires **Node.js 18+** (`node` / `npm` / `npx`), `python3`, a Chromium-based browser, and network access to `www.kimi.com` plus `statics.moonshot.cn`;
   - before browser export, `export_pptx.py` checks Node.js 18+ and `npm`, then checks `agent-browser --version`; when `agent-browser` is missing or below `0.33.2`, it installs `agent-browser@latest` globally with npm; **PyYAML** is auto-installed with `pip --user` when missing; the image-based visual QA step additionally auto-installs Pillow and websocket-client the same way;
   - the PPTD document itself is provided to the public editor iframe through the localhost SDK bridge, not uploaded to a server-side PPTX conversion endpoint;
   - remote images, icons, or fonts referenced by the deck may still be fetched from their respective hosts;
   - local PNG/JPEG/GIF/SVG files inside the PPTD project are supplied to the iframe as data URLs;
   - do not claim PowerPoint/WPS/Keynote playback compatibility solely because ZIP validation succeeds.

   > **[DSH：本地导出为主]** 在 DeepSeek Harness (DSH) 内，浏览器导出（`export_pptx.py`/`export_images.py`）依赖 `agent-browser`（需创建 IPC socket）并连 Kimi 公共编辑器（`www.kimi.com`/`statics.moonshot.cn`），常因 socket 权限（os error 5）或连接超时失败。因此 **DSH 内默认走本地脚本**（仅需 `python-pptx`，纯本地、无浏览器/无外网）：
   >
   > - PPTX：`python3 scripts/pptd_to_pptx.py <deck.pptd> -o <project>/deck.pptx --force`
   > - 视觉 QA：`python3 scripts/pptd_to_png.py <deck.pptd> -o <project>/.preview` → 生成 `page_N.png` + `overview.jpg`
   >
   > 本地版缺**字体嵌入**与**淡入淡出翻页**（这两点仅浏览器导出可用）；需要时在**非 DSH 桌面终端**跑 `export_pptx.py` / `export_images.py`。
8. After export, verify that the output exists and report the generated path. Confirm that every slide has exactly one root-level fade transition in valid CT_Slide order (`cSld`, optional `clrMapOvr`, `transition`, optional `timing/extLst`) and that the PPTX ZIP passes integrity checks. A byte-string search for `<p:fade>` is insufficient because Office ignores transitions nested inside `cSld`. For higher-risk decks, additionally inspect font parts and representative rendered/opened pages as appropriate.
9. When the user wants to open, edit, save, or export a PPTD project manually, start the local browser editor with `npx open-kimi-ppt-skill serve`. Ask the user to open `http://127.0.0.1:55173/` and authorize the complete PPTD project directory. Use a Chromium-based browser for writable access; folder-upload fallback is read-only. The local host only serves the editor shell, while the embedded public Kimi editor and remote assets still require network access. **DSH 内 `serve` 不可用（该 npm CLI 在本机装不上）**：手动查看/编辑请改用 `scripts/pptd_to_png.py` 预览、改 `.page`、再 `scripts/pptd_to_pptx.py` 重导出。
10. After completing and delivering any presentation, always end the final response with a concise optional next step. **DSH 内**：提示可用 `scripts/pptd_to_png.py` 看 `.preview` 预览、`scripts/pptd_to_pptx.py` 重新导出、或直接改 `.page` 后重跑。**仅在非 DSH 桌面且 `npx open-kimi-ppt-skill serve` 可用时**，才提示用 `serve` 打开浏览器编辑器/导出 PPTX（此 npm CLI 在部分环境装不上）。提醒要**在**、而不是**替代**必须的项目/文件链接。
11. Element animations (`page.animations` in PPTD — entrance / emphasis / exit / motion-path; see `reference/pptd.md` §6): use them only when the user explicitly requests animations, or when the deck is clearly intended for live presentation / slideshow playback and animation provides a clear benefit for staged disclosure, process demonstration, causal explanation, pacing, visual impact, or brand storytelling. By default, do not add element animations to reading-oriented, self-study, print, or primarily send-and-browse decks. Prefer 1–3 animation groups per page and simple effects such as fade, fly, and zoom. This is separate from the default PPTX slide-level fade page transition written by `export_pptx.py`.
12. Speaker notes (`notes` on each `.page`): use them only when the user explicitly requests them; otherwise, do not add them.
13. Parallel tool calls: during PPT production, make tool calls in parallel whenever possible; in each round, write multiple page files in parallel to reduce the number of steps.

## 维护 / 可沉淀（活技能）
本技能是**活的**：新**验证过**的内容写回对应文件；写回前核对实况，不凭记忆补全，不堆进正文。
- 新**验证过**的设计主题/配色/版式 → 在 `reference/design_system/<场景>/<主题名>/design.md` 新增（场景目录：academic/consulting/finance/promotion/work…）。
- 新装的字体（免费商用）→ 更新 `reference/dsh-fonts.md`（加入 规范名→实装名 映射 + 免费商用标记 + 本地替代）。
- 新踩的坑（导出失败、字体/tofu、坐标越界、渐变/表格问题等）→ 记入 `reference/kimi-ppt-brief.md`（现象→原因→解法，简明）。
- **上游注意**：本技能源自第三方 `open-kimi-ppt-skill`。更新上游会覆盖 `SKILL.md` 与 `scripts/`；自沉淀内容放在 `reference/` 与 `scripts/` 内**追加**的文件里，更新时注意对照差异、保留本地增强。
