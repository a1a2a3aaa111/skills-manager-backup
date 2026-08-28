---
name: "MK-PRD转Word"
description: "将 Markdown 需求文档和 HTML 高保真原型自动合并、截图并导出为 Word 说明书。仅手动触发。"
disable-model-invocation: true
---

# Word 导出生成器 (MK-PRD转Word)

当需要将现有的 Markdown 格式的需求文档（PRD）和本地的 HTML 高保真原型整合成一份供客户签字的 Word 说明书时，使用此 Skill。

## 工作原理

1. **自动生成映射（可选）**：运行 `auto_generate_map.py` 脚本，它会扫描指定 `prd_dir`（PRD 根目录）下的 `stories/` 目录，通过正则提取故事文件中的各个页面名称，并智能匹配 `prototype_dir`（原型根目录）下最接近的 HTML 文件，生成 `export-map.yaml`。
2. **配置映射**：读取项目根目录下的 `export-map.yaml` 配置文件，该文件定义了 PRD 和原型的根目录路径，以及哪个 Markdown 故事文件中的哪个页面，需要插入哪个 HTML 文件的截图。
3. **自动截图与精准替换**：使用 Python Playwright 自动打开映射好的本地 HTML 文件并截取全屏图片。然后精准替换 Markdown 故事文件中对应页面标题（如 `### 1.1 页面名称：XXX`）下方的 ASCII 线框图（` ```text ` 块）。
4. **多文件有序合并**：由于 PRD 是多文件结构，脚本会自动按规范顺序将文件拼接为单个临时 Markdown 文件：`主文件` -> `models/*` (按序号升序) -> `stories/*` (按序号升序，包含已替换高清截图的内容)。
5. **自动渲染 Mermaid**：自动识别拼接后 Markdown 中的 Mermaid 架构图（如流程图、状态机、实体关系图等），利用无头浏览器在后台实时渲染并截图替换为高清图片。
6. **生成 Word**：使用纯 Python 原生方案（`markdown2` + `python-docx`），将最终的完整 Markdown 文件解析并排版导出为 `.docx` 格式的 Word 文档，无需依赖外部 Pandoc 工具。

## 使用前提

使用此 Skill 前，需要确保环境满足以下条件：
1. 已安装 Python 3。
2. 安装跨平台依赖包：进入本 Skill 根目录，执行 `pip install -r requirements.txt`。
3. 安装浏览器内核（用于截图）：执行 `playwright install chromium`。

## 映射文件规范 (`export-map.yaml`)

可以通过运行脚本自动生成，或者手动在项目根目录创建一个 `export-map.yaml` 文件，格式如下（**强烈建议统一使用相对路径，以确保配置在 Mac、Windows 或云端均可通用**）：

```yaml
# 全局配置
config:
  prd_dir: "./PRD文档/车辆管理-v2"           # 必填：PRD 根目录（建议相对路径，包含 主文件、models/、stories/）
  prototype_dir: "./车辆管理-v2-原型/原型包"         # 必填：原型根目录（建议相对路径，包含 pages/）
  output_dir: "./output"                         # 最终生成的 Word 文档输出目录
  screenshot_dir: "./assets/screenshots"         # 截图保存目录
  output_filename: "产品需求说明书.docx"
  reference_doc: "可选：指定公司专属的 Word 排版模板"

# 映射规则
mapping:
  - markdown: "stories/10-车辆管理员记录保养信息.md"
    pages:
      - page_name: "保养记录列表"                  # 故事文件中定义的页面名称（用于定位）
        html: "pages/保养记录-list.html"         # 对应的原型 HTML 相对路径
      - page_name: "保养记录表单"
        html: "pages/保养记录-form.html"
```

## 执行方式

当用户触发此 Skill（要求导出 Word）时，Agent **必须**遵循以下步骤执行：

1. **主动获取路径（强制）**：不要假设路径，必须主动向用户提问：
   > “为了开始导出 Word，请告诉我：\n 1. **PRD 文档的根目录路径**（包含主文件、models/、stories/ 的文件夹）\n 2. **原型的根目录路径**（包含 pages/ 文件夹的原型包）”
2. **检查环境依赖**：确认当前环境安装了 Playwright, PyYAML, BeautifulSoup4, Pandoc 等依赖。
3. **生成并配置映射文件**：获取到路径后，运行 `./scripts/auto_generate_map.py <prd_dir> <prototype_dir> export-map.yaml` 自动生成映射配置。如果根目录已存在 `export-map.yaml` 且路径正确，可提示用户是否需要重新扫描。
4. **人工确认映射**：让用户检查或帮助用户完善 `export-map.yaml`，特别是自动扫描未能成功匹配的 `手动填写.html` 页面。
5. **执行导出**：运行 `./scripts/export_word.py` 脚本生成最终的 Word。
