# 图片转可编辑 PowerPoint 云端项目

技能安装在 [.agents/skills/image-to-editable-ppt/SKILL.md](.agents/skills/image-to-editable-ppt/SKILL.md)。
上游：<https://github.com/zhoujie97/image-to-editable-ppt-skill>，固定版本 `22945dfc705689354fb167b8a609c7d4dfccb6a5`。
四个上游文件保持原样，文件清单和 SHA256 见 [来源记录](docs/image-to-editable-ppt-source.json)。

## 当前截图的可编辑草稿

[下载 PPTX](exports/editable-ppt-draft.pptx) 或 [下载 ZIP](exports/editable-ppt-draft.zip)。
草稿已重建文字、指标及条形图，右侧六个视频画面仍待提供截图原文件后补齐。
在 GitHub 文件页面使用 **Download raw file** 下载；ZIP 内包含同一份 PPTX。
重建脚本和验证范围见 [重建记录](docs/short-video-rebuild.md)。

## 使用

在后续任务中上传图片，并要求：

> 使用 $image-to-editable-ppt 将这张图片重建为可编辑 PPTX，保持布局，完成便携性处理、渲染和溢出检查。

这是供 Codex 使用的重建技能，需要助手分析图片并编写对象布局，不是无需分析即可自动转换任意图片的命令行程序。
文字必须使用可编辑文本框，表格使用原生表格，关系箭头使用原生连接符。
不得将整页图片作为背景冒充可编辑内容。

先读取本项目技能及其 `references/reconstruction-workflow.md`，再读取内置 `presentations:Presentations` 的说明、样式指引、Artifact Tool 文档和验收要求。
本实例的内置技能目录为 `/opt/codex/skills/builtins/presentations`；若平台调整路径，可设置 `PRESENTATIONS_SKILL_DIR` 指向实际完整技能目录。

## 一条命令复测

在项目根目录执行：

```bash
bash scripts/check-ppt-environment.sh
```

脚本以失败退出状态报告不可用依赖，依次检查：

- 上游四个技能文件、SHA256、SKILL.md 元数据、相对文档引用及 JS 语法。
- 云平台提供的 Node.js、Artifact Tool、sharp 和 Presentations 技能及工具文件。
- sharp 将 SVG 实际栅格化为透明 PNG；zip/unzip 打包、CRC 检验及内容往返一致性。
- 实际创建一页 PPTX，包含文本、形状、原生带箭头连接符、3×2 原生表格、SVG 图标。
- 执行上游后处理脚本，清理组合锁及字体继承，生成 1024×1024 PNG 回退图。
- 执行内置 finalizer，验证文件结构、页面尺寸、字体策略和原生表格，并重新导入 PPTX。
- 独立解析最终 OOXML，验证可编辑对象、内部图片关系、PNG 回退、字体设置和 ZIP 完整性。
- 渲染最终页面并执行内置溢出测试。溢出测试的输出也必须明确成功，不能只依赖退出码。

每次生成独立目录 `.local/ppt-checks/run-*/`，包含 `report.json`、归一化结果、验证回执、PPTX 和 PNG 预览；失败不会被旧结果掩盖。
这些运行产物已忽略，不作为项目源码保存。首次验证结果见 [验证记录](docs/ppt-environment-verification.json)。
复测会重新生成一页测试文稿；它只证明工具链能运行，不代表任意输入图片的识别、保真度或 PowerPoint/WPS 原生交互已经验证。

## 编写生成脚本时激活运行时

```bash
source scripts/ppt-runtime.sh
```

脚本使用平台注入的 `CODEX_PRIMARY_RUNTIME*` 路径，导出 `RUNTIME_NODE`、`RUNTIME_NODE_MODULES`、`RUNTIME_PYTHON`、`RUNTIME_BIN_DIR` 和 `CODEX_NODE_MODULES`。
不要将运行时绝对路径写死在项目生成器中。Node 的 ES module 默认不会按 `NODE_PATH` 查找包；可参考 `scripts/ppt-smoke-test.mjs`，从运行时路径解析后动态导入，或按内置技能在私有构建目录创建 `node_modules` 链接。

本实例已提供所有核心依赖，无需 npm 安装或额外密钥。`@oai/artifact-tool` 是平台提供的私有包；换到不带该包或 Presentations 技能的环境时，检查脚本会失败，不能把普通 Node.js 环境视为具备生成能力，也不能假设从公共 npm 可以补齐私有依赖。

归一化单个文稿（输入与输出必须不同）：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" .agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs \
  --input /absolute/path/raw.pptx \
  --output /absolute/path/portable.pptx \
  --sharp-module "$RUNTIME_NODE_MODULES/sharp" \
  --min-native-connectors 1 \
  --min-target-arrows 1 \
  --min-native-tables 1
```

数量参数须按实际页面调整。后处理成功不能代替最终 PPTX 的渲染、图片关系/回退检查及逐页人工检查。
项目不依赖常驻服务；每个新任务应重新激活运行时并检查依赖。
