# 图片转可编辑 PowerPoint 云端项目

技能安装在 [.agents/skills/image-to-editable-ppt/SKILL.md](.agents/skills/image-to-editable-ppt/SKILL.md)。
上游：<https://github.com/zhoujie97/image-to-editable-ppt-skill>，固定版本 `22945dfc705689354fb167b8a609c7d4dfccb6a5`。
四个上游文件保持原样，文件清单和 SHA256 见 [来源记录](docs/image-to-editable-ppt-source.json)。

## 安徽线下门店展示

[下载可编辑版式草稿](exports/anhui-store-editable-draft.pptx)。标题、说明、卡片及地图可编辑；地图为按截图近似重绘，两张门店照片仍待截图原文件补齐。
源文件和继续生成的方式见 [重建记录](docs/anhui-store-rebuild.md)。

## 安徽业务展示（新版）

[下载可编辑版式草稿](exports/anhui-business-editable-draft.pptx)。按新版截图重建标题、副标题、三组说明卡片、金色标签、底部亮点及蓝金波浪；两张原始门店照片仍待补齐，地图为可编辑近似示意。
生成方式与当前限制见 [新版重建记录](docs/anhui-business-rebuild.md)。

## 台湾加康诺：品牌实力

[下载可编辑版式草稿](exports/jkn-brand-editable-draft.pptx)。公司介绍、经营成果和销售额为可编辑文字，Logo及装饰按截图近似重绘；右侧两份登记文件的原图仍待补齐。
说明见 [品牌实力重建记录](docs/jkn-brand-rebuild.md)。

最新简洁版：[下载可编辑草稿](exports/jkn-brand-simple-editable-draft.pptx)。已更新“市场表现稳定”和浅蓝经营成果卡片，仍待补两份登记文件的原图。说明见 [简洁版重建记录](docs/jkn-brand-simple-rebuild.md)。

## N9 胶囊产品页

[下载可编辑 PPT](exports/n9-product-editable.pptx)。文案、版式及右侧三盒包装均以独立对象重建；包装为近似矢量插图，不是原始产品照片。
生成脚本及验证范围见 [N9 产品页重建记录](docs/n9-product-rebuild.md)。

## 线下私域资源展示

[下载可编辑 PPT](exports/private-resources-editable.pptx)。文字、卡片、流程箭头、设备/场景示意与手机列表均可分别修改；设备照片和群聊截图改为近似矢量重绘，头像及截断文字有简化。
说明见 [私域资源重建记录](docs/private-resources-rebuild.md)。

## 南极磷虾油市场现状与发展机遇

[下载可编辑 PPT](exports/krill-market-editable.pptx)。三栏文字、市场数字及装饰分别可编辑；柱状图为原生图表，内嵌2024年/2030年、11/21亿美元的数据表。
说明见 [磷虾油市场页重建记录](docs/krill-market-rebuild.md)。

## 从鱼油的发展看磷虾油市场机会

[下载可编辑 PPT](exports/fish-krill-opportunity-editable.pptx)。左侧为可编辑时间线和百分比柱状图，右侧为真正的4行3列PowerPoint表格，图表附内嵌数据表。
说明见 [鱼油与磷虾油对比页重建记录](docs/fish-krill-opportunity-rebuild.md)。

## 南极磷虾油核心消费群体

[下载可编辑 PPT](exports/krill-consumers-editable.pptx)。四类人群、年龄标签和8条原生项目符号说明可编辑，人物、磷虾和胶囊为近似重绘。
说明见 [消费群体页重建记录](docs/krill-consumers-rebuild.md)。

## 超级短视频总结的可编辑草稿

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
