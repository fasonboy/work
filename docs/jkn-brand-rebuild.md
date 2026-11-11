# 台湾加康诺品牌实力：重建记录

输出：`exports/jkn-brand-editable-draft.pptx`。当前是**缺少两份登记文件原图的可编辑草稿**。
聊天预览可见，但云端没有收到可读取的原始截图。右侧保留独立框位并明确标记“待补截图原文件”；未仿造文件扫描件、未重画印章、未补写模糊内容。

保留 1672×941 画布。公司名称、介绍、品牌实力、四组说明、经营成果、截止日期、销售额及货币单位均为独立文本。
“1,000万 台币”及“截至2026年10月”按截图保留，未进行销售数据的独立核验。
Logo为原生节点、线条和可编辑文字的近似重绘。图标、奖章、奖杯为 SVG；桂冠、色带及底部波浪可独立选择。这些图形与官方原始美术资产并非同一文件。

运行：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-jkn-brand.mjs
```

收到截图原文件后可补入右侧原件：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-jkn-brand.mjs --reference /absolute/path/screenshot.png
```

输入须与本次截图布局、比例一致。程序只裁切两份文件图，保留其原始内容，不把整张截图铺成背景。
嵌入后，扫描件仍是可移动和替换的图片对象，其小字及印章不是可编辑文字。
新运行输出保存在独立的 `.local/jkn-brand/run-*/output/`，不会覆盖用户原始截图。
原图可作为文件附件、包含原图的 ZIP，或上传到 GitHub 仓库后通过文件路径传入。

已执行官方 finalizer、重新导入、渲染和独立结构检查。报告见 `docs/jkn-brand-verification.json`；当前构建及预览在被忽略的 `.local/jkn-brand/revision-3/build/`。
未在桌面 PowerPoint/WPS 中打开验证。补齐文件图片后需要重新检查最终输出。
