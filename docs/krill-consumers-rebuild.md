# 磷虾油核心消费群体：可编辑重建

输出：`exports/krill-consumers-editable.pptx`，画布 1672×941。
四类人群名称、年龄标签、两条说明和人群洞察按截图保留，未独立验证消费者画像或市场结论。
人物/家庭/盾牌图标为 SVG，并内嵌透明 1024×1024 PNG 回退；磷虾、胶囊及卡片装饰使用原生形状近似重绘，不是原始图形文件。

每张卡片的两条说明是真正的原生项目符号段落，支持继续换行和编辑；不是散落的文字框或手工键入的方块字符。
使用原生蓝色方块、悬挂缩进和段内软换行，保留截图的主要断行位置。
当前导出器会将自定义方块输出成 `buNone`，因此生成器通过标准 DrawingML 属性恢复原生方块、颜色和字号，并清除分组锁；不修改上游技能文件或原始导出。

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-krill-consumers.mjs
```

每次运行位于独立的 `.local/krill-consumers/run-*/`，`output/` 包含 PPTX，`build/` 包含预览和报告。
原始导出、原生格式修复、便携性后处理及最终文件各自保留，互不覆盖。

本次执行了官方 finalizer、重新导入、悬挂缩进和文字检查、8条原生项目符号/7个软换行核对、分组锁检查、SVG 回退及页面溢出检查。
当前内置预览器没有正确应用原生项目符号的颜色/字号，因此另用 LibreOffice 将交付文件转换为 PDF 并检查蓝色方块和换行。
LibreOffice 转换没有重新保存或修改交付的 PPTX；它不能替代桌面 PowerPoint/WPS 的实际编辑测试，本次未进行后者。
验证记录见 `docs/krill-consumers-verification.json`。
