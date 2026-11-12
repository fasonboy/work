# 南极磷虾油市场前景

输出：`exports/krill-market-prospects-editable.pptx`，单页，画布1672×941。
按用户当前截图重建标题、副标题、左右双栏、数据卡片、营养对比表、成分标签和底部结论。
所有文字保持可编辑；对比表使用4行3列原生PowerPoint表格，支持表格行列操作。
数据卡片按截图保留1999—2000年1.3%、2011—2012年12%及约9倍，没有新增图表或用精确比值替换截图文案。
营养成分、营养形式和市场主张按图转录，未独立验证来源或功效。

云端没有可读取的原始截图文件；雪山、水面、磷虾、胶囊和图标为近似矢量重绘，不是原始照片，外观存在差异。
图标内嵌SVG及1024×1024透明PNG回退。数据卡片之间的两条箭头为原生连接符，四条成分关联线也使用原生连接符。
没有把整页截图作为背景。

重复生成：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-krill-market-prospects.mjs
```

脚本生成独立`.local/krill-market-prospects/run-*/`目录，保存最终PPTX、预览和验证回执；执行上游便携性处理、内置finalizer、原生表格检查，并重新导入最终文件渲染。
修改布局后需重新检查最终文件溢出：

```bash
"$RUNTIME_PYTHON" "$PRESENTATIONS_SKILL_DIR/container_tools/slides_test.py" /absolute/path/final.pptx
```

确认日志明确包含`Test passed. No overflow detected.`，不能仅依赖退出码。
本次验证记录见`docs/krill-market-prospects-verification.json`。已检查最终渲染、文字及表格内容、连接符、字体继承、分组锁、媒体回退和ZIP完整性。
未在桌面PowerPoint/WPS中打开编辑测试。
