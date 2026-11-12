# 南极磷虾油市场现状：全球增长与中国趋势

输出：`exports/krill-market-china-editable.pptx`，单页，画布1672×941。
标题、三栏内容、市场指标、需求驱动、底部发展关系和装饰均以独立对象重建；未将整页截图用作背景。
全球市场柱状图为原生PowerPoint图表，附内嵌Excel数据快照：2024年11亿美元、2030年21亿美元。
数据快照依据截图新建，不包含原始工作簿或原始公式。

国内指标保留截图中的“2024年1—9月”“国内线上销售额近10亿元”“同比增长24.3%”；人民币亿元与全球市场的亿美元分别保留。
2025—2030年CAGR 11.1%、预测值、营养及消费趋势文案按截图转录，未独立核验数据来源或重算。
雪山、水面、磷虾及图标为近似矢量重绘，不是原始照片；图标嵌入SVG并附1024×1024透明PNG回退。
语义箭头为2个原生连接符。照片装饰与截图存在视觉差异。

重复生成：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-krill-market-china.mjs
```

每次生成独立的`.local/krill-market-china/run-*/`目录，包含最终PPTX、渲染预览和验证回执。
脚本执行上游便携性处理、内置finalizer、原生图表及工作簿验证，并将最终文件重新导入渲染。
修改布局后，另对最终文件执行溢出检查：

```bash
"$RUNTIME_PYTHON" "$PRESENTATIONS_SKILL_DIR/container_tools/slides_test.py" /absolute/path/final.pptx
```

必须确认输出明确包含`Test passed. No overflow detected.`，不能仅依赖退出码。
本次核对了关键文字、图表缓存、Excel工作表数据、单位、连接符、SVG回退、分组锁、字体继承和ZIP完整性，并人工检查最终文件渲染。
版面检查无警告、溢出检查通过，验证记录见`docs/krill-market-china-verification.json`。
未在桌面PowerPoint/WPS中打开编辑测试。
