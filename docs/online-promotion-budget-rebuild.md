# 线上推广费用规划

输出：`exports/online-promotion-budget-editable.pptx`，单页，画布1672×941。
按当前截图重建标题、副标题、左侧预算圆框、四项费用卡片和底部投放原则。
所有文字、金额及百分比使用原生文本框，背景、圆环、卡片和波浪使用原生形状，关系箭头为原生连接符。
图标是近似矢量重绘，内嵌SVG及1024×1024透明PNG回退。没有使用整页截图背景。

截图数据保持原样：线上预算24.5–28万元，占总推广预算约35%；线上平台广告45%、内容种草合作25%、短视频及直播20%、私域会员运营10%。
四项费用占比合计100%，是线上预算内部的分配；35%是总推广预算占比，两种口径分别保留。
左侧完整圆框是装饰，不作为35%的进度图。本页没有数据图表或原生表格，也没有把四项卡片改成表格。
本次生成独立单页，不修改前一份5页方案中“线上推广约60%”的原图数字；两份截图之间存在口径差异，未擅自统一。
没有另行推算分项金额、重算占比或验证预算决策。

重复生成：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-online-promotion-budget.mjs
```

每次运行生成独立的`.local/online-promotion-budget/run-*/`目录，含最终PPTX、预览和验证回执。
脚本执行上游便携性处理、内置finalizer、字体和结构检查，并重新导入最终PPTX渲染。
布局修改后需对最终文件再运行溢出测试，明确确认成功日志：

```bash
"$RUNTIME_PYTHON" "$PRESENTATIONS_SKILL_DIR/container_tools/slides_test.py" /absolute/path/final.pptx
```

本次实际复测了运行时依赖和功能生成，核对了最终文字、费用占比、原生连接符、分组锁、字体继承、SVG回退及ZIP完整性，并人工检查最终渲染。
验证记录见`docs/online-promotion-budget-verification.json`。未在桌面PowerPoint/WPS中打开编辑测试。
