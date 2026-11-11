# 鱼油与磷虾油市场机会：可编辑重建

输出：`exports/fish-krill-opportunity-editable.pptx`，画布 1672×941。
时间线、文字、数字与原生连接符为独立对象。Logo、图标、雪山、冰面和胶囊近似重绘，不是原始图片资产。

右侧比较内容为 **4行3列的原生PowerPoint表格**，可以直接改文字、增删行列、修改单元格填充与边框。
没有用独立矩形网格代替数据表格。图标与点赞徽章独立位于表格上方，移动表格时需要按需一起选择这些装饰。

左侧柱状图为原生图表，内嵌新的 Excel 数据快照：

| 类别 | 数据值 | 显示 |
| --- | ---: | ---: |
| 1999–2000年 | 0.013 | 1.3% |
| 2011–2012年 | 0.12 | 12.0% |

原始数值使用比例小数，以 `0.0%` 格式显示；图表缓存与内嵌工作簿单元格逐项核对一致。
数据快照由截图中的完整字面值创建，不声称保留原始工作簿或公式血缘。
“约9倍”、时间线、营养对比和市场结论按图保留，没有独立核验统计来源或营养功效，也没有新增外部数据来源。

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-fish-krill-opportunity.mjs
```

每次输出到独立的 `.local/fish-krill-opportunity/run-*/output/`，预览及报告在对应 `build/`。
已执行官方 finalizer、原生表格/图表检查、内嵌单元格对照、重新导入、渲染及页面溢出检查，报告见 `docs/fish-krill-opportunity-verification.json`。
本次未在桌面 PowerPoint/WPS 中打开测试，装饰与参考照片存在差异。
