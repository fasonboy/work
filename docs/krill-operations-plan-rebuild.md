# 线上线下一体化运营方案：5张截图重建

输出：`exports/krill-operations-plan-editable.pptx`，5页，每页1672×941。
按用户本轮附件发送顺序保留：

1. 项目整体规划：首期资金140–150万元，产品采购70万、市场推广70–80万，以及产品、私域和渠道资源。
2. 项目实施与经营目标：市场启动1–3个月、渠道增长4–6个月、规模运营7–12个月，以及四项经营指标。
3. 资金投入与推广计划：推广预算70–80万元，约60%线上、约33%线下、约7%机动储备。
4. 线下运营方案：区域经销商、实体门店、健康机构、企业团购及北京、安徽试点策略。
5. 线上运营方案：电商平台、社交内容、微信私域及传播到复购的流程。

原图页眉中的“第4页”“第5页”“03”“2 /”是截图内部标签，按图保留；实际文件页序按附件顺序，不擅自重排。
文字、数字、卡片和产品示意均为独立对象；语义箭头为原生连接符。预算为原生环形图，内嵌依据60/33/7新建的Excel数据快照，不声称保留原始工作簿或原始公式。
预算与比例按截图保留，没有推算或新增分项金额、经营目标数或商业预测。所有产品属性和市场主张按图转录，未独立核验。

原始截图文件未在云端提供；瓶装产品、雪山、地图、门店、机构、团购及平台标识均为近似矢量重绘，不能作为实拍或官方Logo原文件使用。
第1页标明“产品示意重绘”，第4页四块场景标明“场景示意”。文档没有将整页截图嵌为背景。
SVG图标内嵌高分辨率透明PNG回退，不依赖外部媒体文件。

重复生成：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-krill-operations-plan.mjs
```

每次输出位于独立的`.local/krill-operations-plan/run-*/`目录，含最终PPTX、逐页预览和验证回执。
运行时依赖可用`bash scripts/check-ppt-environment.sh`实际复测。
生成器执行上游便携性处理、官方finalizer、图表数据与字体验证，并将最终文件重新导入渲染。
修改布局后需再次对最终文件运行溢出测试：

```bash
"$RUNTIME_PYTHON" "$PRESENTATIONS_SKILL_DIR/container_tools/slides_test.py" /absolute/path/final.pptx
```

确认输出明确包含`Test passed. No overflow detected.`，不能仅依赖退出码。
当前Artifact Tool预览器未显示环形图；独立通过LibreOffice导出PDF并检查最终文件全部5页，验证环形图显示和标签位置。
最终文件的文字、数据工作簿、字体继承、分组锁、连接符、媒体回退、ZIP完整性及版面检查记录见`docs/krill-operations-plan-verification.json`。
未在桌面PowerPoint/WPS中打开编辑测试。
