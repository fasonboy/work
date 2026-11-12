# 国内生产格局：新版单页

输出：`exports/production-landscape-editable.pptx`，单页1677×938。按本轮新版截图重建，不覆盖此前4页版。

标题、副标题、上游原料说明、四家企业名称/定位/介绍及底部项目说明全部为独立可编辑文本框，卡片、分隔线和冰山装饰为原生形状。6个图标为近似SVG重绘，附1024×1024透明PNG回退。没有整页截图背景，没有外部图片链接。原图没有数据表、图表或关系箭头，因此本页未引入这些对象。

保留新版关键区别：

- 辽渔集团全资**子**公司及陆上精深加工与销售定位。
- 上游挪威Aker BioMarine与VIK等品牌“存在原料及研发合作”的措辞。
- 辽渔具备捕捞、船载加工、陆上精深加工、科研平台与保健食品开发基础的原图陈述。
- 拟议56%磷脂/60粒SKU**不等同于**辽渔2026年公开的60%新品，仍须单独核验规格、产品属性、备案/注册、委托生产关系及批次检测报告。

企业产能、许可、合作、产品等信息按当前截图转录，未独立核验。船名按当前图面转录为“福六海”，未用此前版本或外部资料静默替换；正式使用时应核对企业原始材料。图标与右上小幅冰山装饰是近似重绘，保留相近的颜色与位置，不能视为原始照片。

重复生成：

```bash
cd /workspace/work
bash scripts/check-ppt-environment.sh
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-production-landscape.mjs
```

每次生成新的`.local/production-landscape/run-*/`目录，保留最终文件、预览及验证回执。脚本使用平台Artifact Tool，通过上游便携性处理、官方finalizer、字体策略及重新导入渲染；布局修改后对最终文件执行：

```bash
"$RUNTIME_PYTHON" "$PRESENTATIONS_SKILL_DIR/container_tools/slides_test.py" /absolute/path/final.pptx
```

确认日志明确出现`Test passed. No overflow detected.`且没有`ERROR:`，审阅后再将最终文件复制到`exports/`。

本次已实际复测依赖及功能生成，检查22项关键文案、页面尺寸、ZIP完整性、编辑格式、组合锁、外部媒体、6组透明PNG回退，并审阅最终Artifact Tool预览和LibreOffice独立渲染。布局错误/警告均为0，溢出检查通过。未在桌面PowerPoint/WPS中实际交互编辑。详见`docs/production-landscape-verification.json`。
