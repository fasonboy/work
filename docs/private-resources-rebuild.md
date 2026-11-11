# 线下私域资源展示：可编辑重建

输出：`exports/private-resources-editable.pptx`，画布 1676×939。
保留资源卡片、运营支持、社群示例的三栏关系。标题、说明、四个原生箭头连接符及手机列表文字可编辑。
设备、场景、手机框和群聊列表以原生形状与文本近似重绘；图标为 SVG，内嵌透明 1024×1024 PNG 回退。

**本文件不是原始照片及聊天截图的无损转换。**
当前环境只有可见的聊天预览，没有原始截图文件，故血压计、恒温箱和体检场景使用示意插图。
血压计显示横线，不写入无法确认的测量读数。房间纹理、设备细节与原照片不同。
群聊保留“示例”标题，仅转录可辨认的名称与截断预览；头像使用简化色块，不补全未知消息或推断人物身份。
标题下方的截图校对红线未作为演示文稿设计元素保留。

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-private-resources.mjs
```

每次输出位于独立的 `.local/private-resources/run-*/output/`，预览与报告在对应 `build/`。
本次验证记录见 `docs/private-resources-verification.json`，包含官方 finalizer、重新导入、渲染、页面溢出、关键文案、原生对象、SVG 回退及外部图片关系检查。
本次未在桌面 PowerPoint/WPS 中打开文件；验证通过不表示与原照片或截图逐像素相同。
