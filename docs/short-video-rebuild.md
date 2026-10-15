# 超级短视频总结：截图重建记录

根据当前聊天中可见的截图重建为一页 PowerPoint，保留 1694×928 的原始比例。
文字、指标、进度条、卡片、图标及连接线为独立对象；未将整页截图铺成背景。
主字体为云端可用的 Noto Sans CJK SC，与原图字体并非经过源文件确认的同一字体。

本次环境只拿到了聊天预览，工作目录中没有截图原文件。因此当前输出是**待补六个视频缩略图的草稿**，而不是完整复原版。
观看量、点赞量和六张卡片下方的标题已转为可编辑文字；缩略图中原有的照片及图内文字尚未恢复。

生成源文件：`scripts/rebuild-short-video.mjs`。

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-short-video.mjs
```

获得原始截图文件后可重新生成，程序只裁取右侧六个画面，不嵌入整张原图：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-short-video.mjs --reference /absolute/path/screenshot.png
```

输入需与示例截图布局、比例一致。生成器会读取实际尺寸并检查比例，然后按参考坐标裁取画面。
输出在独立的 `.local/short-video-rebuild/run-*/output/`；构建报告和 PNG 在该次运行的 `build/`。
每次运行写入新目录，不覆盖原始截图或既有 PPTX。
按照 Presentations 技能要求，在每项独立制作任务的首次 authoring 前记录一次操作；`PPT_OPERATION_MARKED=1` 仅供同一任务内已记录操作后的修订使用。

已对 revision-3 草稿执行后处理、官方 finalizer、重新导入、渲染和独立 OOXML 检查：保留关键数值，66 个可编辑文本对象，3 个原生箭头连接符，30 个 SVG 均有透明 1024×1024 PNG 回退，无外部图片关系，字体继承及组合锁检查通过。
未在桌面 PowerPoint 或 WPS 中打开验证。补入截图后必须重新检查最终文件和全部页面；当前草稿的验证结果不能代替新输出的验证。
