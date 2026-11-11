# 台湾加康诺：最新简洁版

输出：`exports/jkn-brand-simple-editable-draft.pptx`。按用户最新的简洁版截图更新，保留 1672×941 画布。
公司介绍、Logo、标题色带、四组说明及浅蓝销售额卡片均可独立编辑。第四组文字为“市场表现稳定”“深受消费者信赖，持续增长”。
“1,000万 台币”与“截至2026年10月”按图保留，未独立核验数据。Logo和图标为近似重绘。

当前云端没有可读取的原始截图，两份登记文件原图尚未嵌入，右侧明确标为待补，不是完整复原版。
未仿造印章或扫描件，未猜写模糊文件内容。原始文件可通过 ZIP 附件或放入 GitHub 仓库提供。

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-jkn-brand-simple.mjs
```

获得与最新截图布局相同的原始图片后，补入原件：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-jkn-brand-simple.mjs --reference /absolute/path/screenshot.png
```

程序只裁切两份登记文件，不把整页截图当作背景。嵌入后扫描件仍是图片对象，内部文字与印章并非可编辑元素。
每次输出到独立的 `.local/jkn-brand-simple/run-*/output/`，不覆盖用户原图。
本次运行经过官方 finalizer、重新导入、渲染及独立结构检查，验证记录见 `docs/jkn-brand-simple-verification.json`。
未在桌面 PowerPoint/WPS 中打开测试；补入原件后的新输出需要重新验收。
