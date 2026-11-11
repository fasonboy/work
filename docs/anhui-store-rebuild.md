# 安徽线下门店展示：重建记录

用户提供的图片在聊天中可见，但本次云端没有收到可读取的截图文件。
因此当前交付的是**待补两张门店照片的版式草稿**，不能视为完整复原版。

- 保留 1672×941 原图比例及三栏布局。
- 标题、底部三组说明、卡片、图标和分隔线为独立对象。
- 地图使用原生形状和可编辑文字近似重绘，保留截图可见的主要地名和道路；其位置及细节存在近似，不能作为导航地图或用于核对地址。
- 照片区域明确标记“待补截图原文件”，没有使用其他门店照片替代。
- 字体使用当前环境提供的 Noto Sans CJK SC，并未确认与截图原字体完全相同。

生成脚本：`scripts/rebuild-anhui-store.mjs`。

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-anhui-store.mjs
```

可将截图原文件上传到项目任意目录，例如 `assets/anhui-store.png`，然后运行：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-anhui-store.mjs --reference assets/anhui-store.png
```

原图需与聊天截图布局及比例一致。程序只裁取两张照片并内嵌，不把整张截图作为 PPT 背景。
补入照片后，照片本身仍是图片对象；照片内的建筑、家具和招牌文字不因此变成可编辑形状。

本次检查：30 个可编辑文本框；3 个 SVG 图标均带 1024×1024 PNG 回退；ZIP 完整性、关键文字保留、无外部图片关系、组合锁、字体继承、官方 finalizer、重新导入和渲染通过。
当前运行的 PNG 与详细报告保存在被忽略的 `.local/anhui-store/revision-1/build/`。本次未在桌面 PowerPoint/WPS 中打开文件。补入照片后须对新输出重新验收。
