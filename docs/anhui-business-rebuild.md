# 安徽业务展示：新版截图重建

输出：`exports/anhui-business-editable-draft.pptx`。这是**缺少两张原始门店照片的版式草稿**。
当前聊天能显示图片预览，但云端没有可供裁切的截图文件；未用其他照片或生成图片冒充原照片。
地图是依据可见内容近似重绘的示意，部分细节省略，不能视为原地图或用于导航。

保留 1672×941 画布。标题、副标题、正文、三组卡片、金色标签、地址气泡、核心亮点及蓝金波浪均为独立对象。
图标为嵌入 SVG，并带透明 1024×1024 PNG 回退。照片补齐后是可单独替换的图片，照片内的招牌、建筑和家具并非可编辑对象。

运行：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-anhui-business.mjs
```

获取新版截图原文件后可补入原照片：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-anhui-business.mjs --reference /absolute/path/screenshot.png
```

图片须与新版截图的布局和比例一致。程序仅裁切两张照片，不将整页截图作为背景。
可通过原始文件附件、包含原图的 ZIP，或上传到 GitHub 仓库后提供文件路径来传入。
新输出存入独立的 `.local/anhui-business/run-*/output/`，原图保持不变。

已执行官方 finalizer、重新导入、渲染、结构及溢出检查。当前验证记录见 `docs/anhui-business-verification.json`。
未在桌面 PowerPoint/WPS 中打开测试；补齐原照片后的输出必须重新验收。
