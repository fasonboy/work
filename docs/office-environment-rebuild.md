# 办公环境实拍展示：待补照片的可编辑版式

输出：`exports/office-environment-editable-draft.pptx`，画布 1672×941。
这是**缺少两张实拍原照片的草稿，不是完整转换结果**。标题、办公区/会议与接待区标签、底部四项说明、Logo和蓝色装饰已重建为独立对象。
当前云端能看到聊天预览，但没有可读取的原始图片文件；没有用重绘或生成的办公室照片替代实拍。
Logo与图标是近似矢量重绘，图标内嵌1024×1024透明PNG回退。

生成版式草稿：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-office-environment.mjs
```

收到两张独立原照片后，可以原样内嵌（保持比例，裁切适应框位）：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-office-environment.mjs \
  --office-image /absolute/path/office.jpg \
  --reception-image /absolute/path/reception.jpg
```

也支持从与聊天图片布局及比例相同的整页截图裁取两块照片：

```bash
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-office-environment.mjs --reference /absolute/path/screenshot.png
```

从整页截图裁切时，照片中原来烘焙的蓝色标签也会随图片保留；上方另有可编辑标签。独立无标注原照片可以避免这个问题。
照片内部的家具、建筑等仍是栅格图像，不会自动变为可编辑形状。
原图可作为ZIP附件发送，或放进GitHub仓库后提供路径。用户原照片补入分支本次未实测，因为原文件缺失。

每次输出位于独立的 `.local/office-environment/run-*/output/`；只在两张照片均提供后才输出不带draft的文件名。
草稿经过官方finalizer、重新导入、渲染、关键文字、组锁及图片便携性检查；报告见 `docs/office-environment-verification.json`。
本次未在桌面PowerPoint/WPS中打开编辑测试。原照片补齐后必须重新验证新输出。
