# N9 产品页：可编辑重建

输出：`exports/n9-product-editable.pptx`，画布 1676×938。
重建标题、四个卖点、三项优势、规格/食用方式/适用场景，以及右侧三盒产品包装。
三盒包装分别包含独立的原生顶面、侧面、正面、N9 标记和标签文字，可以单独编辑。
图标为嵌入 SVG，带透明 1024×1024 PNG 回退；Logo节点与线条为原生对象。

**包装是依据聊天截图近似重绘的可编辑插图，不是原始产品照片或官方印刷稿。**
N9 条纹字形、透视、反光、材料纹理和小字有简化。不可辨认的小字及配方不补写。
未将整张截图铺成背景，也没有使用整块包装照片代替可编辑对象。
来源图片中的功效、安全性、专利、营养师推荐及饮酒相关文案仅按图转录，没有独立验证或新增宣传断言。

运行：

```bash
cd /workspace/work
source scripts/ppt-runtime.sh
"$RUNTIME_NODE" scripts/rebuild-n9-product.mjs
```

每次输出到独立的 `.local/n9-product/run-*/output/`，预览及报告在该次运行的 `build/`。
本次报告见 `docs/n9-product-verification.json`。
检查包含官方 finalizer、重新导入、渲染、溢出、关键文案、三盒包装的独立对象、组合锁、字体继承、SVG 回退及外部图片关系。
本次未在桌面 PowerPoint/WPS 中打开文件，未验证包装与原产品照片逐像素一致。
