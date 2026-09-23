# 最终交付入口

只交付 `finalize_review.py` 输出的 Word 和实际修复报告，不能把 candidate 改名直接交付。

新版 `review_pipeline.py` 生成版本4清单，必须包含标题编号、题注、文字、模板、表格、真实图片对象和实际更新域的检查结果。Agent 逐页、逐图填写最新观察，逐项关闭表格 review 后，执行：

```bash
python scripts/finalize_review.py work/review-manifest.json final-visual-review.json --out final.docx --json-out release.json
```

清单/页面/原图/候选/模板哈希不匹配、漏检、图内缺陷未关闭、字段刷新无真实执行记录，均不输出交付文件。任何后续修改使原验收失效，须重新更新域、审计和渲染。失败不覆盖已有成果。

这不是宿主级不可绕过的 hook。若宿主允许 Agent 任意复制文件并直接回复，单靠 Skill/脚本无法禁止绕过；宿主需把“可交付文件”的权限绑定到此入口。图片判断来自真实看图，不能以可编辑 JSON 代替观察。
