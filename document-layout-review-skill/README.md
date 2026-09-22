# document-layout-review

通用 Word/DOCX 排版检查与修复 Skill。

## 结构

- `SKILL.md`：Agent 的工作流、决策规则和修复闭环
- `references/`：模板、表格、流程图、文字和验收细则
- `scripts/`：确定性检查、修复、渲染工具
- `assets/`：默认模板和已经确认的固定样式 JSON

## 核心闭环

**建立排版基准 → 全面检查 → 问题分类 → 修复 → 重新渲染 → 验收 → 未通过继续修复**

发布前运行：

```bash
python scripts/verify_skill_package.py .
```
