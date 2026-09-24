# 渲染引擎与失败处理

## 默认行为

**Microsoft Word → WPS Writer → LibreOffice。** 默认是按这个顺序自动选择，不是强制只用 Word。没有安装、接口不可调用、启动失败、超时或导出失败时，记录具体原因，再尝试下一项。某项生成有效 PDF 和完整页图后停止切换；全部失败则阻断交付，不能跳过视觉验收。

此规则同时落实在预检、`render_docx.py`、流水线和最终放行入口。不得另行直接调用 LibreOffice 绕过默认顺序。引擎调用成功不代表文档合格；既有模板、内容保护、图表题注、编号、逐页逐图验收及失败修复规则不变。

## 调用

正常执行不必增加参数：

```bash
python scripts/review_pipeline.py input.docx --work-dir work
```

单独渲染也默认按同一顺序：

```bash
python scripts/render_docx.py input.docx --out-dir render
```

`--renderer word`、`--renderer wps`、`--renderer libreoffice` 表示本次**只允许指定引擎**，失败就报错，不再切换。这与默认优先级不是一回事。用户明确限定渲染器时，必须在后续 `--audit-only` 续跑中保留该参数。

渲染与更新域是两个独立动作。更新域默认 `--field-engine auto`，同样按 Word → WPS → LibreOffice 选择；独立刷新脚本用 `--engine`，不写成 `--renderer`：

```bash
python scripts/field_refresh.py repaired.docx --engine wps --out refreshed.docx --json-out refreshed-fields.json
```

仅明确要求域更新用 Word，不自动限定渲染器。两者都要求 Word 时使用 `--field-engine word --renderer word`。分别记录实际引擎；不同引擎的结果不能冒称一致，更不能把 WPS/LibreOffice 的结果写成 Word F9 实测。

## 执行环境

Word/WPS 的本次实现是 **Windows COM**：通过 PowerShell 调用 `Word.Application` / `KWPS.Application`。检测和调用发生在**脚本实际运行的电脑**；用户电脑安装了软件，不等于远程 Linux 容器能够调用。当前实现没有 macOS Word/WPS 自动化适配器。

三种引擎生成 PDF 后，均需要 `pdftoppm` 生成页面图片；可用时还通过 `pdfinfo` 核对 PDF 页数。Word/WPS 可用时不再强制安装 LibreOffice。LibreOffice 仅做 PDF 渲染不需要 UNO Python；使用它实际更新域时仍需要能导入 `uno` 的 Python，必要时通过 `--uno-python` 指定。

COM 能力以实际探测和调用为准，不仅查安装路径。WPS 的具体版本必须能够提供脚本所调用的对象和导出接口；接口不支持会记录失败并继续默认降级，不虚报成功。如果 COM 返回含用户已打开文档的共享实例，脚本拒绝操作该实例，防止关闭用户文档或改动交互会话。

探测与转换有超时。脚本正常结束会关闭本次打开的文档并恢复应用选项；异常退出后的残留 Office 进程不保证均可清理。不能用按进程名批量终止 WINWORD/WPS 的方式处理，以免关闭用户正在编辑的文档。

## 产物与验收

渲染读取独立输入副本，不保存回输入 DOCX。每次尝试使用独立临时目录，LibreOffice 另外使用隔离用户配置；失败尝试的产物不能进入验收。成功后发布到新的 `render-<随机标识>` 目录，不能扫描并混入旧页图。

`render-report.json` 记录实际引擎、每次尝试及原因、源 DOCX/PDF/全部 PNG 的哈希、页数和页面清单。渲染成功时 `visual_review_status` 仍是 `pending`，Agent 还必须完成最新页面的视觉验收。最终报告写出实际引擎和降级原因。

最终放行入口重新核对实际引擎、执行顺序、源文件、PDF、页图和原始渲染记录。旧清单缺少新证据时须重新渲染并验收，不能手改版本号或补假记录放行。

文档损坏、外部内容安全阻断、已确认的域结果错误、内容误改以及其他必需检查失败，不是通过切换引擎即可消除的问题。按主技能失败闭环处理，不以“另一引擎能够导出”为理由放行。

## 回归与本机实测

独立单元测试（原生 Office 调用使用模拟）：

```bash
python -m unittest discover -s tests -p test_renderer_priority.py -v
```

真实引擎测试会创建临时样例，执行三页渲染、缩为一页后重渲染、源文件与旧证据校验；它不是用户文档的视觉验收：

```bash
python tests/smoke_renderer.py --renderer auto --work-dir smoke-auto
python tests/smoke_renderer.py --renderer word --work-dir smoke-word
python tests/smoke_renderer.py --renderer wps --work-dir smoke-wps
```

Word/WPS 两条应在装有所需软件且可调用 COM 的 Windows 环境运行。真实样例测试的实际引擎和页图位置写入 `smoke-results.json`；不要把模拟测试通过写成本机接口已经实测通过。

## 接口参考

实现参考官方 API；具体安装环境仍以实测为准：

- Microsoft Word：`https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat`
- WPS Application：`https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Application/obj`
- LibreOffice 命令行：`https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html`
