# AGENTS.md

你为 M5Stack 的 AI 工程师服务。目标不是只给答案，而是把请求推进成可验证、可复用、可交付的工程结果。

## 基本工作方式

- 默认使用中文，除非用户明确要求英文或目标文件已有固定语言风格。
- 先检查真实文件、日志、配置和命令输出，再下结论；不要只按记忆或惯例猜测。
- 对简单问题直接给结论；对复杂改动先说明计划、影响面和验证方式，再执行。
- 保持改动最小且聚焦，避免无必要重构；若需要扩大范围，先说明原因和回归验证方式。
- 遇到用户短句追问时，切换为结论优先、少解释的回答方式。
- 文件路径在回复中使用完整可点击路径，尤其是 Windows 路径。

## 项目背景

本仓库是 Windows 上运行的 N9918A-Controller，用 Web 控制台承载 SA 频谱/EMI 测试和 NA 天线 S11 测量流程，控制 N9918A FieldFox、Mini-Circuits USB RF Switch，并支持频谱/天线数据处理、AI 异常分析和报告导出。项目范围以 Web 前端可见、可操作的流程为准。

关键文件：

- `web_app.py`: Flask API 和静态 Web 前端入口，默认监听 `127.0.0.1:5000`。
- `web_frontend/`: Web 控制台 HTML/CSS/JS，覆盖 SA/NA 模式入口、连接、配置、测量、校准、图表、峰值/谷值、AI 和报告操作。
- `sa_test_service.py`: Web API 使用的模式协调服务层，封装 SA/NA 状态、线程、结果、AI 和报告。
- `n9918a_backend.py`: SA/EMC 后端控制逻辑，使用 PyVISA/SCPI 配置仪器、采样、峰值搜索和数据保存。
- `n9918a_na_backend.py`: NA/S11 后端控制逻辑，使用 PyVISA/SCPI 配置 NA、执行 QuickCal、读取 `FDATa?`/`SDATA?`、计算谷值/带宽/Smith 数据。
- `Switch.py`: Mini-Circuits 切换器封装，通过 `pythonnet` 加载 `mcl_RF_Switch_Controller64.dll`。
- `chat.py`: AI 分析封装，使用 Volcengine Ark/OpenAI 兼容接口；密钥应来自环境变量。
- `utils/create_pdf.py`: ReportLab PDF 报告生成，依赖仓库字体或 Windows 中文字体。
- `doc/`: N9918A/FieldFox 官方说明资料，涉及 SCPI、规格或规则时优先对照这里。
- `run.bat`: 激活 `visa` conda 环境并启动 Web 控制台。

## Codex 任务四件套

处理工程任务时，尽量补齐这四项：

1. 目标：这次改动要达成的用户结果或工程结果。
2. 上下文：相关文件、日志、设备 IP、仪器模式、SDK/驱动、最近 diff 或复现步骤。
3. 约束：不要改什么、兼容边界、硬件状态、网络/API、资源限制、并发/UI 线程边界。
4. 验收：构建、静态检查、脚本、GUI/硬件验证步骤、风险和回滚方案。

## 本项目约束

- 运行环境以 Windows 为主；README 推荐 Python 3.8，Mini-Circuits DLL/`pythonnet` 也要注意 64-bit Python 兼容性。
- 不要删除或替换 `mcl_RF_Switch_Controller64.dll`、`assets/m5logo2022.png`、`utils/` 字体文件和 `doc/` 官方 PDF，除非用户明确要求并说明替代方案。
- 现有文件包含中文内容；写入中文文件必须使用 UTF-8，避免 PowerShell 默认编码导致乱码。
- 不要把 API key、token、Wi-Fi 密码、客户测试数据或私有日志写入代码、README、截图、报告或回答中。
- AI 分析密钥使用环境变量：优先 `ARK_API_KEY`，也可用 `VOLCENGINE_API_KEY` 或 `OPENAI_API_KEY`；模型和地址可用 `ARK_MODEL`、`ARK_BASE_URL` 覆盖。
- 生成的测试图、PDF、日志、缓存文件一般视为运行产物；除非任务要求，不要扩大提交范围到新的产物文件。
- 未在 Web 控制台、服务层或文档中引用的历史测试脚本、截图和临时入口默认视为可清理对象；清理前用 `rg` 确认没有引用。
- 仪器默认 IP、VISA resource 字符串、SCPI 模式切换和频段参数属于硬件行为配置，修改时要说明兼容影响。
- 不要复活旧的 `NA-mode/` 历史代码；NA 功能以 `n9918a_na_backend.py`、`sa_test_service.py` 和 Web API 中的当前实现为准。
- NA 预设、点数、IFBW、switchbox 位置和校准顺序属于硬件流程配置；修改前要对照 `doc/N9918A编程说明.pdf` 与用户确认的夹具路径。

## 硬件与安全边界

- 连接 N9918A、切换 RF Switch、执行测量、长时间采样、导出报告或调用外部 AI 服务，都可能改变设备状态或暴露数据；用户未明确要求时不要主动执行。
- 任何会改变仪器状态的命令必须先说明风险；涉及开关切换、长测、外部发送客户数据时，等待用户确认。
- `Switch.py` 的开关位置语义是 Mini-Circuits DLL 的 0/1 映射到位置 1/2；修改时必须确认 A/B/C/D 的物理链路含义。
- NA 校准固定顺序按 FieldFox QuickCal 文档执行：OPEN=`B2D1` + `CORR:COLL:INT 1;*OPC?` 必须先于 LOAD=`B1D1` + `CORR:COLL:LOAD 1;*OPC?`，最后 ANTENNA=`B2D2`；否则仪器可能返回 `Must acquire open port first`。
- SCPI 调试先做只读查询或短流程复现，例如 `*IDN?`、资源枚举、当前配置查询；避免直接上来改模式或触发扫描。
- 对超时、断连、VISA IO 错误、DLL 加载失败、RF Switch 未连接等情况，要保留清晰错误提示和降级路径。

## Web 与并发规则

- Web 前端只通过 `web_app.py` 的 API 改变设备状态；不要在浏览器 JS 中硬编码硬件流程。
- 耗时的连接、配置、测量、AI 分析、PDF 生成使用后台线程或服务层方法，避免阻塞 Web 请求。
- 测量过程中要正确禁用/恢复按钮，避免重复点击导致并发测量或仪器状态错乱。
- 新增状态字段时要检查失败、异常、取消和 finally 路径，确保 `measurement_in_progress` 和按钮状态可恢复。
- SA/NA 共用测量互斥状态，同一时间只允许一个模式执行测量或校准；校准期间必须禁用模式切换、测量和手动 switchbox 切换。
- 连接、配置、测量、AI 分析和导出 PDF 的错误提示要区分用户可处理问题与程序异常，避免只显示泛化失败。

## 数据、AI 与报告规则

- EMC 数据处理要保留频率、幅度、限值、Margin、Status 等字段语义，不要把单位或符号方向改错。
- NA 数据处理要同时保留 S11 dB、复数 Gamma、中心谷、所有候选谷、绝对阈值带宽和相对谷值带宽；全扫宽不显示 Smith Chart。
- AI 分析只应围绕异常点、Fail 点和临界 Margin 点展开；不要输出与异常无关的泛泛说明。
- PDF 报告生成要验证中文字体、logo、频谱图路径、项目信息、数据表和总结文本是否存在。
- 保存 CSV、PDF 或图像时要避免覆盖用户重要文件；如需覆盖，先提示文件路径和风险。
- 涉及客户名、EUT、型号、工程师、备注和测试结果时，按敏感数据处理，不要上传到外部服务，除非用户明确允许。

## 调试流程

- 不直接猜修法；按“现象分类 -> Top 3 假设 -> 证据 -> 最小实验 -> 修改”的顺序推进。
- 连接失败先分层：Python/依赖、VISA backend、网络/IP、防火墙、仪器模式、DLL/USB 驱动、权限。
- GUI 卡死先查耗时操作是否跑在主线程、线程回调是否安全、异常是否被吞掉。
- 测量结果异常先查 SCPI 配置、仪器模式、RBW/VBW/sweep time、频段点数、单位转换和后处理逻辑。
- PDF 失败先查字体注册、图片路径、输出路径权限、数据结构和 ReportLab 版本。
- AI 失败先查密钥来源、base URL/model、网络、异常日志和输入数据是否过大或包含不可发送内容。

## Code Review 关注点

- 硬件状态改变：SCPI 写命令、模式切换、开关位置、长时间测量是否受控。
- 并发与 UI：后台线程是否只通过主线程回调更新 UI，按钮状态是否可恢复。
- 内存与资源：VISA resource、DLL 对象、matplotlib figure、文件句柄是否关闭或复用合理。
- 错误处理：超时、断连、无设备、无字体、无密钥、无数据、PDF 路径异常是否有清晰提示。
- API 兼容：默认 IP、预设频段、数据字段、报告格式、README 操作流程是否保持兼容。
- 安全合规：是否引入硬编码密钥、外发客户数据、日志泄露或不可控文件覆盖。
- 测试缺口：是否有 py_compile、最小 smoke test、硬件验证步骤或可复现手动测试。

## 常用验证命令

仅修改文档时：

```powershell
$env:PYTHONIOENCODING='utf-8'
python -c "from pathlib import Path; text=Path('AGENTS.md').read_text(encoding='utf-8'); print(len(text), text.splitlines()[0])"
```

修改 Python 代码后优先跑语法检查：

```powershell
python -m py_compile n9918a_backend.py n9918a_na_backend.py sa_test_service.py web_app.py Switch.py chat.py utils/create_pdf.py tests/test_web_app.py
```

启动 GUI 做人工验证时：

```powershell
python web_app.py
```

Web API smoke test（不连接硬件）：

```powershell
python -c "from web_app import app; c=app.test_client(); print(c.get('/api/presets').json['ok']); print(c.get('/api/status').json['ok'])"
```

依赖/关键文件诊断（不连接硬件）：

```powershell
python -c "from web_app import app; c=app.test_client(); print(c.get('/api/diagnostics').json['data'])"
```

Demo 数据 smoke test（不连接硬件）：

```powershell
python -c "from web_app import app; c=app.test_client(); r=c.post('/api/demo/load', json={'duration_seconds': 15}).json; print(r['ok'], len(r['data']['peaks']))"
```

完整无硬件 Web smoke test：

```powershell
python -m unittest tests.test_web_app
```

前端语法检查：

```powershell
node --check web_frontend/app.js
```

## 输出标准

- 最终回答尽量包含：改了什么、为什么安全、验证跑了什么、还缺哪一步硬件/线上确认。
- 不能运行硬件验证时，说明原因并给出可复制的命令或人工步骤。
- 如果发现仓库已有未提交改动，不要回滚；只说明自己触碰的文件。
- 如果发现 API key、客户数据或私有日志，不复述敏感内容，直接提示用户轮换或清理。
- 推荐下一步时用编号列表，便于用户直接回复数字。
