# N9918A Web SA Frontend Design

## 目标

把已实现的 SA/EMC 测试流程整理为 Web 控制台：连接仪器、连接 RF Switch、选择预设频段、自动切换、单次扫描、15s/5min EMI 采样、停止、保存数据、AI 分析、PDF 报告导出和现场诊断。

## 方案

采用轻量 Flask + 原生 HTML/CSS/JS：

- `web_app.py`: HTTP API 和静态文件服务。
- `sa_test_service.py`: 测试流程服务层，统一管理 `N9918AController`、`MiniCircuitsSwitchController`、测量线程、当前数据、AI 结果和报告。
- `web_frontend/`: 浏览器控制台，不直接接触硬件，只调用 API。
- `n9918a_backend.py`: 保留已验证的 PyVISA/SCPI 设备控制和数据处理。

这样不重写底层硬件控制，降低回归风险；Web 层只替换用户体验和流程编排。

界面主文案使用中文，并在连接配置、RF Switch、测量流程、频谱、峰值、报告、AI 和日志模块旁提供 `?` 帮助提示，降低首次使用和现场排障成本。

## 功能映射

| 用户流程 | Web 实现 |
| --- | --- |
| 设备 IP 连接/断开 | `/api/device/connect`, `/api/device/disconnect` |
| 测试配置 + 应用配置 | `/api/presets`, `/api/configure` |
| RF Switch 连接/状态/设置 | `/api/switch/*` |
| 单次扫描 | `/api/measure/single` |
| 15 秒 / 5 分钟采样 | `/api/measure/timed` |
| 停止测量 | `/api/measure/stop` |
| 保存数据 | `/api/data/save` |
| AI 异常分析 | `/api/ai/analyze` |
| 导出 PDF 报告 | `/api/report/export` + download URL |
| 无硬件预览 | `/api/demo/load` 生成模拟频谱、峰值和 AI 文本 |
| 现场诊断 | `/api/diagnostics` 检查依赖、关键文件和 AI 环境变量 |

## 文档核对结果

参考 `doc/N9918A编程说明.pdf`：

- 官方 SA 示例在切换模式时使用 `INST:SEL 'SA';*OPC?` 等待完成。
- 单次扫描示例使用 `INIT:IMM;*OPC?` 后读取 `TRACE:DATA?`。
- SA/PAA/NF 模式读取数据使用 `TRACe:DATA?`。
- 频率、点数和带宽命令族与当前代码使用的 `SENS:FREQ:START/STOP`、`SENS:SWE:POIN`、`SENS:BAND...` 一致。

据此做了小幅安全改动：连接时等待 SA 模式切换完成，单次扫描优先使用 `:INIT:IMM;*OPC?`，长时间采样继续沿用已投入使用的连续扫描读取 trace 流程，并新增停止回调，以便 Web `停止测量` 按钮更可靠地退出采样循环。

## 验证计划

不连接硬件的本地验证：

```powershell
python -m py_compile n9918a_backend.py sa_test_service.py web_app.py Switch.py chat.py utils/create_pdf.py tests/test_web_app.py
python -c "from web_app import app; c=app.test_client(); print(c.get('/api/presets').json['ok']); print(c.get('/api/status').json['ok'])"
python -c "from web_app import app; c=app.test_client(); print(c.get('/api/diagnostics').json['ok'])"
python -c "from web_app import app; c=app.test_client(); r=c.post('/api/demo/load', json={'duration_seconds': 15}).json; print(r['ok'], len(r['data']['peaks']))"
python -m unittest tests.test_web_app
```

需要硬件的人工验证：

1. `python web_app.py` 后打开 `http://127.0.0.1:5000`。
2. 连接 N9918A，确认状态变为已连接。
3. 连接 RF Switch，确认 A/B/C/D 状态读取正常。
4. 选择每个 Test Config，确认仪器配置和自动切换位置符合预期。
5. 先跑单次扫描，再跑 15 秒 EMI，确认曲线、峰值表、保存数据、AI 分析和 PDF 下载。
6. 最后再跑 5 分钟 EMI，检查停止测量、长测稳定性和报告生成。
