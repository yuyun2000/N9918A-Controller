# N9918A-Controller

基于 Keysight N9918A FieldFox 的 SA/NA Web 控制台，集成 Mini-Circuits RF Switch 自动切换、SA 单次频谱采集、15s/5min EMI 采样、NA 天线 S11/回波损耗/驻波比测量、Fast Cal 自动校准、峰值/谷值判定、AI 异常分析和报告导出。

当前入口是 Web 前端：`web_app.py` + `web_frontend/`。页面顶部可选择 `SA 频谱测试` 或 `NA 天线测量`；仓库只保留 Web 控制台正在使用的流程和必要支撑文件，不再复用旧的 `NA-mode/` 历史代码。

## 系统要求

- 操作系统：Windows
- Python：推荐 3.8，与 `pythonnet`、Mini-Circuits 64-bit DLL 保持一致
- 硬件：
  - Keysight N9918A FieldFox
  - Mini-Circuits USB RF Switch
  - USB 连接线与仪器网络连接

## 安装

```powershell
pip install -r requirements.txt
```

如需使用 AI 分析，请在本地配置密钥，不要写入可提交代码。当前 AI 客户端使用 OpenAI-compatible Responses API，默认模型为 `gpt-5.5`，默认地址为本地代理 `http://192.168.20.38:3000/`，默认 reasoning effort 为 `xhigh`。

推荐在仓库根目录创建本地忽略文件 `ai_config.local.json`：

```json
{
  "api_key": "你的密钥",
  "base_url": "http://192.168.20.38:3000/",
  "model": "gpt-5.5",
  "reasoning_effort": "xhigh"
}
```

也可以临时用环境变量覆盖：

```powershell
$env:N9918A_AI_API_KEY="你的密钥"
$env:N9918A_AI_BASE_URL="http://192.168.20.38:3000/"
$env:N9918A_AI_MODEL="gpt-5.5"
$env:N9918A_AI_REASONING_EFFORT="xhigh"
```

## 启动 Web 前端

```powershell
python web_app.py
```

启动后默认会自动打开浏览器。如果浏览器没有弹出，可手动打开：

```text
http://127.0.0.1:5000
```

如果暂时没有连接硬件，可以点击 Web 页面里的 `加载演示数据`，加载一组模拟频谱、峰值和 AI 分析文本，用于检查页面布局、图表、峰值表和报告按钮状态。演示数据不会访问 N9918A 或 RF Switch，不能替代真实硬件验收。

如果你创建了名为 `visa` 的 conda 环境，也可以双击默认启动脚本：

```powershell
.\run.bat
```

`run.bat` 会启动 Web 服务，并自动打开 `http://127.0.0.1:5000`。

## Web 控制台功能

页面主控件已中文化，模块标题旁的 `?` 会提示配置含义、推荐操作顺序、数据外发风险、校准要求和排障要求。

1. 连接 N9918A：输入设备 IP，默认 `192.168.20.233`。
2. 连接 RF Switch：可查看型号、SN、固件、温度、USB 状态和 A/B/C/D 位置。
3. 选择 Test Config 并配置仪器：
   - `EMC_30MHz_1GHz`
   - `LF_9kHz_150kHz`
   - `MF_150kHz_30MHz`
   - `HF_1GHz_3GHz`
4. 自动切换 RF Switch：
   - `< 30MHz`: A2 + D2，B/C 保持 1
   - `30MHz - 3GHz`: A2 + D1，B/C 保持 1
   - 其他范围：A/B/C/D 回到 1
5. 测量：
   - `加载演示数据`: 加载无硬件演示数据，用于 UI/流程预览
   - `手动清理 Trace`: 停止连续扫描、清理 SCPI 状态、清写/置空 Trace，并清空 Web 端上一轮 SA 图表和峰值
   - `单次扫描`: 单次快速扫描，用于先确认频谱是否正常；流程会清写 Trace，等待 `INIT:IMM;*OPC?` 完成后读取 `TRAC1:DATA?`
   - `15 秒采样`: 15 秒 EMI 筛查采样，按仪器 sweep time 逐次触发完整 sweep，支持 AI 和 PDF
   - `5 分钟采样`: 5 分钟 EMI 筛查采样，支持 AI 和 PDF
   - `停止测量`: 请求停止当前采样，并发送 `INIT:CONT OFF`
6. 数据与报告：
   - `保存数据`: 保存原始采样、峰值和频谱 CSV
   - `AI 异常分析`: 对超限/临界 Margin 频点做异常分析
   - `导出 PDF 报告`: 生成报告并提供浏览器下载链接
   - `环境诊断`: 检查关键 Python 包、DLL、logo、字体、doc PDF 和 AI 环境变量

### NA 天线测量流程

1. 顶部切到 `NA 天线测量`；仪器已连接时会发送 `INST:SEL "NA";*OPC?`。
2. 选择预设并应用：
   - `ANT_315`: 215-415MHz
   - `ANT_433`: 300-500MHz
   - `ANT_868`: 768-968MHz
   - `ANT_915`: 815-1015MHz
   - `ANT_2450`: 2.2-2.7GHz
   - `ANT_5G`: 4.8-6.0GHz
   - `ANT_FULL`: 30kHz-26.5GHz，全扫宽只分页列出谷值，不显示 Smith Chart
3. 自动校准必须同时连接 N9918A 与 switchbox，顺序按 FieldFox QuickCal 文档固定为：
   - OPEN: `B2C1`，先执行 `CORR:COLL:METH:QCAL:CAL 1`，再执行 `CORR:COLL:INT 1;*OPC?`
   - LOAD: `B1C1`，执行 `CORR:COLL:LOAD 1;*OPC?`
   - SAVE/ANTENNA: 执行 `CORR:COLL:SAVE 0`，再切到 `B2C2`
4. 测量读取 `CALC:DATA:FDATa?` 得到 S11 dB，读取 `CALC:DATA:SDATA?` 得到复数 Gamma；普通预设显示 S11 曲线、VSWR 曲线、实际中心谷、理想频点对比、理想频点附近损耗、回波损耗、驻波比、绝对/相对 3dB/10dB 带宽和带 50Ω 阻抗网格的 Smith Chart。
5. 带宽有两套口径：绝对阈值 `S11 <= -3dB/-10dB`，相对谷值 `S11 <= valley+3dB/valley+10dB`，端点用线性插值估算。
6. NA 报告导出为 A4 纵向 PDF，包含 M5Stack logo、项目/工程师信息、理想频点与实际中心谷偏移、端点 S11、回波损耗、驻波比、S11 曲线、VSWR 曲线、Smith Chart 阻抗标记和谷值列表。

## 代码结构

```text
n9918a_backend.py       # N9918A SA/EMC PyVISA + SCPI 控制和数据处理
n9918a_na_backend.py    # N9918A NA/S11 控制、校准流程、谷值/带宽/Smith 数据处理
sa_test_service.py      # Web API 使用的测试流程服务层
web_app.py              # Flask API 与静态资源服务
web_frontend/           # Web 控制台 HTML/CSS/JS
run.bat                 # 默认启动 Web 控制台
Switch.py               # Mini-Circuits RF Switch 封装
chat.py                 # AI 分析客户端
utils/create_pdf.py     # PDF 报告生成
doc/                    # N9918A/FieldFox 官方资料
assets/m5logo2022.png   # PDF 报告 logo
```

## SCPI 流程依据

`doc/N9918A编程说明.pdf` 中的官方示例说明：

- SA 模式切换使用 `INST:SEL 'SA';*OPC?` 等待完成。
- 单次扫描示例使用 `INIT:IMM;*OPC?` 后读取 `TRACE:DATA?`。
- SA/PAA/NF 模式读取数据使用 `TRACe:DATA?`。
- 频率范围、点数、带宽使用 `SENS:FREQ:START`、`SENS:FREQ:STOP`、`SENS:SWE:POIN`、`SENS:BAND...` 命令族。
- NA 模式切换、S11 配置和读取使用 `INST:SEL "NA";*OPC?`、`CALC:PAR:DEF S11`、`CALC:FORM MLOG`、`INIT:IMM;*OPC?`、`CALC:DATA:FDATa?` 和 `CALC:DATA:SDATA?`；编程手册也支持 `CALC:FORM SMITh/SWR` 及 `MMEM:STORe:IMAGe` 保存当前仪器屏幕 PNG，但本项目报告默认用 `SDATA?` 自绘 Smith/VSWR，便于统一标注和离线验证。
- QuickCal/校准采集使用 `CORR:COLL:METH:QCAL:CAL 1`、`CORR:COLL:INT 1;*OPC?`、`CORR:COLL:LOAD 1;*OPC?`、`CORR:COLL:SAVE 0`，OPEN/INT 必须先于 LOAD。

当前代码已按这些要点保持流程：连接时等待 SA 模式切换完成；SA 配置/采样前设置 `INIT:CONT OFF`、`:SENS:AMPL:UNIT DBUV`、`:TRAC1:TYPE CLRW` 并关闭/清理其他历史 trace；单次扫描和长时间采样都逐次执行 `:INIT:IMM;*OPC?` 后读取 `:TRAC1:DATA?`，采样间隔不快于仪器 `:SENS:SWE:TIME?`。页面也提供 `手动清理 Trace`，用于现场主动清空仪器屏幕旧曲线和 Web 端旧结果；下一次扫描仍会自动重新进入 Clear/Write。

SA 结果默认是筛查口径，不等同正式 FCC/CE 合规报告：

- `PEAK` 是同步 sweep 结果的最大值。
- `AVERAGE` 在电压域求筛查平均，不在 dB 域直接平均。
- `QUASI_PEAK` 是软件估算准峰值；若仪器安装 EMI Option 361，应优先使用仪器 EMI/QPD/EAV 检测器做正式确认。
- FCC/CE 曲线为带 Class、距离、Detector 和单位说明的筛查参考；正式判定必须确认天线因子、线缆损耗、switchbox 损耗、外置前置放大器增益、测试距离、场地和标准版本。

## 注意事项

- 不要删除 `mcl_RF_Switch_Controller64.dll`，它是切换器控制库。
- PDF 导出仅面向 15s/5min 等 EMI 测量结果；单次扫描用于快速确认。
- NA 报告导出当前生成 HTML 报告和 CSV/JSON 数据；NA v1 不调用 AI 分析。
- 未经用户确认，不要在真实硬件上执行 NA 校准、switchbox 切换或长扫频。
- AI 分析会通过 Responses API 把峰值表和测试信息发送到配置的 AI 服务，涉及客户数据时需先确认可外发；密钥可放在已忽略的 `ai_config.local.json`，不要写入可提交源码、README、报告或截图。
- 生成的 PDF、CSV、日志和缓存属于运行产物，默认不纳入版本控制。

## 验证

代码语法检查：

```powershell
python -m py_compile n9918a_backend.py n9918a_na_backend.py sa_test_service.py web_app.py Switch.py chat.py utils/create_pdf.py tests/test_web_app.py
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
