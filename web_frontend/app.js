const state = {
  presets: {},
  result: null,
  busy: false,
  wasMeasuring: false,
};

const $ = (id) => document.getElementById(id);

const elements = {
  statusDot: $("statusDot"),
  statusTitle: $("statusTitle"),
  statusDetail: $("statusDetail"),
  deviceIp: $("deviceIp"),
  presetSelect: $("presetSelect"),
  connectBtn: $("connectBtn"),
  disconnectBtn: $("disconnectBtn"),
  configureBtn: $("configureBtn"),
  connectSwitchBtn: $("connectSwitchBtn"),
  refreshSwitchBtn: $("refreshSwitchBtn"),
  disconnectSwitchBtn: $("disconnectSwitchBtn"),
  switchGrid: $("switchGrid"),
  switchReadout: $("switchReadout"),
  demoBtn: $("demoBtn"),
  singleBtn: $("singleBtn"),
  slowBtn: $("slowBtn"),
  fastBtn: $("fastBtn"),
  stopBtn: $("stopBtn"),
  saveBtn: $("saveBtn"),
  analyzeBtn: $("analyzeBtn"),
  pdfBtn: $("pdfBtn"),
  diagnosticsBtn: $("diagnosticsBtn"),
  metricConfig: $("metricConfig"),
  metricRange: $("metricRange"),
  metricBandwidth: $("metricBandwidth"),
  metricPoints: $("metricPoints"),
  canvas: $("spectrumCanvas"),
  peakRows: $("peakRows"),
  aiResult: $("aiResult"),
  eventLog: $("eventLog"),
  downloadSlot: $("downloadSlot"),
  diagnosticsResult: $("diagnosticsResult"),
  customer: $("customer"),
  eut: $("eut"),
  model: $("model"),
  engineer: $("engineer"),
  remark: $("remark"),
};

function logEvent(message) {
  const li = document.createElement("li");
  li.textContent = `${new Date().toLocaleTimeString()} ${message}`;
  elements.eventLog.prepend(li);
  while (elements.eventLog.children.length > 80) {
    elements.eventLog.lastElementChild.remove();
  }
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "请求失败");
  }
  return payload.data;
}

const post = (path, data = {}) =>
  api(path, {
    method: "POST",
    body: JSON.stringify(data),
  });

function setBusy(isBusy) {
  state.busy = isBusy;
  document.body.classList.toggle("is-busy", isBusy);
}

function userInfoPayload() {
  return {
    customer: elements.customer.value,
    eut: elements.eut.value,
    model: elements.model.value,
    engineer: elements.engineer.value,
    remark: elements.remark.value,
  };
}

function updateStatus(status) {
  const connected = Boolean(status.connected);
  const measuring = Boolean(status.measurement_in_progress);
  state.wasMeasuring = state.wasMeasuring || measuring;
  elements.statusDot.classList.toggle("connected", connected && !measuring);
  elements.statusDot.classList.toggle("busy", measuring);
  elements.statusTitle.textContent = measuring ? "测量中" : connected ? "已连接" : "就绪";
  const demoSuffix = status.demo_mode ? " · 演示数据" : "";
  elements.statusDetail.textContent = `${status.progress_message || status.last_error || "等待连接仪器"}${demoSuffix}`;

  elements.metricConfig.textContent = status.current_config || "未配置";
  const start = status.start_freq ? (status.start_freq / 1e6).toFixed(3) : "--";
  const stop = status.stop_freq ? (status.stop_freq / 1e6).toFixed(3) : "--";
  elements.metricRange.textContent = `${start} - ${stop} MHz`;
  elements.metricBandwidth.textContent = status.rbw ? `${status.rbw} / ${status.vbw} Hz` : "--";
  elements.metricPoints.textContent = status.n_points || "--";

  elements.connectBtn.disabled = connected || measuring;
  elements.disconnectBtn.disabled = !connected || measuring;
  elements.configureBtn.disabled = !connected || measuring;
  elements.singleBtn.disabled = !connected || !status.current_config || measuring;
  elements.slowBtn.disabled = !connected || !status.current_config || measuring;
  elements.fastBtn.disabled = !connected || !status.current_config || measuring;
  elements.stopBtn.disabled = !measuring;
  elements.saveBtn.disabled = measuring || (!status.has_single_data && !status.has_emi_data);
  elements.analyzeBtn.disabled = measuring || (!status.has_single_data && !status.has_emi_data);
  elements.pdfBtn.disabled = measuring || !status.has_emi_data;
  elements.demoBtn.disabled = measuring;

  if (status.user_info) {
    for (const [key, value] of Object.entries(status.user_info)) {
      if (elements[key] && document.activeElement !== elements[key]) {
        elements[key].value = value;
      }
    }
  }
}

function populatePresets(presets) {
  state.presets = presets;
  elements.presetSelect.innerHTML = "";
  for (const [key, config] of Object.entries(presets)) {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = `${config.name} · ${config.description}`;
    elements.presetSelect.append(option);
  }
}

function renderSwitch(status) {
  if (!status || !status.connected) {
    elements.switchReadout.textContent = "切换器未连接。";
    return;
  }
  elements.switchReadout.textContent = [
    `型号: ${status.model || "--"}`,
    `序列号: ${status.serial || "--"}`,
    `固件: ${status.firmware || "--"}`,
    `温度: ${status.temperature || "--"}`,
    `USB 状态: ${status.usb_status || "--"}`,
  ].join("\n");

  const positions = status.positions || {};
  for (const switchName of ["A", "B", "C", "D"]) {
    const module = elements.switchGrid.querySelector(`[data-switch="${switchName}"]`);
    if (module) {
      module.querySelector("strong").textContent = `${switchName} / 位置 ${positions[switchName] || "--"}`;
    }
  }
}

function createSwitchControls() {
  elements.switchGrid.innerHTML = "";
  for (const switchName of ["A", "B", "C", "D"]) {
    const module = document.createElement("div");
    module.className = "switch-module";
    module.dataset.switch = switchName;
    module.innerHTML = `
      <strong>${switchName} / 位置 --</strong>
      <div>
        <button data-pos="1">位置 1</button>
        <button data-pos="2">位置 2</button>
      </div>
    `;
    module.querySelectorAll("button").forEach((button) => {
      button.addEventListener("click", async () => {
        await runAction(`切换器 ${switchName} -> 位置 ${button.dataset.pos}`, async () => {
          const data = await post("/api/switch/set", {
            switch: switchName,
            position: Number(button.dataset.pos),
          });
          renderSwitch(data);
        });
      });
    });
    elements.switchGrid.append(module);
  }
}

function fccLimit(freqHz) {
  const mhz = freqHz / 1e6;
  if (mhz < 30) return 34;
  if (mhz < 88) return 40;
  if (mhz < 216) return 43.5;
  if (mhz < 960) return 46;
  return 54;
}

function ceLimit(freqHz) {
  const mhz = freqHz / 1e6;
  if (mhz < 30) return 34;
  if (mhz < 230) return 40;
  if (mhz < 1000) return 47;
  return 54;
}

function renderChart(series, peaks = []) {
  const canvas = elements.canvas;
  const ctx = canvas.getContext("2d");
  const rect = canvas.getBoundingClientRect();
  const scale = window.devicePixelRatio || 1;
  canvas.width = Math.max(900, Math.floor(rect.width * scale));
  canvas.height = Math.max(420, Math.floor(rect.height * scale));
  ctx.setTransform(scale, 0, 0, scale, 0, 0);

  const width = canvas.width / scale;
  const height = canvas.height / scale;
  ctx.clearRect(0, 0, width, height);

  if (!series || !series.frequency_mhz?.length) {
    ctx.fillStyle = "#60717a";
    ctx.font = "18px Bahnschrift, sans-serif";
    ctx.fillText("等待测量数据...", 36, 52);
    return;
  }

  const xVals = series.frequency_mhz;
  const yVals = series.amplitude_dbuv;
  const minX = Math.min(...xVals);
  const maxX = Math.max(...xVals);
  const minY = Math.min(10, Math.min(...yVals) - 8);
  const maxY = Math.max(80, Math.max(...yVals) + 8);
  const pad = { left: 68, right: 22, top: 28, bottom: 50 };
  const plotW = width - pad.left - pad.right;
  const plotH = height - pad.top - pad.bottom;
  const logMin = Math.log10(Math.max(minX, 0.001));
  const logMax = Math.log10(Math.max(maxX, minX + 1));
  const x = (mhz) => pad.left + ((Math.log10(Math.max(mhz, 0.001)) - logMin) / (logMax - logMin)) * plotW;
  const y = (dbuv) => pad.top + (1 - (dbuv - minY) / (maxY - minY)) * plotH;

  ctx.strokeStyle = "rgba(31, 57, 63, .14)";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 8; i++) {
    const gy = pad.top + (plotH / 8) * i;
    ctx.beginPath();
    ctx.moveTo(pad.left, gy);
    ctx.lineTo(width - pad.right, gy);
    ctx.stroke();
  }

  const drawLine = (values, color, dash = []) => {
    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    ctx.setLineDash(dash);
    ctx.beginPath();
    values.forEach((value, i) => {
      const px = x(xVals[i]);
      const py = y(value);
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    });
    ctx.stroke();
    ctx.restore();
  };

  drawLine(yVals, "#0a6a72");
  drawLine(xVals.map((mhz) => fccLimit(mhz * 1e6)), "#b7442e", [8, 6]);
  drawLine(xVals.map((mhz) => ceLimit(mhz * 1e6)), "#23744a", [4, 5]);

  for (const peak of peaks.slice(0, 28)) {
    const px = x(peak.frequency_mhz);
    const py = y(peak.amplitude_dbuv);
    const fail = peak.exceed_fcc || peak.exceed_ce;
    ctx.fillStyle = fail ? "#b7442e" : "#d9822b";
    ctx.beginPath();
    ctx.arc(px, py, fail ? 5 : 3.5, 0, Math.PI * 2);
    ctx.fill();
  }

  ctx.fillStyle = "#60717a";
  ctx.font = "12px Cascadia Code, monospace";
  ctx.fillText(`${minX.toFixed(3)} MHz`, pad.left, height - 18);
  ctx.fillText(`${maxX.toFixed(3)} MHz`, width - pad.right - 112, height - 18);
  ctx.save();
  ctx.translate(20, height / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText("幅度 dBμV", 0, 0);
  ctx.restore();
}

function renderPeaks(peaks = []) {
  elements.peakRows.innerHTML = "";
  if (!peaks.length) {
    elements.peakRows.innerHTML = `<tr><td colspan="6">暂无数据</td></tr>`;
    return;
  }
  for (const [index, peak] of peaks.entries()) {
    const fail = peak.exceed_fcc || peak.exceed_ce;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${peak.frequency_mhz.toFixed(3)}</td>
      <td>${peak.amplitude_dbuv.toFixed(2)}</td>
      <td>${peak.fcc_margin.toFixed(2)}</td>
      <td>${peak.ce_margin.toFixed(2)}</td>
      <td class="${fail ? "status-fail" : "status-pass"}">${fail ? "失败" : "通过"}</td>
    `;
    elements.peakRows.append(tr);
  }
}

function renderDiagnostics(data) {
  const formatItems = (title, items) => [
    `[${title}]`,
    ...items.map((item) => `${item.ok ? "正常" : "缺失"} ${item.name} - ${item.detail || ""}`),
  ];
  const lines = [
    ...formatItems("Python 依赖", data.packages || []),
    "",
    ...formatItems("关键文件", data.files || []),
    "",
    ...formatItems("环境变量", data.environment || []),
  ];
  if (data.switch_import_error) {
    lines.push("", `[切换器模块导入提示]`, data.switch_import_error);
  }
  elements.diagnosticsResult.textContent = lines.join("\n");
}

async function refreshStatus() {
  const status = await api("/api/status");
  updateStatus(status);
  return status;
}

async function refreshResult() {
  const result = await api("/api/result");
  state.result = result;
  updateStatus(result.status);
  renderChart(result.series, result.peaks);
  renderPeaks(result.peaks);
  if (result.ai_result) {
    elements.aiResult.textContent = result.ai_result;
  }
  return result;
}

async function runAction(label, action) {
  try {
    setBusy(true);
    logEvent(`${label} ...`);
    const result = await action();
    await refreshResult();
    logEvent(`${label} 完成`);
    return result;
  } catch (error) {
    logEvent(`${label} 失败: ${error.message}`);
    alert(error.message);
  } finally {
    setBusy(false);
  }
}

function bindEvents() {
  elements.connectBtn.addEventListener("click", () =>
    runAction("连接仪器", () => post("/api/device/connect", { ip_address: elements.deviceIp.value })),
  );
  elements.disconnectBtn.addEventListener("click", () =>
    runAction("断开仪器", () => post("/api/device/disconnect")),
  );
  elements.configureBtn.addEventListener("click", () =>
    runAction("应用测试配置", () => post("/api/configure", { preset_key: elements.presetSelect.value })),
  );
  elements.demoBtn.addEventListener("click", () =>
    runAction("加载演示数据", () =>
      post("/api/demo/load", {
        preset_key: elements.presetSelect.value || "EMC_30MHz_1GHz",
        duration_seconds: 15,
      }),
    ),
  );
  elements.connectSwitchBtn.addEventListener("click", () =>
    runAction("连接切换器", async () => {
      renderSwitch(await post("/api/switch/connect"));
    }),
  );
  elements.refreshSwitchBtn.addEventListener("click", () =>
    runAction("刷新切换器", async () => {
      renderSwitch(await api("/api/switch/status"));
    }),
  );
  elements.disconnectSwitchBtn.addEventListener("click", () =>
    runAction("断开切换器", async () => {
      renderSwitch(await post("/api/switch/disconnect"));
    }),
  );
  elements.singleBtn.addEventListener("click", () =>
    runAction("单次扫描", () => post("/api/measure/single")),
  );
  elements.slowBtn.addEventListener("click", () =>
    runAction("15 秒 EMI 采样", () => post("/api/measure/timed", { duration_seconds: 15 })),
  );
  elements.fastBtn.addEventListener("click", () =>
    runAction("5 分钟 EMI 采样", () => post("/api/measure/timed", { duration_seconds: 300 })),
  );
  elements.stopBtn.addEventListener("click", () =>
    runAction("停止测量", () => post("/api/measure/stop")),
  );
  elements.saveBtn.addEventListener("click", () =>
    runAction("保存数据", async () => {
      const result = await post("/api/data/save");
      alert(`已保存:\n${result.files.join("\n")}`);
    }),
  );
  elements.analyzeBtn.addEventListener("click", () =>
    runAction("AI 异常分析", async () => {
      const data = await post("/api/ai/analyze");
      elements.aiResult.textContent = data.result;
    }),
  );
  elements.pdfBtn.addEventListener("click", () =>
    runAction("导出 PDF 报告", async () => {
      const data = await post("/api/report/export", { user_info: userInfoPayload(), auto_analyze: true });
      elements.downloadSlot.innerHTML = `<a href="${data.download_url}">下载报告：${data.file}</a>`;
    }),
  );
  elements.diagnosticsBtn.addEventListener("click", () =>
    runAction("环境诊断", async () => {
      renderDiagnostics(await api("/api/diagnostics"));
    }),
  );
  for (const input of [elements.customer, elements.eut, elements.model, elements.engineer, elements.remark]) {
    input.addEventListener("change", () => post("/api/user-info", userInfoPayload()).catch(console.warn));
  }
  window.addEventListener("resize", () => renderChart(state.result?.series, state.result?.peaks || []));
}

async function init() {
  createSwitchControls();
  bindEvents();
  populatePresets(await api("/api/presets"));
  await refreshResult();
  logEvent("Web 控制台已初始化");
  setInterval(async () => {
    try {
      const status = await refreshStatus();
      if (status.measurement_in_progress || state.wasMeasuring) {
        await refreshResult();
      }
      state.wasMeasuring = status.measurement_in_progress;
    } catch (error) {
      logEvent(`轮询失败: ${error.message}`);
    }
  }, 1800);
}

init().catch((error) => {
  logEvent(`初始化失败: ${error.message}`);
  alert(error.message);
});
