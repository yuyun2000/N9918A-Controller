const state = {
  presets: {},
  naPresets: {},
  result: null,
  naResult: null,
  busy: false,
  wasMeasuring: false,
  mode: "SA",
  valleyPage: 1,
  valleysPerPage: 12,
  lastStatus: null,
};

const $ = (id) => document.getElementById(id);

const elements = {
  statusDot: $("statusDot"),
  statusTitle: $("statusTitle"),
  statusDetail: $("statusDetail"),
  modeSA: $("modeSA"),
  modeNA: $("modeNA"),
  modeHint: $("modeHint"),
  saWorkspace: $("saWorkspace"),
  naWorkspace: $("naWorkspace"),
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
  naPresetSelect: $("naPresetSelect"),
  naConfigureBtn: $("naConfigureBtn"),
  naCalibrateBtn: $("naCalibrateBtn"),
  naMeasureBtn: $("naMeasureBtn"),
  naStopBtn: $("naStopBtn"),
  naCalibrationLog: $("naCalibrationLog"),
  naConfigText: $("naConfigText"),
  naCalibrationText: $("naCalibrationText"),
  naSwitchText: $("naSwitchText"),
  naStatusText: $("naStatusText"),
  naDownloadSlot: $("naDownloadSlot"),
  naSaveBtn: $("naSaveBtn"),
  naExportBtn: $("naExportBtn"),
  naS11Canvas: $("naS11Canvas"),
  smithCanvas: $("smithCanvas"),
  naPrimaryText: $("naPrimaryText"),
  naTargetText: $("naTargetText"),
  bandwidthGrid: $("bandwidthGrid"),
  pointRows: $("pointRows"),
  targetRows: $("targetRows"),
  valleyRows: $("valleyRows"),
  valleyPrevBtn: $("valleyPrevBtn"),
  valleyNextBtn: $("valleyNextBtn"),
  valleyPageInfo: $("valleyPageInfo"),
};

function logEvent(message) {
  const li = document.createElement("li");
  li.textContent = `${new Date().toLocaleTimeString()} ${message}`;
  elements.eventLog.prepend(li);
  while (elements.eventLog.children.length > 100) {
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

function setModeUi(mode) {
  state.mode = mode;
  elements.modeSA.classList.toggle("active", mode === "SA");
  elements.modeNA.classList.toggle("active", mode === "NA");
  elements.saWorkspace.hidden = mode !== "SA";
  elements.naWorkspace.hidden = mode !== "NA";
  elements.saWorkspace.classList.toggle("active", mode === "SA");
  elements.naWorkspace.classList.toggle("active", mode === "NA");
  elements.modeHint.textContent = mode === "NA"
    ? "NA 模式会调用 FieldFox 网络分析功能；校准会自动执行 OPEN(B2C1) → LOAD(B1C1) → ANTENNA(B2C2) 的 switchbox 顺序。"
    : "SA 模式保留原有频谱扫描、EMI 采样、AI 分析和 PDF 报告流程。";
  if (mode === "SA") {
    renderChart(state.result?.series, state.result?.peaks || []);
  } else {
    renderNaS11(state.naResult?.series, state.naResult?.primary_valley, state.naResult?.bandwidths, state.naResult?.points_of_interest);
    renderSmith(state.naResult?.smith, state.naResult?.is_full_sweep);
  }
}

async function switchMode(mode) {
  await runAction(`切换到 ${mode} 模式`, async () => {
    const data = await post("/api/mode", { mode });
    setModeUi(data.current_mode || mode);
    await refreshStatus();
    if ((data.current_mode || mode) === "NA") {
      await refreshNaResult();
    } else {
      await refreshResult();
    }
  });
}

function updateStatus(status) {
  state.lastStatus = status;
  const connected = Boolean(status.connected);
  const measuring = Boolean(status.measurement_in_progress);
  state.wasMeasuring = state.wasMeasuring || measuring;
  setModeUi(status.current_mode || state.mode);

  elements.statusDot.classList.toggle("connected", connected && !measuring);
  elements.statusDot.classList.toggle("busy", measuring || status.switching_mode);
  elements.statusTitle.textContent = measuring ? "测量/校准中" : connected ? "已连接" : "就绪";
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
  elements.configureBtn.disabled = !connected || measuring || state.mode !== "SA";
  elements.singleBtn.disabled = !connected || !status.current_config || measuring || state.mode !== "SA";
  elements.slowBtn.disabled = !connected || !status.current_config || measuring || state.mode !== "SA";
  elements.fastBtn.disabled = !connected || !status.current_config || measuring || state.mode !== "SA";
  elements.stopBtn.disabled = !measuring;
  elements.saveBtn.disabled = measuring || (!status.has_single_data && !status.has_emi_data);
  elements.analyzeBtn.disabled = measuring || (!status.has_single_data && !status.has_emi_data);
  elements.pdfBtn.disabled = measuring || !status.has_emi_data;
  elements.demoBtn.disabled = measuring;
  elements.modeSA.disabled = measuring || status.switching_mode;
  elements.modeNA.disabled = measuring || status.switching_mode;
  elements.connectSwitchBtn.disabled = measuring;
  elements.refreshSwitchBtn.disabled = measuring;
  elements.disconnectSwitchBtn.disabled = measuring;
  elements.switchGrid.querySelectorAll("button").forEach((button) => {
    button.disabled = measuring;
  });

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

function populateNaPresets(presets) {
  state.naPresets = presets;
  elements.naPresetSelect.innerHTML = "";
  for (const [key, config] of Object.entries(presets)) {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = `${config.name} · ${formatHz(config.start_freq)}-${formatHz(config.stop_freq)} · ${config.description}`;
    elements.naPresetSelect.append(option);
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
      const position = positions[switchName] || "--";
      module.dataset.currentPosition = position;
      module.querySelector("strong").textContent = `${switchName} / 位置 ${position}`;
      module.querySelectorAll("button").forEach((button) => {
        button.classList.toggle("active", String(position) === button.dataset.pos);
      });
    }
  }
}

function createSwitchControls() {
  elements.switchGrid.innerHTML = "";
  const switchOrder = ["A", "B", "D", "C"];
  for (const switchName of switchOrder) {
    const module = document.createElement("div");
    module.className = "switch-module";
    module.dataset.switch = switchName;
    module.innerHTML = `
      <strong>${switchName} / 位置 --</strong>
      <div>
        <button class="switch-pos pos-one" data-pos="1"><span></span>位置 1</button>
        <button class="switch-pos pos-two" data-pos="2"><span></span>位置 2</button>
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

function prepareCanvas(canvas, minWidth = 900, minHeight = 420) {
  const ctx = canvas.getContext("2d");
  const rect = canvas.getBoundingClientRect();
  const scale = window.devicePixelRatio || 1;
  canvas.width = Math.max(minWidth, Math.floor(rect.width * scale));
  canvas.height = Math.max(minHeight, Math.floor(rect.height * scale));
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  return { ctx, width: canvas.width / scale, height: canvas.height / scale };
}

function renderChart(series, peaks = []) {
  const { ctx, width, height } = prepareCanvas(elements.canvas);
  ctx.clearRect(0, 0, width, height);

  if (!series || !series.frequency_mhz?.length) {
    drawEmpty(ctx, "等待 SA 测量数据...", width, height);
    return;
  }

  const xVals = series.frequency_mhz;
  const yVals = series.amplitude_dbuv;
  const minX = Math.min(...xVals);
  const maxX = Math.max(...xVals);
  const minY = Math.min(10, Math.min(...yVals) - 8);
  const maxY = Math.max(80, Math.max(...yVals) + 8);
  const pad = { left: 68, right: 22, top: 28, bottom: 50 };
  const { x, y } = scaledPlotters(minX, maxX, minY, maxY, width, height, pad, true);

  drawGrid(ctx, width, height, pad);
  drawLine(ctx, xVals, yVals, x, y, "#0a6a72", 2.2);
  drawLine(ctx, xVals, xVals.map((mhz) => fccLimit(mhz * 1e6)), x, y, "#b7442e", 1.6, [8, 6]);
  drawLine(ctx, xVals, xVals.map((mhz) => ceLimit(mhz * 1e6)), x, y, "#23744a", 1.6, [4, 5]);

  for (const peak of peaks.slice(0, 28)) {
    const px = x(peak.frequency_mhz);
    const py = y(peak.amplitude_dbuv);
    const fail = peak.exceed_fcc || peak.exceed_ce;
    ctx.fillStyle = fail ? "#b7442e" : "#d9822b";
    ctx.beginPath();
    ctx.arc(px, py, fail ? 5 : 3.5, 0, Math.PI * 2);
    ctx.fill();
  }

  drawAxisLabels(ctx, width, height, pad, `${minX.toFixed(3)} MHz`, `${maxX.toFixed(3)} MHz`, "幅度 dBμV");
}

function renderNaS11(series, primaryValley, bandwidths = {}, points = []) {
  const { ctx, width, height } = prepareCanvas(elements.naS11Canvas);
  ctx.clearRect(0, 0, width, height);

  if (!series || !series.frequency_mhz?.length) {
    drawEmpty(ctx, "等待 NA S11 测量数据...", width, height);
    return;
  }

  const xVals = series.frequency_mhz;
  const yVals = series.s11_db;
  const minX = Math.min(...xVals);
  const maxX = Math.max(...xVals);
  const minY = Math.min(-35, Math.min(...yVals) - 4);
  const maxY = Math.max(2, Math.max(...yVals) + 3);
  const pad = { left: 70, right: 26, top: 30, bottom: 52 };
  const useLog = maxX / Math.max(minX, 0.001) > 4;
  const { x, y } = scaledPlotters(minX, maxX, minY, maxY, width, height, pad, useLog);

  for (const [key, bw] of Object.entries(bandwidths || {})) {
    if (!bw?.left_hz || !bw?.right_hz) continue;
    const color = key.includes("10") ? "rgba(183, 68, 46, 0.16)" : "rgba(217, 130, 43, 0.14)";
    ctx.fillStyle = color;
    const left = x(bw.left_hz / 1e6);
    const right = x(bw.right_hz / 1e6);
    ctx.fillRect(Math.min(left, right), pad.top, Math.abs(right - left), height - pad.top - pad.bottom);
  }

  drawGrid(ctx, width, height, pad);
  drawLine(ctx, xVals, yVals, x, y, "#0a6a72", 2.4);

  const thresholds = [
    { value: -3, color: "#d9822b", dash: [8, 6], label: "-3dB" },
    { value: -10, color: "#b7442e", dash: [4, 5], label: "-10dB" },
  ];
  for (const item of thresholds) {
    drawHorizontal(ctx, y(item.value), pad.left, width - pad.right, item.color, item.dash);
    ctx.fillStyle = item.color;
    ctx.font = "12px Cascadia Code, monospace";
    ctx.fillText(item.label, width - pad.right - 50, y(item.value) - 6);
  }

  const pointColors = {
    center: "#b7442e",
    target: "#23744a",
    absolute_3db_left: "#d9822b",
    absolute_3db_right: "#d9822b",
    absolute_10db_left: "#b7442e",
    absolute_10db_right: "#b7442e",
  };
  const plottedPoints = (points || []).filter((point) => pointColors[point.type]);
  const centerPoint = plottedPoints.find((point) => point.type === "center");
  const targetPoint = plottedPoints.find((point) => point.type === "target");
  if (targetPoint) {
    const px = x(targetPoint.frequency_mhz);
    drawVertical(ctx, px, pad.top, height - pad.bottom, pointColors.target, [6, 5]);
    drawTag(ctx, `理想 ${targetPoint.frequency_mhz.toFixed(3)}MHz / RL ${formatDb(targetPoint.return_loss_db)}`, clamp(px + 8, pad.left, width - 260), pad.top + 10, pointColors.target);
  }
  if (centerPoint) {
    const px = x(centerPoint.frequency_mhz);
    drawVertical(ctx, px, pad.top, height - pad.bottom, pointColors.center, [2, 4]);
    const yOffset = targetPoint && Math.abs(x(targetPoint.frequency_mhz) - px) < 180 ? 44 : 10;
    drawTag(ctx, `实际 ${centerPoint.frequency_mhz.toFixed(3)}MHz / RL ${formatDb(centerPoint.return_loss_db)}`, clamp(px + 8, pad.left, width - 260), pad.top + yOffset, pointColors.center);
  }
  for (const point of plottedPoints) {
    const px = x(point.frequency_mhz);
    const py = y(point.s11_db);
    ctx.fillStyle = pointColors[point.type];
    ctx.beginPath();
    ctx.arc(px, py, point.type === "center" ? 6 : 4.5, 0, Math.PI * 2);
    ctx.fill();
  }

  if (!plottedPoints.length && primaryValley) {
    const px = x(primaryValley.frequency_mhz);
    const py = y(primaryValley.s11_db);
    ctx.fillStyle = "#d9822b";
    ctx.beginPath();
    ctx.arc(px, py, 6, 0, Math.PI * 2);
    ctx.fill();
  }

  drawAxisLabels(ctx, width, height, pad, `${minX.toFixed(3)} MHz`, `${maxX.toFixed(3)} MHz`, "S11 dB");
}

function renderSmith(smith, isFullSweep) {
  const { ctx, width, height } = prepareCanvas(elements.smithCanvas, 520, 420);
  ctx.clearRect(0, 0, width, height);
  const size = Math.min(width, height) * 0.78;
  const cx = width / 2;
  const cy = height / 2;
  const r = size / 2;

  ctx.fillStyle = "rgba(255, 255, 255, 0.62)";
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = "rgba(31,57,63,.18)";
  ctx.lineWidth = 1;
  for (const fraction of [0.25, 0.5, 0.75, 1]) {
    ctx.beginPath();
    ctx.arc(cx, cy, r * fraction, 0, Math.PI * 2);
    ctx.stroke();
  }
  ctx.beginPath();
  ctx.moveTo(cx - r, cy);
  ctx.lineTo(cx + r, cy);
  ctx.moveTo(cx, cy - r);
  ctx.lineTo(cx, cy + r);
  ctx.stroke();

  if (isFullSweep) {
    drawCentered(ctx, "全扫宽结果不显示 Smith Chart", width, height);
    return;
  }
  if (!smith || !smith.real?.length) {
    drawCentered(ctx, "等待复数 Gamma 数据...", width, height);
    return;
  }

  const toX = (value) => cx + value * r;
  const toY = (value) => cy - value * r;
  ctx.strokeStyle = "#0a6a72";
  ctx.lineWidth = 2;
  ctx.beginPath();
  smith.real.forEach((real, index) => {
    const px = toX(Math.max(-1.1, Math.min(1.1, real)));
    const py = toY(Math.max(-1.1, Math.min(1.1, smith.imag[index])));
    if (index === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  });
  ctx.stroke();

  for (const marker of smith.markers || []) {
    if (marker.real == null || marker.imag == null) continue;
    const isCenter = marker.type === "center";
    ctx.fillStyle = isCenter ? "#b7442e" : "#d9822b";
    ctx.beginPath();
    ctx.arc(toX(marker.real), toY(marker.imag), isCenter ? 5.5 : 3.8, 0, Math.PI * 2);
    ctx.fill();
  }
}

function renderBandwidths(data) {
  const primary = data?.primary_valley;
  if (primary) {
    elements.naPrimaryText.textContent = `中心频率 ${primary.frequency_mhz.toFixed(6)} MHz，S11 ${formatDb(primary.s11_db)}，回波损耗 ${formatDb(primary.return_loss_db)}，VSWR ${formatVswr(primary.vswr)}`;
  } else {
    elements.naPrimaryText.textContent = "暂无中心谷值。";
  }

  const labels = {
    absolute_3db: "绝对 -3dB 带宽",
    absolute_10db: "绝对 -10dB 带宽",
    relative_3db: "相对 +3dB 带宽",
    relative_10db: "相对 +10dB 带宽",
  };
  elements.bandwidthGrid.innerHTML = "";
  for (const [key, label] of Object.entries(labels)) {
    const bw = data?.bandwidths?.[key] || {};
    const div = document.createElement("div");
    div.className = "bandwidth-card";
    div.innerHTML = `
      <span>${label}</span>
      <strong>${bw.width_hz == null ? "--" : formatHz(bw.width_hz)}</strong>
      <small>左 ${formatHz(bw.left_hz)} · S11 ${formatDb(bw.left_s11_db)} · VSWR ${formatVswr(bw.left_vswr)}</small>
      <small>右 ${formatHz(bw.right_hz)} · S11 ${formatDb(bw.right_s11_db)} · VSWR ${formatVswr(bw.right_vswr)}</small>
      <small>${bw.complete ? "完整" : "不完整或未跨阈值"}</small>
    `;
    elements.bandwidthGrid.append(div);
  }
}

function renderTargetSummary(summary) {
  if (!summary) {
    elements.naTargetText.className = "target-compare";
    elements.naTargetText.textContent = "暂无理想频点对比。";
    return;
  }
  elements.naTargetText.className = `target-compare ${summary.status || "unknown"}`;
  const errorText = summary.frequency_error_mhz == null
    ? "--"
    : `${summary.frequency_error_mhz >= 0 ? "+" : ""}${Number(summary.frequency_error_mhz).toFixed(3)} MHz`;
  const rlDelta = summary.return_loss_delta_db == null
    ? "--"
    : `${summary.return_loss_delta_db >= 0 ? "+" : ""}${Number(summary.return_loss_delta_db).toFixed(2)} dB`;
  elements.naTargetText.textContent =
    `理想 ${Number(summary.target_frequency_mhz).toFixed(3)} MHz：S11 ${formatDb(summary.target_s11_db)} / RL ${formatDb(summary.target_return_loss_db)} / VSWR ${formatVswr(summary.target_vswr)}；` +
    `实际谷值偏移 ${errorText}，实际谷值相对理想频点回波损耗差 ${rlDelta}，${summary.status_label || "暂无判定"}`;
}

function renderTargetWindow(points = []) {
  elements.targetRows.innerHTML = "";
  if (!points.length) {
    elements.targetRows.innerHTML = `<tr><td colspan="5">暂无理想频点附近数据</td></tr>`;
    return;
  }
  for (const point of points) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${point.offset_percent > 0 ? "+" : ""}${Number(point.offset_percent).toFixed(0)}%</td>
      <td>${Number(point.frequency_mhz).toFixed(6)}</td>
      <td>${formatDb(point.s11_db)}</td>
      <td>${formatDb(point.return_loss_db)}</td>
      <td>${formatVswr(point.vswr)}</td>
    `;
    elements.targetRows.append(tr);
  }
}

function renderPoints(points = []) {
  elements.pointRows.innerHTML = "";
  if (!points.length) {
    elements.pointRows.innerHTML = `<tr><td colspan="5">暂无中心或端点数据</td></tr>`;
    return;
  }
  const preferredOrder = {
    center: 0,
    target: 1,
    absolute_3db_left: 2,
    absolute_3db_right: 3,
    absolute_10db_left: 4,
    absolute_10db_right: 5,
    relative_3db_left: 6,
    relative_3db_right: 7,
    relative_10db_left: 8,
    relative_10db_right: 9,
  };
  const rows = [...points].sort((a, b) => (preferredOrder[a.type] ?? 99) - (preferredOrder[b.type] ?? 99));
  for (const point of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${point.label || point.type}</td>
      <td>${Number(point.frequency_mhz).toFixed(6)}</td>
      <td>${formatDb(point.s11_db)}</td>
      <td>${formatDb(point.return_loss_db)}</td>
      <td>${formatVswr(point.vswr)}</td>
    `;
    elements.pointRows.append(tr);
  }
}

function renderValleys(valleys = []) {
  elements.valleyRows.innerHTML = "";
  if (!valleys.length) {
    elements.valleyRows.innerHTML = `<tr><td colspan="8">暂无 NA 测量数据</td></tr>`;
    elements.valleyPageInfo.textContent = "第 0 / 0 页";
    elements.valleyPrevBtn.disabled = true;
    elements.valleyNextBtn.disabled = true;
    return;
  }
  const pageCount = Math.max(1, Math.ceil(valleys.length / state.valleysPerPage));
  state.valleyPage = Math.min(Math.max(1, state.valleyPage), pageCount);
  const start = (state.valleyPage - 1) * state.valleysPerPage;
  const pageRows = valleys.slice(start, start + state.valleysPerPage);
  for (const [offset, valley] of pageRows.entries()) {
    const abs10 = valley.bandwidths?.absolute_10db;
    const rel3 = valley.bandwidths?.relative_3db;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${start + offset + 1}</td>
      <td>${valley.frequency_mhz.toFixed(6)}</td>
      <td>${formatDb(valley.s11_db)}</td>
      <td>${formatDb(valley.return_loss_db)}</td>
      <td>${formatVswr(valley.vswr)}</td>
      <td>${(valley.prominence_db || 0).toFixed(3)}</td>
      <td>${abs10?.width_hz == null ? "--" : formatHz(abs10.width_hz)}</td>
      <td>${rel3?.width_hz == null ? "--" : formatHz(rel3.width_hz)}</td>
    `;
    elements.valleyRows.append(tr);
  }
  elements.valleyPageInfo.textContent = `第 ${state.valleyPage} / ${pageCount} 页，共 ${valleys.length} 个谷值`;
  elements.valleyPrevBtn.disabled = state.valleyPage <= 1;
  elements.valleyNextBtn.disabled = state.valleyPage >= pageCount;
}

function renderNaStatus(data) {
  state.naResult = data;
  const status = data.status || {};
  const config = status.config || data.config || {};
  const calibration = status.calibration || {};
  elements.naConfigText.textContent = config.name ? `${config.name} · ${formatHz(config.start_freq)}-${formatHz(config.stop_freq)}` : "未配置";
  elements.naCalibrationText.textContent = calibration.in_progress ? "校准中" : calibration.complete ? "已校准" : "未校准";
  elements.naSwitchText.textContent = status.switch_position || "--";
  elements.naStatusText.textContent = status.progress_message || status.error || "等待操作";

  const events = calibration.events || [];
  elements.naCalibrationLog.textContent = events.length
    ? events.map((event) => `${event.ok ? "✓" : "✗"} ${event.label || event.step} · ${event.switch_position || "--"} · ${event.scpi || "switch"} ${event.message ? `· ${event.message}` : ""}`).join("\n")
    : "等待 NA 配置。";

  const connected = Boolean(status.connected || state.lastStatus?.connected);
  const busy = Boolean(status.measurement_in_progress || state.lastStatus?.measurement_in_progress);
  elements.naConfigureBtn.disabled = busy;
  elements.naCalibrateBtn.disabled = busy || !connected || !status.configured;
  elements.naMeasureBtn.disabled = busy || !connected || !status.configured || !calibration.complete;
  elements.naStopBtn.disabled = !busy;
  elements.naSaveBtn.disabled = busy || !data.series;
  elements.naExportBtn.disabled = busy || !data.series;

  renderNaS11(data.series, data.primary_valley, data.bandwidths, data.points_of_interest);
  renderSmith(data.smith, data.is_full_sweep);
  renderBandwidths(data);
  renderTargetSummary(data.target_summary);
  renderPoints(data.points_of_interest || []);
  renderTargetWindow(data.target_window || []);
  renderValleys(data.valleys || []);
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
    lines.push("", "[切换器模块导入提示]", data.switch_import_error);
  }
  elements.diagnosticsResult.textContent = lines.join("\n");
}

async function refreshMode() {
  const mode = await api("/api/mode");
  setModeUi(mode.current_mode || state.mode);
  return mode;
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

async function refreshNaResult() {
  const result = await api("/api/na/result");
  renderNaStatus(result);
  return result;
}

async function runAction(label, action) {
  try {
    setBusy(true);
    logEvent(`${label} ...`);
    const result = await action();
    await refreshMode();
    await refreshStatus();
    if (state.mode === "NA") {
      await refreshNaResult();
    } else {
      await refreshResult();
    }
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
  elements.modeSA.addEventListener("click", () => switchMode("SA"));
  elements.modeNA.addEventListener("click", () => switchMode("NA"));
  elements.connectBtn.addEventListener("click", () =>
    runAction("连接仪器", async () => {
      const requestedMode = state.mode;
      const data = await post("/api/device/connect", { ip_address: elements.deviceIp.value });
      if (requestedMode === "NA") {
        await post("/api/mode", { mode: "NA" });
      }
      return data;
    }),
  );
  elements.disconnectBtn.addEventListener("click", () => runAction("断开仪器", () => post("/api/device/disconnect")));
  elements.configureBtn.addEventListener("click", () =>
    runAction("应用 SA 测试配置", () => post("/api/configure", { preset_key: elements.presetSelect.value })),
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
  elements.singleBtn.addEventListener("click", () => runAction("单次扫描", () => post("/api/measure/single")));
  elements.slowBtn.addEventListener("click", () => runAction("15 秒 EMI 采样", () => post("/api/measure/timed", { duration_seconds: 15 })));
  elements.fastBtn.addEventListener("click", () => runAction("5 分钟 EMI 采样", () => post("/api/measure/timed", { duration_seconds: 300 })));
  elements.stopBtn.addEventListener("click", () => runAction("停止测量", () => post("/api/measure/stop")));
  elements.saveBtn.addEventListener("click", () =>
    runAction("保存数据", async () => {
      const result = await post("/api/data/save");
      alert(`已保存\n${result.files.join("\n")}`);
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

  elements.naConfigureBtn.addEventListener("click", () =>
    runAction("应用 NA 预设", () => post("/api/na/configure", { preset_key: elements.naPresetSelect.value })),
  );
  elements.naCalibrateBtn.addEventListener("click", () => runAction("NA 自动校准", () => post("/api/na/calibrate")));
  elements.naMeasureBtn.addEventListener("click", () => runAction("NA 天线测量", () => post("/api/na/measure")));
  elements.naStopBtn.addEventListener("click", () => runAction("停止 NA 流程", () => post("/api/na/stop")));
  elements.naSaveBtn.addEventListener("click", () =>
    runAction("保存 NA 数据", async () => {
      const result = await post("/api/na/data/save");
      alert(`已保存\n${result.files.join("\n")}`);
    }),
  );
  elements.naExportBtn.addEventListener("click", () =>
    runAction("导出 NA 报告", async () => {
      const data = await post("/api/na/report/export", { user_info: userInfoPayload() });
      elements.naDownloadSlot.innerHTML = `<a href="${data.download_url}">下载 NA 报告：${data.file}</a>`;
    }),
  );
  elements.valleyPrevBtn.addEventListener("click", () => {
    state.valleyPage -= 1;
    renderValleys(state.naResult?.valleys || []);
  });
  elements.valleyNextBtn.addEventListener("click", () => {
    state.valleyPage += 1;
    renderValleys(state.naResult?.valleys || []);
  });
  for (const input of [elements.customer, elements.eut, elements.model, elements.engineer, elements.remark]) {
    input.addEventListener("change", () => post("/api/user-info", userInfoPayload()).catch(console.warn));
  }
  window.addEventListener("resize", () => {
    renderChart(state.result?.series, state.result?.peaks || []);
    renderNaS11(state.naResult?.series, state.naResult?.primary_valley, state.naResult?.bandwidths, state.naResult?.points_of_interest);
    renderSmith(state.naResult?.smith, state.naResult?.is_full_sweep);
  });
}

function drawEmpty(ctx, message, width, height) {
  ctx.fillStyle = "#60717a";
  ctx.font = "18px Bahnschrift, sans-serif";
  ctx.fillText(message, 36, 52);
}

function drawCentered(ctx, message, width, height) {
  ctx.fillStyle = "#60717a";
  ctx.font = "18px Bahnschrift, sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(message, width / 2, height / 2);
  ctx.textAlign = "left";
}

function drawGrid(ctx, width, height, pad) {
  ctx.strokeStyle = "rgba(31, 57, 63, .14)";
  ctx.lineWidth = 1;
  const plotH = height - pad.top - pad.bottom;
  for (let i = 0; i <= 8; i++) {
    const gy = pad.top + (plotH / 8) * i;
    ctx.beginPath();
    ctx.moveTo(pad.left, gy);
    ctx.lineTo(width - pad.right, gy);
    ctx.stroke();
  }
}

function scaledPlotters(minX, maxX, minY, maxY, width, height, pad, useLog) {
  const plotW = width - pad.left - pad.right;
  const plotH = height - pad.top - pad.bottom;
  const logMin = Math.log10(Math.max(minX, 0.001));
  const logMax = Math.log10(Math.max(maxX, minX + 0.001));
  const x = (value) => {
    if (useLog) {
      return pad.left + ((Math.log10(Math.max(value, 0.001)) - logMin) / Math.max(logMax - logMin, 1e-9)) * plotW;
    }
    return pad.left + ((value - minX) / Math.max(maxX - minX, 1e-9)) * plotW;
  };
  const y = (value) => pad.top + (1 - (value - minY) / Math.max(maxY - minY, 1e-9)) * plotH;
  return { x, y };
}

function drawLine(ctx, xVals, yVals, x, y, color, width = 2, dash = []) {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.setLineDash(dash);
  ctx.beginPath();
  yVals.forEach((value, index) => {
    const px = x(xVals[index]);
    const py = y(value);
    if (index === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  });
  ctx.stroke();
  ctx.restore();
}

function drawHorizontal(ctx, y, left, right, color, dash = []) {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.4;
  ctx.setLineDash(dash);
  ctx.beginPath();
  ctx.moveTo(left, y);
  ctx.lineTo(right, y);
  ctx.stroke();
  ctx.restore();
}

function drawVertical(ctx, x, top, bottom, color, dash = []) {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.25;
  ctx.setLineDash(dash);
  ctx.beginPath();
  ctx.moveTo(x, top);
  ctx.lineTo(x, bottom);
  ctx.stroke();
  ctx.restore();
}

function drawTag(ctx, text, x, y, color) {
  ctx.save();
  ctx.font = "12px Bahnschrift, sans-serif";
  const width = ctx.measureText(text).width + 16;
  const height = 24;
  ctx.fillStyle = "rgba(255, 248, 232, 0.93)";
  ctx.strokeStyle = color;
  ctx.lineWidth = 1;
  roundRect(ctx, x, y, width, height, 9);
  ctx.fill();
  ctx.stroke();
  ctx.fillStyle = color;
  ctx.fillText(text, x + 8, y + 16);
  ctx.restore();
}

function roundRect(ctx, x, y, width, height, radius) {
  const r = Math.min(radius, width / 2, height / 2);
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + width, y, x + width, y + height, r);
  ctx.arcTo(x + width, y + height, x, y + height, r);
  ctx.arcTo(x, y + height, x, y, r);
  ctx.arcTo(x, y, x + width, y, r);
  ctx.closePath();
}

function drawAxisLabels(ctx, width, height, pad, leftLabel, rightLabel, yLabel) {
  ctx.fillStyle = "#60717a";
  ctx.font = "12px Cascadia Code, monospace";
  ctx.fillText(leftLabel, pad.left, height - 18);
  ctx.fillText(rightLabel, width - pad.right - Math.min(150, rightLabel.length * 7), height - 18);
  ctx.save();
  ctx.translate(20, height / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText(yLabel, 0, 0);
  ctx.restore();
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function formatHz(value) {
  if (value == null || Number.isNaN(Number(value))) return "--";
  const number = Number(value);
  const abs = Math.abs(number);
  if (abs >= 1e9) return `${(number / 1e9).toFixed(abs >= 10e9 ? 3 : 4)} GHz`;
  if (abs >= 1e6) return `${(number / 1e6).toFixed(abs >= 100e6 ? 3 : 4)} MHz`;
  if (abs >= 1e3) return `${(number / 1e3).toFixed(3)} kHz`;
  return `${number.toFixed(3)} Hz`;
}

function formatDb(value) {
  if (value == null || Number.isNaN(Number(value))) return "--";
  return `${Number(value).toFixed(3)} dB`;
}

function formatVswr(value) {
  if (value == null || Number.isNaN(Number(value))) return "∞";
  return Number(value).toFixed(3);
}

async function init() {
  createSwitchControls();
  bindEvents();
  populatePresets(await api("/api/presets"));
  populateNaPresets(await api("/api/na/presets"));
  await refreshMode();
  await refreshResult();
  await refreshNaResult();
  logEvent("Web 控制台已初始化");
  setInterval(async () => {
    try {
      const status = await refreshStatus();
      if (state.mode === "NA") {
        await refreshNaResult();
      }
      if (status.measurement_in_progress || state.wasMeasuring) {
        if (state.mode === "NA") await refreshNaResult();
        else await refreshResult();
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
