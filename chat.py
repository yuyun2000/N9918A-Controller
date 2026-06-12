# -*- coding: utf-8 -*-
"""AI analysis client for SA screening results."""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace

LOCAL_CONFIG_PATH = Path(__file__).resolve().with_name("ai_config.local.json")


def _load_local_config():
    if not LOCAL_CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(LOCAL_CONFIG_PATH.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise ValueError(f"AI local config is invalid: {LOCAL_CONFIG_PATH} ({exc})") from exc


LOCAL_CONFIG = _load_local_config()


def _config_value(key, env_names, default=None):
    if key in LOCAL_CONFIG and LOCAL_CONFIG[key] not in (None, ""):
        return LOCAL_CONFIG[key]
    for name in env_names:
        value = os.getenv(name)
        if value:
            return value
    return default


DEFAULT_BASE_URL = _config_value("base_url", ("N9918A_AI_BASE_URL", "OPENAI_BASE_URL"), "http://192.168.20.38:3000/")
DEFAULT_MODEL = _config_value("model", ("N9918A_AI_MODEL", "OPENAI_MODEL"), "gpt-5.5")
DEFAULT_REASONING_EFFORT = _config_value("reasoning_effort", ("N9918A_AI_REASONING_EFFORT", "OPENAI_REASONING_EFFORT"), "xhigh")
DEFAULT_TIMEOUT_SECONDS = float(_config_value("timeout_seconds", ("N9918A_AI_TIMEOUT_SECONDS",), "180"))
DEFAULT_MAX_OUTPUT_TOKENS = int(_config_value("max_output_tokens", ("N9918A_AI_MAX_OUTPUT_TOKENS",), "8192"))

AI_KEY_ENV_NAMES = (
    "N9918A_AI_API_KEY",
    "OPENAI_API_KEY",
    "ARK_API_KEY",
    "VOLCENGINE_API_KEY",
)

sys_prompt = """
你是 M5Stack 的 EMC/SA 筛查结果分析助手，面向硬件、射频、结构和整改工程师。

任务目标：
- 只分析输入表格中 Fail 点、FCC/CE Margin 为正的点，以及 Margin 接近 0 的临界点（默认 <= 2 dB）。
- 识别异常点之间的规律：倍频/分频、时钟源、DC/DC、MCU/无线模块、线缆天线效应、屏蔽/接地/结构缝隙、测试夹具或切换路径等。
- 输出可执行的工程整改建议，而不是泛泛科普或正式法规判定。

重要口径：
- 当前数据是 SA 筛查模式结果，不等同正式 FCC/CE 合规报告。
- Margin = 测量/修正值 - 参考限值；Margin > 0 表示超限风险，接近 0 表示临界风险。
- QUASI_PEAK 若来自本软件，为软件估算；正式确认应使用具备 EMI 选件的仪器检测器和完整修正链。
- 若输入缺少天线因子、线缆损耗、switchbox 损耗、前置放大器增益、测试距离或 detector 信息，必须把这类限制写入“验证前提/不确定性”。

输出格式（中文）：
1. 结论优先：用 3-5 条列出最高风险频点/频段、风险等级和最可能根因。
2. 异常点清单：表格列出频率、幅度、FCC/CE margin、状态、风险说明；不要列出无关合格点。
3. 规律归纳：说明是否存在倍频、集中频段、模块相关或线缆/结构路径相关性。
4. 根因假设：按“电源/时钟/无线或高速数字/线缆与接口/结构屏蔽/测试链路”分层列出证据和排查方法。
5. 整改建议：按优先级给出可验证动作，例如近场探头定位、断开线缆 A/B、关闭模块、改滤波/磁珠/屏蔽/接地、复测步骤。
6. 复测建议：说明下一次 SA/EMI 复测应如何确认整改有效。

约束：
- 不输出与异常点无关的法规介绍、硬件常识或大段背景说明。
- 不把筛查结果写成正式合规结论。
- 如没有 Fail 或临界点，简短说明“未发现需要 AI 展开的异常点”，并列出仍需硬件确认的前提。
""".strip()


@dataclass
class AITextResponse:
    """Small compatibility wrapper for code that previously expected choices[0].message."""

    text: str
    raw: dict

    @property
    def output_text(self):
        return self.text

    @property
    def choices(self):
        return [SimpleNamespace(message=SimpleNamespace(content=self.text))]


class ChatBot:
    def __init__(
        self,
        api_key: str = None,
        base_url: str = None,
        model: str = None,
        system_message: str = "You are a helpful assistant.",
        reasoning_effort: str = None,
        timeout_seconds: float = None,
        max_output_tokens: int = None,
    ):
        """
        Initialize the Responses API client.

        Args:
            api_key: API key. Defaults to ai_config.local.json or compatible env vars.
            base_url: OpenAI-compatible base URL. Defaults to the local proxy.
            model: Responses model name. Defaults to gpt-5.5.
            system_message: Instructions passed to the Responses API.
            reasoning_effort: Reasoning effort, defaults to xhigh for this project.
        """
        self.api_key = api_key or self._read_api_key()
        if not self.api_key:
            names = ", ".join(AI_KEY_ENV_NAMES)
            raise ValueError(f"API key is required. Set one of: {names}.")

        self.base_url = base_url or DEFAULT_BASE_URL
        self.endpoint = self._responses_endpoint(self.base_url)
        self.model = model or DEFAULT_MODEL
        self.system_message = system_message
        self.reasoning_effort = reasoning_effort or DEFAULT_REASONING_EFFORT
        self.timeout_seconds = timeout_seconds or DEFAULT_TIMEOUT_SECONDS
        self.max_output_tokens = max_output_tokens or DEFAULT_MAX_OUTPUT_TOKENS

    @staticmethod
    def _read_api_key():
        if LOCAL_CONFIG.get("api_key"):
            return LOCAL_CONFIG["api_key"]
        for name in AI_KEY_ENV_NAMES:
            value = os.getenv(name)
            if value:
                return value
        return None

    @staticmethod
    def _responses_endpoint(base_url: str) -> str:
        base = (base_url or "").strip().rstrip("/")
        if not base:
            base = "http://192.168.20.38:3000"
        if base.endswith("/responses"):
            return base
        if base.endswith("/v1"):
            return f"{base}/responses"
        return f"{base}/v1/responses"

    def _build_payload(self, message: str) -> dict:
        return {
            "model": self.model,
            "instructions": self.system_message,
            "input": message,
            "reasoning": {"effort": self.reasoning_effort},
            "max_output_tokens": self.max_output_tokens,
            "store": False,
        }

    def _post_responses(self, payload: dict) -> dict:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            self.endpoint,
            data=body,
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Responses API HTTP {exc.code}: {detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"Responses API connection failed: {exc.reason}") from exc

    @staticmethod
    def extract_output_text(data: dict) -> str:
        if not isinstance(data, dict):
            return ""
        if isinstance(data.get("output_text"), str):
            return data["output_text"].strip()

        chunks = []
        for item in data.get("output", []) or []:
            for content in item.get("content", []) or []:
                text = content.get("text")
                if isinstance(text, str):
                    chunks.append(text)
        if chunks:
            return "".join(chunks).strip()

        # Some OpenAI-compatible proxies return Chat-like or simplified shapes.
        choices = data.get("choices") or []
        if choices:
            message = choices[0].get("message") or {}
            content = message.get("content")
            if isinstance(content, str):
                return content.strip()
        if isinstance(data.get("text"), str):
            return data["text"].strip()
        return ""

    def responses_no_stream(self, message: str) -> AITextResponse:
        print(
            "AI Responses request: "
            f"model={self.model}, reasoning={self.reasoning_effort}, length={len(message)} chars"
        )
        raw = self._post_responses(self._build_payload(message))
        text = self.extract_output_text(raw)
        if not text:
            raise RuntimeError("Responses API returned no output text.")
        return AITextResponse(text=text, raw=raw)

    def chat_no_stream(self, message: str) -> AITextResponse:
        """Backward-compatible method name; internally uses Responses API, not Chat Completions."""
        return self.responses_no_stream(message)
