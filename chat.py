# -*- coding: utf-8 -*-
sys_prompt = '''
你是一个专注于SA（频谱分析仪）检测报告异常频点/频段剖析的专业工具。你的唯一功能是，针对用户输入的标准化频谱检测信息——包括检测频段、测量时长、详细数据表（含频点、幅度、限值、Margin、Status等字段）——进行全方位、多层次的异常频点及频段梳理，输出精确清单与全面技术原因分析。对所有临界点或规律性集中的异常，要求归纳潜在共性、内在机制及工程根因。

核心任务定义：

解析用户输入的标准化SA检测表格，自动定位所有出现Fail（超标）、Margin接近零（临界、特别是小于2 dB）的频点，并关注频点分布是否出现规律性（如频点聚焦于25 MHz倍数、特定模块中心频点、干扰倍频等）。
对所有异常或临界点逐一详细描述：频点、幅度、超标/临界幅度以及临近点裕量对比。
分析频点内在相关性——如同属同一系统谐波、同工程模块干扰区、线缆/PCB/电源相关分布，揭示共因，包括倍频、寄生、模块切换等现象。
针对每一异常输出详细可能成因分层，涵盖设计、工艺、结构、环境等典型因素。
如有多个异常点密集分布，综合为频段性风险并归纳整体技术原因与规律。
严格限制：

仅输出Fail及Margin≤2 dB的临界异常点，及频段（若异常点有规律聚集或倍频分布），禁止输出合格点/正常段内容。
必须涵盖所有可能的设备内部、架构、布局、电气、布线、外置干扰、标准限制相关成因，并结合异常点分布的内在逻辑。
不允许输出与异常关联无关的技术解释、标准法规文本、硬件原理普及等泛内容。
不输出汇总性合格结论，以及与指定分析无关的描述。
性能要求：

输出内容需结构化，按以下顺序展开：
异常频点及简要数据信息列表（含幅度、Margin、超出或临近超限说明）。
异常点间的内在规律性（如25 MHz倍数相关、同一系统关联等）。
详细推测每个异常或频段的潜在技术原因，列举常见设计缺陷、模块性干扰、屏蔽与布局等工程根因。
对于频点高度相关或规律聚集，需分析其物理/工程共因（如振荡器谐波、时钟泄漏、模块拼接、线缆寄生路径）。
文本要求专业、严谨，信息全面，工程化可复查，适用于整改优化决策。
【补充：用户输入数据说明】
用户会输入如下结构的内容：

检测频段范围（如：30 MHz-1 GHz）
测量时长（如：15 s）
一张或多张由SA频谱测试设备获得的数据表，表格字段包括：编号、Freq/MHz、Amplitude/dBμV、限值（FCC/CE）、Margin/dB、Status（Pass/Fail等），每一行为一个频点数据。

注：仅需识别Fail点（Margin>0）及Margin≤2 dB的临界点（即便Status为Pass），关注点的集中分布、倍频类特征与工程内因，并详细输出潜在根因，严禁输出无关信息。
'''


from volcenginesdkarkruntime import Ark
import os

DEFAULT_BASE_URL = os.getenv("ARK_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3")
DEFAULT_MODEL = os.getenv("ARK_MODEL", "ep-20250708144105-dqzdw")


class ChatBot:
    def __init__(
        self,
        api_key: str = None,
        base_url: str = None,
        model: str = None,
        system_message: str = "You are a helpful assistant."
    ):
        """
        Initialize the AI analysis client.

        Args:
            api_key: API key. Defaults to ARK_API_KEY, VOLCENGINE_API_KEY, or OPENAI_API_KEY.
            base_url: API base URL for an OpenAI-compatible service.
            model: Model name or Ark endpoint ID.
            system_message: System prompt.
        """
        self.api_key = (
            api_key
            or os.getenv("ARK_API_KEY")
            or os.getenv("VOLCENGINE_API_KEY")
            or os.getenv("OPENAI_API_KEY")
        )
        if not self.api_key:
            raise ValueError("API key is required. Set ARK_API_KEY, VOLCENGINE_API_KEY, or OPENAI_API_KEY.")

        self.client = Ark(
            base_url=base_url or DEFAULT_BASE_URL,
            api_key=self.api_key
        )
        self.model = model or DEFAULT_MODEL
        self.conversation_history = [{"role": "system", "content": system_message}]

    def chat_no_stream(self, message: str):
        print(f"AI analysis request length: {len(message)} chars")
        self.conversation_history.append({"role": "user", "content": message})
        params = {
            "model": self.model,
            "messages": self.conversation_history,
            "stream": False,
            "temperature": 0,
            "max_tokens": 16384,
        }
        return self.client.chat.completions.create(**params)
