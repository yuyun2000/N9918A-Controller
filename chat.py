# -*- coding: utf-8 -*-
sys_prompt = '''
你是一个专注于SA（频谱分析仪）检测报告异常频点/频段剖析的专业工具。你的唯一功能是，针对用户输入的标准化频谱检测信息——包括检测频段、测量时长、详细数据表（含频点、幅度、限值、Margin、Status等字段）以及可选频谱图片——进行全方位、多层次的异常频点及频段梳理，输出精确清单与全面技术原因分析。对所有临界点或规律性集中的异常，要求归纳潜在共性、内在机制及工程根因。

核心任务定义：

解析用户输入的标准化SA检测表格和补充频谱图片（如提供），自动定位所有出现Fail（超标）、Margin接近零（临界、特别是小于2 dB）的频点，并关注频点分布是否出现规律性（如频点聚焦于25 MHz倍数、特定模块中心频点、干扰倍频等）。
对所有异常或临界点逐一详细描述：频点、幅度、超标/临界幅度以及临近点裕量对比。
分析频点内在相关性——如同属同一系统谐波、同工程模块干扰区、线缆/PCB/电源相关分布，揭示共因，包括倍频、寄生、模块切换等现象。
针对每一异常输出详细可能成因分层，涵盖设计、工艺、结构、环境等典型因素。
如有多个异常点密集分布，综合为频段性风险并归纳整体技术原因与规律。
严格限制：

仅输出Fail及Margin≤2 dB的临界异常点，及频段（若异常点有规律聚集或倍频分布），禁止输出合格点/正常段内容。
必须涵盖所有可能的设备内部、架构、布局、电气、布线、外置干扰、标准限制相关成因，并结合异常点分布的内在逻辑。
不允许输出与异常关联无关的技术解释、标准法规文本、硬件原理普及等泛内容。
若表格或图片无法识别异常点或结构不合规范，直接回复“仅支持结构化SA检测表格和可辨频点的频谱图，请检查数据格式。”
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
可选：频谱曲线类图片，用于辅助频点分布判断。
注：仅需识别Fail点（Margin>0）及Margin≤2 dB的临界点（即便Status为Pass），关注点的集中分布、倍频类特征与工程内因，并详细输出潜在根因，严禁输出无关信息。
'''


import json
from typing import List, Dict, Any, Callable
import time
from volcenginesdkarkruntime import Ark
import os


class ChatBot:
    def __init__(
        self, 
        api_key: str = None, 
        base_url: str = "https://api.openai.com/v1", 
        model: str = "gpt-3.5-turbo",
        system_message: str = "You are a helpful assistant."
    ):
        """
        初始化ChatBot类
        
        Args:
            api_key: OpenAI API密钥，如果为None则从环境变量OPENAI_API_KEY获取
            base_url: API基础URL，可自定义为其他兼容OpenAI API的服务
            model: 使用的模型名称或推理接入点ID
            system_message: 系统预设指令
        """
        self.api_key = api_key 
        if not self.api_key:
            raise ValueError("API key is required. Either pass it directly or set OPENAI_API_KEY environment variable.")
        
        # 初始化OpenAI客户端
        self.client = Ark(
            base_url=base_url,
            api_key=self.api_key
        )
        
        self.model = model
        self.conversation_history = [{"role": "system", "content": system_message}]
        self.tools = []
        self.function_map = {}
        
    def get_system_message(self) -> str:
        """获取当前系统预设指令"""
        return self.conversation_history[0]["content"] if self.conversation_history and self.conversation_history[0]["role"] == "system" else ""
    
    def clear_history(self, keep_system_message: bool = True) -> None:
        """清除对话历史，默认保留系统预设指令"""
        if keep_system_message and self.conversation_history and self.conversation_history[0]["role"] == "system":
            self.conversation_history = [self.conversation_history[0]]
        else:
            self.conversation_history = []
 
    def chat_no_stream(self, message: str):
        self.conversation_history.append({"role": "user", "content": message})
        # 准备请求参数
        params = {
            "model": self.model,
            "messages": self.conversation_history,
            "stream": False,
            "temperature":0,
            "max_tokens":16384,
        }

        # 如果有工具定义，添加到请求中
        if self.tools:
            params["tools"] = self.tools
            params["tool_choice"] = "auto"
        # 升级方舟 SDK 到最新版本 pip install -U 'volcengine-python-sdk[ark]'
        
        # 非流式请求
        response = self.client.chat.completions.create(**params)
        
        return response

    def get_conversation_history(self) -> List[Dict[str, Any]]:
        """获取完整对话历史"""
        return self.conversation_history
    



if __name__ == "__main__":
    bot = ChatBot(
        api_key="b8800336-579b-4322-b2e9-ca0f4443db71",
        base_url="https://ark.cn-beijing.volces.com/api/v3",
        model="ep-20250708144105-dqzdw",
        system_message=sys_prompt
    )

    print(f"系统指令: {bot.get_system_message()}")
    print("-------------------------------------")

    user_input = '''
    频段:9kHZ-150kHz 测量时长：15s 测量数据：
QUASI_PEAK Mode Results:
====================================================================================================
No   Freq [MHz]   Amplitude [dBμV]   FCC Limit [dBμV]   FCC Margin [dB]    Status         
----------------------------------------------------------------------------------------------------
1    0.036        32.96              34.0               -1.04              Pass           
2    0.072        29.38              40.0               -10.62             Pass           
3    0.070        25.60              40.0               -14.40             Pass           
4    0.033        19.30              34.0               -14.70             Pass           
5    0.031        18.91              34.0               -15.09             Pass           
6    0.062        22.82              40.0               -17.18             Pass           
7    0.009        16.77              34.0               -17.23             Pass           
8    0.108        21.67              40.0               -18.33             Pass           
9    0.064        20.88              40.0               -19.12             Pass           
10   0.018        14.10              34.0               -19.90             Pass           
11   0.065        19.97              40.0               -20.03             Pass           
12   0.011        13.36              34.0               -20.64             Pass           
13   0.039        12.45              34.0               -21.55             Pass           
14   0.098        18.28              40.0               -21.72             Pass           
15   0.030        12.25              34.0               -21.75             Pass           
16   0.096        18.15              40.0               -21.85             Pass           
17   0.094        17.98              40.0               -22.02             Pass           
18   0.095        17.08              40.0               -22.92             Pass           
19   0.061        15.28              40.0               -24.72             Pass           
20   0.058        15.21              40.0               -24.79             Pass           
21   0.131        13.23              40.0               -26.77             Pass           
22   0.071        12.64              40.0               -27.36             Pass           
23   0.126        11.42              40.0               -28.58             Pass           
24   0.113        11.19              40.0               -28.81             Pass           
25   0.143        9.74               40.0               -30.26             Pass       
'''
    response = bot.chat_no_stream(user_input)
    msg_obj = response.choices[0].message
    content = msg_obj.content if hasattr(msg_obj, "content") else msg_obj.get("content", "")
    print(content)
