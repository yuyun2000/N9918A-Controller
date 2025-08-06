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
    频段:30MHZ-1GHz 测量时长：15s 测量数据：

QUASI_PEAK Mode Results:
====================================================================================================
No   Freq [MHz]   Amplitude [dBμV]   FCC Limit [dBμV]   FCC Margin [dB]    Status         
----------------------------------------------------------------------------------------------------
1    175.015      42.82              40.0               2.82               FCC Fail       
2    274.925      47.79              46.0               1.79               FCC Fail       
3    46.975       39.91              40.0               -0.09              Pass           
4    224.970      44.75              46.0               -1.25              Pass           
5    499.965      38.77              46.0               -7.23              Pass           
6    76.075       31.28              40.0               -8.72              Pass           
7    240.005      36.50              46.0               -9.50              Pass           
8    159.980      27.64              40.0               -12.36             Pass           
9    72.680       27.63              40.0               -12.37             Pass           
10   52.795       26.01              40.0               -13.99             Pass           
11   450.010      31.46              46.0               -14.54             Pass           
12   350.100      31.41              46.0               -14.59             Pass           
13   170.650      24.65              40.0               -15.35             Pass           
14   64.435       24.60              40.0               -15.40             Pass           
15   69.285       24.26              40.0               -15.74             Pass           
16   60.555       24.15              40.0               -15.85             Pass           
17   400.055      29.60              46.0               -16.40             Pass           
18   43.095       23.51              40.0               -16.49             Pass           
19   56.190       23.09              40.0               -16.91             Pass           
20   82.380       22.59              40.0               -17.41             Pass           
21   178.895      21.67              40.0               -18.33             Pass           
22   215.270      21.23              40.0               -18.77             Pass           
23   374.835      27.02              46.0               -18.98             Pass           
24   125.060      19.90              40.0               -20.10             Pass           
25   281.230      25.79              46.0               -20.21             Pass           
26   30.970       19.10              40.0               -20.90             Pass           
27   300.145      24.91              46.0               -21.09             Pass           
28   204.115      18.19              40.0               -21.81             Pass           
29   267.650      23.84              46.0               -22.16             Pass           
30   261.345      23.08              46.0               -22.92             Pass           
31   260.375      22.93              46.0               -23.07             Pass           
32   255.040      22.46              46.0               -23.54             Pass           
33   287.535      22.02              46.0               -23.98             Pass           
34   221.575      20.97              46.0               -25.03             Pass           
35   294.810      19.41              46.0               -26.59             Pass           
36   324.880      19.03              46.0               -26.97             Pass           
37   115.845      12.35              40.0               -27.65             Pass           
38   35.820       12.19              40.0               -27.81             Pass           
39   320.030      16.93              46.0               -29.07             Pass           
40   549.920      16.81              46.0               -29.19             Pass           
41   725.005      16.67              46.0               -29.33             Pass           
42   434.005      14.83              46.0               -31.17             Pass           
43   625.095      13.91              46.0               -32.09             Pass           
44   675.050      13.62              46.0               -32.38             Pass           
45   618.305      13.47              46.0               -32.53             Pass           
46   480.080      13.36              46.0               -32.64             Pass           
47   774.960      12.74              46.0               -33.26             Pass           
48   429.640      12.15              46.0               -33.85             Pass           
49   893.785      11.74              46.0               -34.26             Pass           
50   389.870      11.72              46.0               -34.28             Pass           


'''
    response = bot.chat_no_stream(user_input)
    msg_obj = response.choices[0].message
    content = msg_obj.content if hasattr(msg_obj, "content") else msg_obj.get("content", "")
    print(content)
