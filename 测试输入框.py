from toolbox import safe_update_ui
import os
import glob


def 测试输入框(main_input, llm_kwargs=None, plugin_kwargs=None, chatbot=None, history=None, system_prompt=None, user_request=None):
    if chatbot is None: chatbot = []
    if history is None: history = []

    chatbot.append(["测试插件已启动", "模拟准备读取文件..."])
    yield from safe_update_ui(chatbot, history)

    # 模拟路径处理
    if os.path.isdir(main_input):
        excel_files = glob.glob(os.path.join(main_input, "*.xlsx"))
        filepath = excel_files[0] if excel_files else "未找到文件"
    else:
        filepath = main_input

    chatbot.append(["读取路径完成", f"路径为：{filepath}"])
    yield from safe_update_ui(chatbot, history)

    # 👇 插入动态变量
    chatbot.append(["请输入一个测试参数", "$测试参数$"])
    yield from safe_update_ui(chatbot, history)

    # 👇 读取用户输入的变量（来自 plugin_kwargs）
    try:
        user_value = plugin_kwargs.get("测试参数", "")
        chatbot.append(["你输入的内容是：", str(user_value)])
    except:
        chatbot.append(["❌ 错误", "无法读取你输入的测试参数"])
    yield from safe_update_ui(chatbot, history)
