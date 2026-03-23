import webbrowser
from pathlib import Path

def open_cost_index(txt, llm_kwargs, plugin_kwargs, chatbot, history, *_):
    try:
        html_path = Path("static/index.html").resolve().as_uri()
        webbrowser.open_new_tab(html_path)
        chatbot.append(["📁 成本测算页面已打开", f"请在新标签页中查看：{html_path}"])
        yield "", chatbot, "", ""
    except Exception as e:
        chatbot.append(["❌ 页面打开失败", str(e)])
        yield "", chatbot, "", ""
