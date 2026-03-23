import requests

def 知识库问答(*args):
    print("\U0001F4E5 插件实际收到参数 args:", args)

    if not args or not args[0]:
        return None, [["用户", ""] , ["助手", "[错误] 未收到输入"]], "", ""

    text = args[0]
    print("\U0001F4E8 用户输入：", text)

    # ✅ 使用你的 Chat 应用密钥
    api_key = "app-3ahepjGhcu9um47yzSjvZBdG"
    api_url = "http://localhost:5001/v1/chat-messages"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    payload = {
        "inputs": {},
        "query": text,
        "response_mode": "blocking",
        "conversation_id": "",
        "user": "gpt-academic"
    }

    print("\U0001F4E4 实际发送 payload：", payload)

    try:
        response = requests.post(api_url, headers=headers, json=payload, timeout=90)
        print("🟦 状态码: ", response.status_code)
        print("📩 原始响应文本: ", response.text)

        try:
            result = response.json()
        except Exception as json_error:
            print("❌ JSON解析失败: ", json_error)
            return None, [["用户", text], ["助手", f"[响应非JSON] {response.text}"]], "", ""

        answer = result.get("answer", "[Dify未返回answer字段]")
        return None, [["用户", text], ["助手", answer]], "", ""

    except Exception as e:
        print("❌ 请求异常: ", e)
        return None, [["用户", text], ["助手", f"[请求失败] {e}"]], "", ""
