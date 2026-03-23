from toolbox import update_ui, CatchException, report_exception, write_history_to_file, promote_file_to_downloadzone
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from crazy_functions.crazy_utils import request_gpt_model_in_new_thread_with_ui_alive
import os
import glob

@CatchException
def 成本动因分析(main_input, llm_kwargs, plugin_kwargs, chatbot, history, system_prompt, user_request):
    chatbot.append(["📊 成本动因分析任务启动", "开始读取并清洗数据..."])
    yield from update_ui(chatbot=chatbot, history=history)

    # 自动获取Excel路径
    if os.path.isdir(main_input):
        excel_files = glob.glob(os.path.join(main_input, "*.xlsx"))
        if not excel_files:
            chatbot.append(["❌ 错误", f"目录 {main_input} 中未找到Excel文件"])
            yield from update_ui(chatbot=chatbot, history=history)
            return
        filepath = excel_files[0]
    else:
        filepath = main_input

    if not os.path.exists(filepath):
        chatbot.append(["❌ 错误", f"路径不存在: {filepath}"])
        yield from update_ui(chatbot=chatbot, history=history)
        return

    # 读取与预处理
    df = pd.read_excel(filepath)
    df.columns = [col.strip() for col in df.columns]
    df.replace("–", np.nan, inplace=True)
    df = df.applymap(lambda x: str(x).replace(",", "") if isinstance(x, str) else x)
    df = df.apply(pd.to_numeric, errors="ignore")

    # 模糊字段识别
    def find_col(name_hint):
        return next((col for col in df.columns if name_hint in col), None)

    cost_col = find_col("合计金额")
    area1_col = find_col("建筑面积")
    area2_col = find_col("占地面积")

    if not all([cost_col, area1_col, area2_col]):
        msg = f"字段识别失败：\n合计金额字段：{cost_col}\n建筑面积字段：{area1_col}\n占地面积字段：{area2_col}"
        chatbot.append(["❌ 无法识别字段", msg + f"\n当前字段有：{list(df.columns)}"])
        yield from update_ui(chatbot=chatbot, history=history)
        return

    # 过滤缺失
    subdf = df[[cost_col, area1_col, area2_col]].dropna()
    cost = subdf[cost_col].astype(float)
    area1 = subdf[area1_col].astype(float)
    area2 = subdf[area2_col].astype(float)

    # 分析
    corr1 = np.corrcoef(cost, area1)[0, 1]
    corr2 = np.corrcoef(cost, area2)[0, 1]

    model1 = LinearRegression().fit(area1.values.reshape(-1, 1), cost)
    r1 = model1.score(area1.values.reshape(-1, 1), cost)

    model2 = LinearRegression().fit(area2.values.reshape(-1, 1), cost)
    r2 = model2.score(area2.values.reshape(-1, 1), cost)

    std1 = (cost / area1).std()
    std2 = (cost / area2).std()

    # 构造分析摘要交给LLM
    summary = f"""
你是一个数据分析专家。请根据以下字段润色为一段结构化的成本动因分析内容：
1. 成本与建筑面积的相关系数为：{corr1:.2f}，R²={r1:.2f}
2. 成本与占地面积的相关系数为：{corr2:.2f}，R²={r2:.2f}
3. 建筑单位成本波动（std）：{std1:.2f}；占地单位成本波动（std）：{std2:.2f}
要求用专业术语、正式风格清晰表达。
"""

    gpt_reply = yield from request_gpt_model_in_new_thread_with_ui_alive(
        inputs=summary,
        inputs_show_user="请润色成本动因分析内容",
        llm_kwargs=llm_kwargs,
        chatbot=chatbot,
        history=history,
        sys_prompt="你是物业数据分析专家，请将分析数据润色为标准化报告段落。"
    )

    chatbot.append(["✅ 分析完成（字段自适应 + 缺失处理 + LLM润色）", gpt_reply])
    yield from update_ui(chatbot=chatbot, history=history)

    history.append(("分析结果", gpt_reply))
    res = write_history_to_file(history)
    promote_file_to_downloadzone(res, chatbot=chatbot)
    chatbot.append(("📎 已生成下载文件", res))
    yield from update_ui(chatbot=chatbot, history=history)
