
import pandas as pd
import re
import requests

COST_PATH = "crazy_functions/cost_standard.csv"
TEMPLATE_PATH = "crazy_functions/formula_template.csv"

def read_csv_auto(path):
    try:
        return pd.read_csv(path, encoding="utf-8")
    except:
        return pd.read_csv(path, encoding="gbk")

def build_prompt_from_tables(txt):
    cost_df = read_csv_auto(COST_PATH)
    formula_df = read_csv_auto(TEMPLATE_PATH)

    cost_df.columns = cost_df.iloc[0]
    cost_df = cost_df[1:].reset_index(drop=True)
    cost_df.columns = ["模块", "校区类型", "岗位类型", "配置标准", "人均服务面积/单位成本", "人均成本"]
    cost_df.fillna("", inplace=True)

    formula_df.columns = formula_df.iloc[0]
    formula_df = formula_df[1:].reset_index(drop=True)

    prompt = ""

    # 保洁岗逻辑
    if "保洁" in txt and "保安" not in txt:
        prompt += "你是一位高校后勤成本测算专家。\n\n"
        prompt += "【保洁岗位测算规则】\n"
        prompt += "- 建筑面积在 0 至 36000 平方米：人均服务面积 = 2305 平方米/人，人均成本 = 5.85 万元/人。\n"
        prompt += "- 建筑面积在 36000 至 82000 万平方米：人均服务面积 = 2564 平方米/人，人均成本 = 6 万元/人。\n"
        prompt += "测算公式如下：\n"
        prompt += "1. 岗位人数（人） = 建筑面积 （平方米）÷ 人均服务面积（平方米/人）\n"
        prompt += "2. 年人力成本（万元） = 岗位人数（人） × 人均成本（万元/人）\n"
        prompt += "3. 总成本（万元）= 年人力成本（万元） + 物料消耗（万元）\n\n"
        prompt += f"请根据上述规则，结合用户输入“{txt}”，推理各项测算值，依次输出：1. 岗位人数计算过程（含四舍五入说明）；2. 年总人力成本计算过程；3. 总成本计算过程；4. 所有步骤最终结果核查与结论总结。"
        
        return prompt

    # 保安岗逻辑
    cost_descriptions = [
        f"- {row['校区类型']} 的 {row['岗位类型']} 配置为 {row['配置标准']}，人均成本为 {row['人均成本']} 万元。"
        for _, row in cost_df.iterrows()
        if row['模块'] == "保安" and row['校区类型'] and row['岗位类型']
    ]

    formula_descriptions = [
        f"{col.strip()}：{str(val).strip()}" for col, val in zip(formula_df.columns, formula_df.iloc[0])
    ]

    prompt = """你是人力成本建模分析专家。
    【场景描述】
    对“保安岗”进行建模测算。校区分为多门校区和单门校区：
    - 多门校区包含：主门门岗、副门门岗、监控、消控岗、管理岗等；
    - 单门校区包含：单门门岗、监控、消控岗、管理岗等。

    【重要原则】
    以下公式是不可更改的固定定律，不允许对公式结构、变量顺序进行任何调整：
    1. 岗位数 = 校门数 × 配置标准
    2. 岗位人数（人） = 岗位数 × 24 × 30 ÷ 8 ÷ 21.75
    3. 年人力成本 = 岗位人数（人） × 人均成本（万元/人/年）
    4. 总成本 = 所有岗位的年人力成本之和
    请只计算用户输入中提到的岗位，忽略未提及的其他岗位，未提及的巡逻岗、值班岗、管理岗数量设为0。
    请严格遵循以上公式结构。所有变量仅在表格中定义，不需要说明来源，也不要尝试改写公式。

    【岗位配置与成本标准】
    1. 多门校区 - 主门门岗：配置标准为 2 岗/门，人均成本为 6.6 万元/人/年；
    2. 多门校区 - 副门门岗：配置标准为 1.5 岗/门，人均成本为 6.6 万元/人/年；
    3. 监控、消控岗：配置标准为 2 岗/间，人均成本为 6.6 万元/人/年；
    4. 管理岗：配置标准为 1 岗/校区，人均成本为 6.6 万元/人/年。


    【岗位数说明】
    当用户输入 “只计算保安岗位，校区类型为多门校区，含有3个主门和2个副门”，其中“3个主门”即表示主校门数量，必须代入公式计算岗位数 → 岗位人数 → 年人力成本。其中“3个副门”即表示副校门数量，必须代入公式计算岗位数 → 岗位人数 → 年人力成本。
    请注意，岗位数计算时必须严格按照用户输入的校区类型和配置标准进行，计算年人力成本时一定得使用岗位人数（人）代入计算。

    【任务】
    请根据上述信息，处理以下用户输入：

    【用户输入】：
    """ + txt + """

    请你依次输出：
    1. 每一项岗位的岗位数计算过程；
    2. 岗位人数计算过程（含四舍五入说明）；
    3. 年人力成本计算过程，计算年人力成本时一定得使用岗位人数（人）及人均成本6.6（万元/人/年）代入计算；
    4. 每项最终成本与总成本合计；
    5. 所有步骤最终结果核查与结论总结。
    """
    return prompt

def call_llm(prompt):
    payload = {
        "model": "deepseek-r1:1.5b",
        "prompt": prompt,
        "stream": False
    }
    try:
        response = requests.post("http://localhost:11434/api/generate", json=payload, timeout=999)
        return response.json().get("response", "").strip()
    except Exception as e:
        return f"❌ 本地 LLM 调用失败：{e}"

def formula_prompt_infer_wrapper(txt, llm_kwargs, plugin_kwargs, chatbot, history, system_prompt, request):
    try:
        prompt = build_prompt_from_tables(txt)
        reply = call_llm(prompt)
    except Exception as e:
        reply = f"❌ 插件执行失败：{e}"
    chatbot.append(["📊 成本建模推理输出", reply])
    yield "", chatbot, "", reply
