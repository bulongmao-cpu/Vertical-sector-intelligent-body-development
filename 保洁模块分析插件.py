from fastapi import UploadFile
from plugins import PluginFunction
from 保洁模块1 import preprocess_cleaning_data
import pandas as pd


@PluginFunction(name="保洁成本动因分析")
def analyze_cleaning_cost(file: UploadFile):
    # 读取上传的Excel文件
    df = pd.read_excel(file.file, header=1)

    # 模块1：预处理
    df = preprocess_cleaning_data(df)
    yield "✅ 数据预处理完成"

    # 模块2：模型建议区间（待实现）
    suggested_ranges, analysis = model_suggest_ranges(df)
    yield f"📐 区间划分建议：\n{analysis}"

    # 模块3：分段统计（待实现）
    stats = calculate_by_area_range(df, suggested_ranges)
    yield f"📊 分析结果如下：\n{stats}"

    # 模块4：趋势总结（待实现）
    summary = model_generate_summary(stats)
    yield f"📋 成本趋势总结：\n{summary}"

    # 模块5：Word 导出（待实现）
    docx_path = export_to_word(df, stats, summary)
    yield f"📄 点击下载完整报告：[下载]({docx_path})"
