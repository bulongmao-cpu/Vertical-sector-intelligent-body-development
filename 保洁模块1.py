import pandas as pd

def match_column(columns, keywords):
    """
    在列名中模糊匹配包含所有关键词的列名，返回第一个匹配到的列。
    """
    for col in columns:
        if all(k in col for k in keywords):
            return col
    raise ValueError(f"❌ 未找到包含关键词 {keywords} 的列，请检查表头。")

def preprocess_cleaning_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip()

    # 自动识别列（根据你实际的表头关键字）
    col_area   = match_column(df.columns, ["建筑", "面积"])
    col_cost   = match_column(df.columns, ["保洁", "成本"])
    col_post   = match_column(df.columns, ["保洁", "岗位"])
    col_people = match_column(df.columns, ["填报", "岗位"])
    col_salary = match_column(df.columns, ["人员", "成本"])

    # 标准化字段命名
    df = df.rename(columns={
        col_area:   "校区建筑面积",
        col_cost:   "保洁成本",
        col_post:   "保洁基础岗位数",
        col_people: "填报人数",
        col_salary: "人员成本"
    })

    # 数值化 + 异常处理
    for col in ["校区建筑面积", "保洁成本", "保洁基础岗位数", "填报人数", "人员成本"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 派生字段计算
    df["平均服务面积"] = df["校区建筑面积"] / df["保洁基础岗位数"]
    df["单位面积成本"] = df["保洁成本"] / df["校区建筑面积"]
    df["人均成本"] = df["人员成本"] / df["填报人数"]

    return df

# 调试代码
if __name__ == "__main__":
    df = pd.read_excel("保洁面积试算划分表格.xlsx", header=1)
    df_processed = preprocess_cleaning_data(df)
    print(df_processed[["校区建筑面积", "保洁成本", "平均服务面积", "单位面积成本", "人均成本"]].head(30))
