import pandas as pd

def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df["平均服务面积"] = df["校区建筑面积"] / df["保洁基础岗位数"]
    df["单位面积成本"] = df["保洁成本"] / df["校区建筑面积"]
    df["人均成本"] = df["人员成本"] / df["填报人数"]
    return df

# 加载 Excel
df = pd.read_excel("保洁面积试算划分表格.xlsx")
df_processed = preprocess(df)
print(df_processed.head())
