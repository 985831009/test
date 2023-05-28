from itertools import combinations
import concurrent
import itertools
import os.path
import sys
import os
import time
import datetime
import pandas as pd
from concurrent import  futures
from openpyxl import Workbook

# 可行性的代码，但是运算速度很慢，3对3匹配多进程12模式下需要程序运行时间为：847.20秒
def runs(df_a, df_b):

    df_a2 = df_a.copy()
    df_b2 = df_b.copy()

    for more_x in range(1, 5):
        for more_y in range(1, 5):
            # 用于存储匹配结果的列表
            matches = []
            matched_rows_a_dict = {}
            matched_rows_b_dict = {}
            for a_comb in combinations(df_a['金额'], more_x):
                for b_comb in combinations(df_b['金额'], more_y):
                    if sum(a_comb) == sum(b_comb):
                                        matched_rows_a = []
                                        matched_rows_b = []
                                        for i in range(more_x):
                                            temp_df_a = df_a.loc[(df_a['金额'] == a_comb[i])]
                                            if len(temp_df_a) > 0:
                                                matched_rows_a1 =  temp_df_a.index.values.tolist()[0]
                                                df_a.drop(index=matched_rows_a1, inplace=True)
                                                matched_rows_a.append(matched_rows_a1)

                                        for i in range(more_y):
                                            temp_df_b = df_b.loc[(df_b['金额'] == b_comb[i])]
                                            if len(temp_df_b) > 0:
                                                matched_rows_b1 = temp_df_b.index.values.tolist()[0]
                                                df_b.drop(index=matched_rows_b1, inplace=True)
                                                matched_rows_b.append(matched_rows_b1)

                                        if len(matched_rows_a) == more_x and len(matched_rows_b) == more_y:
                                            matched_rows_a_str = ", ".join([str(i+2) for i in matched_rows_a])
                                            matched_rows_b_str = ", ".join([str(i+2) for i in matched_rows_b])
                                            matches.append((matched_rows_a_str, matched_rows_b_str))
                                            for w in matched_rows_a:
                                                df_a2.loc[w, "匹配结果"] = f"{more_x}对{more_y}匹配：与kyriba表第{matched_rows_b_str}行匹配"
                                            for ws in matched_rows_b:
                                                df_b2.loc[ws, "匹配结果"] = f"{more_y}对{more_x}匹配：与ERP表第{matched_rows_a_str}行匹配"
    return df_a2, df_b2

def process_rows_parallel(df_a, df_b):
    grouped_a = df_a.groupby("日期")
    grouped_b = df_b.groupby("日期")

    results = [] # 用于存放任务结果的临时列表
    with concurrent.futures.ProcessPoolExecutor(max_workers=32) as executor:
        for date, group_a in grouped_a:
            if date in grouped_b.groups:
                group_b = grouped_b.get_group(date)
                result = executor.submit(runs, group_a.copy(), group_b.copy())
                results.append(result)

        for result in concurrent.futures.as_completed(results):
            res_a, res_b = result.result()
            if res_a is not None:
                df_a.loc[res_a.index, :] = res_a
            if res_b is not None:
                df_b.loc[res_b.index, :] = res_b

    return df_a, df_b


def Run_Files():
    # try:
    df_a = pd.read_excel("E:/code/meralli/Bank Recon-JAPAN/example.xlsx", sheet_name="A")
    df_b = pd.read_excel("E:/code/meralli/Bank Recon-JAPAN/example.xlsx", sheet_name="B")
    df_a_match, df_b_match = process_rows_parallel(df_a, df_b)

    with pd.ExcelWriter(r"D:\RPADATA\GBS\P33_BankRecon_CKCN\Bot01\output/result6-duibi.xlsx") as writer:
        df_a_match.to_excel(writer, sheet_name="ERP", index=False)
        df_b_match.to_excel(writer, sheet_name="Kryiba", index=False)


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    print(f"程序开始时间为：{start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("开始运行4对4的匹配，CPU为I7-12700,内存大小为32G，多进程为32")
    x = Run_Files()
    end_time = datetime.datetime.now()
    print(f"程序结束时间为：{end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"程序运行时间为：{(end_time - start_time).total_seconds():.2f}秒")
