import concurrent
import itertools
import os.path
import sys
import os
import time
from datetime import datetime
import pandas as pd
from concurrent import  futures

def runs(df_a, df_b):
    #
    # 一对一匹配
        df_a2 = df_a.copy()
        df_b2 = df_b.copy()
        for index_a, row_a in df_a.iterrows():
            hash_b = {sum(pair): pair for pair in itertools.combinations(df_b['金额'], 1)}
            amount_a = row_a['金额']
            if amount_a in hash_b:
                pair_b = hash_b[amount_a]
                matched_rows_b = []
                for i in range(1):
                    temp_df_b = df_b.loc[(df_b['金额'] == pair_b[i])]
                    if len(temp_df_b) > 0:
                        matched_rows_a = temp_df_b.index.tolist()[0]
                        df_b.drop(index=matched_rows_a, inplace=True)
                        matched_rows_b.append(matched_rows_a)
                matched_rows_b = list(set(matched_rows_b))
                if len(matched_rows_b) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_b]
                    matched_rows_b_str = ", ".join(macheds)
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "一对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对一匹配：与b表第{matched_rows_b_str}行匹配")
                    # print(f"一对一匹配：a表第{index_a + 2}行与b表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)

        if df_a.empty:
            return df_a2, df_b2
        if df_b.empty and not df_a.empty:
            for index_a in df_a.index:
                print(f"a表第{index_a + 2}行在b表中没有匹配")
            return df_a2, df_b2

    # 一对多匹配

        # 一对2匹配

        for index_a, row_a in df_a.iterrows():
            hash_b = {sum(pair): pair for pair in itertools.combinations(df_b['金额'], 2)}
            amount_a = row_a['金额']
            if amount_a in hash_b:
                pair_b = hash_b[amount_a]
                matched_rows_b = []
                for i in range(2):
                    temp_df_b = df_b.loc[(df_b['金额'] == pair_b[i])]
                    if len(temp_df_b) > 0:
                        matched_rows_a = temp_df_b.index.tolist()[0]
                        df_b.drop(index=matched_rows_a, inplace=True)
                        matched_rows_b.append(matched_rows_a)
                matched_rows_b = list(set(matched_rows_b))
                if len(matched_rows_b) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_b]
                    matched_rows_b_str = ", ".join(macheds)
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "2对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[1]) - 2, "匹配结果"] = "2对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对2匹配：与b表第{matched_rows_b_str}行匹配")
                    # print(f"一对二匹配：a表第{index_a + 2}行与b表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)
        # 2对一匹配
        for index_b, row_b in df_b.iterrows():
            hash_a = {sum(pair): pair for pair in itertools.combinations(df_a['金额'], 2)}
            amount_b = row_b['金额']
            if amount_b in hash_a:
                pair_a = hash_a[amount_b]
                matched_rows_a = []
                for i in range(2):
                    temp_df_a = df_a.loc[(df_a['金额'] == pair_a[i])]
                    if len(temp_df_a) > 0:
                        matched_rows_b = temp_df_a.index.tolist()[0]
                        df_a.drop(index=matched_rows_b, inplace=True)
                        matched_rows_a.append(matched_rows_b)
                matched_rows_a = list(set(matched_rows_a))
                if len(matched_rows_a) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_a]
                    df_a2.loc[int(int(macheds[0]) - 2), "匹配结果"] = "2对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[1]) - 2), "匹配结果"] = "2对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_b2.loc[index_b, "匹配结果"] = (f"一对2匹配：与ERP表第{macheds}行匹配")
                    matched_rows_a_str = ", ".join(macheds)
                    # print(f"二对一匹配：a表第{matched_rows_a_str}行与b表第{index_b + 2}行匹配")
                    df_b.drop(index=index_b, inplace=True)
        # 一对3匹配

        for index_a, row_a in df_a.iterrows():
            hash_b = {sum(pair): pair for pair in itertools.combinations(df_b['金额'], 3)}
            amount_a = row_a['金额']
            if amount_a in hash_b:
                pair_b = hash_b[amount_a]
                matched_rows_b = []
                for i in range(3):
                    temp_df_b = df_b.loc[(df_b['金额'] == pair_b[i])]
                    if len(temp_df_b) > 0:
                        matched_rows_a = temp_df_b.index.tolist()[0]
                        df_b.drop(index=matched_rows_a, inplace=True)
                        matched_rows_b.append(matched_rows_a)
                matched_rows_b = list(set(matched_rows_b))
                if len(matched_rows_b) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_b]
                    matched_rows_b_str = ", ".join(macheds)
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "3对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[1]) - 2, "匹配结果"] = "3对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[2]) - 2, "匹配结果"] = "3对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对3匹配：与b表第{matched_rows_b_str}行匹配")
                    # print(f"一对三匹配：a表第{index_a + 2}行与b表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)
        # 3对一匹配
        for index_b, row_b in df_b.iterrows():
            hash_a = {sum(pair): pair for pair in itertools.combinations(df_a['金额'], 3)}
            amount_b = row_b['金额']
            if amount_b in hash_a:
                pair_a = hash_a[amount_b]
                matched_rows_a = []
                for i in range(3):
                    temp_df_a = df_a.loc[(df_a['金额'] == pair_a[i])]
                    if len(temp_df_a) > 0:
                        matched_rows_b = temp_df_a.index.tolist()[0]
                        df_a.drop(index=matched_rows_b, inplace=True)
                        matched_rows_a.append(matched_rows_b)
                matched_rows_a = list(set(matched_rows_a))
                if len(matched_rows_a) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_a]
                    df_a2.loc[int(int(macheds[0]) - 2), "匹配结果"] = "3对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[1]) - 2), "匹配结果"] = "3对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[2]) - 2), "匹配结果"] = "3对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_b2.loc[index_b, "匹配结果"] = (f"一对3匹配：与ERP表第{macheds}行匹配")
                    matched_rows_a_str = ", ".join(macheds)
                    # print(f"三对一匹配：a表第{matched_rows_a_str}行与b表第{index_b + 2}行匹配")
                    df_b.drop(index=index_b, inplace=True)
        # # # 一对4匹配
        #
        for index_a, row_a in df_a.iterrows():
            hash_b = {sum(pair): pair for pair in itertools.combinations(df_b['金额'], 4)}
            amount_a = row_a['金额']
            if amount_a in hash_b:
                pair_b = hash_b[amount_a]
                matched_rows_b = []
                for i in range(4):
                    temp_df_b = df_b.loc[(df_b['金额'] == pair_b[i])]
                    if len(temp_df_b) > 0:
                        matched_rows_a = temp_df_b.index.tolist()[0]
                        df_b.drop(index=matched_rows_a, inplace=True)
                        matched_rows_b.append(matched_rows_a)
                matched_rows_b = list(set(matched_rows_b))
                if len(matched_rows_b) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_b]
                    matched_rows_b_str = ", ".join(macheds)
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "4对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[1]) - 2, "匹配结果"] = "4对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[2]) - 2, "匹配结果"] = "4对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[3]) - 2, "匹配结果"] = "4对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对4匹配：与kyriba表第{matched_rows_b_str}行匹配")
                    # print(f"一对四匹配：a表第{index_a + 2}行与b表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)
        # # 4对一匹配
        for index_b, row_b in df_b.iterrows():
            hash_a = {sum(pair): pair for pair in itertools.combinations(df_a['金额'], 4)}
            amount_b = row_b['金额']
            if amount_b in hash_a:
                pair_a = hash_a[amount_b]
                matched_rows_a = []
                for i in range(4):
                    temp_df_a = df_a.loc[(df_a['金额'] == pair_a[i])]
                    if len(temp_df_a) > 0:
                        matched_rows_b = temp_df_a.index.tolist()[0]
                        df_a.drop(index=matched_rows_b, inplace=True)
                        matched_rows_a.append(matched_rows_b)
                matched_rows_a = list(set(matched_rows_a))
                if len(matched_rows_a) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_a]
                    df_a2.loc[int(int(macheds[0]) - 2), "匹配结果"] = "4对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[1]) - 2), "匹配结果"] = "4对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[2]) - 2), "匹配结果"] = "4对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[3]) - 2), "匹配结果"] = "4对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_b2.loc[index_b, "匹配结果"] = (f"一对4匹配：与ERP表第{macheds}行匹配")
                    matched_rows_a_str = ", ".join(macheds)
                    # print(f"四对一匹配：a表第{matched_rows_a_str}行与b表第{index_b + 2}行匹配")
                    df_b.drop(index=index_b, inplace=True)

        # # 一对5匹配
        for index_a, row_a in df_a.iterrows():
            hash_b = {sum(pair): pair for pair in itertools.combinations(df_b['金额'], 5)}
            amount_a = row_a['金额']
            if amount_a in hash_b:
                pair_b = hash_b[amount_a]
                matched_rows_b = []
                for i in range(5):
                    temp_df_b = df_b.loc[(df_b['金额'] == pair_b[i])]
                    if len(temp_df_b) > 0:
                        matched_rows_a = temp_df_b.index.tolist()[0]
                        df_b.drop(index=matched_rows_a, inplace=True)
                        matched_rows_b.append(matched_rows_a)
                matched_rows_b = list(set(matched_rows_b))
                if len(matched_rows_b) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_b]
                    matched_rows_b_str = ", ".join(macheds)
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "5对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[1]) - 2, "匹配结果"] = "5对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[2]) - 2, "匹配结果"] = "5对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[3]) - 2, "匹配结果"] = "5对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_b2.loc[int(macheds[4]) - 2, "匹配结果"] = "5对一匹配：与ERP表第" + str(index_a + 2) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对5匹配：与kyriba表第{matched_rows_b_str}行匹配")
                    # print(f"一对五匹配：a表第{index_a + 2}行与b表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)

        # # 5对一匹配
        for index_b, row_b in df_b.iterrows():
            hash_a = {sum(pair): pair for pair in itertools.combinations(df_a['金额'], 5)}
            amount_b = row_b['金额']
            if amount_b in hash_a:
                pair_a = hash_a[amount_b]
                matched_rows_a = []
                for i in range(5):
                    temp_df_a = df_a.loc[(df_a['金额'] == pair_a[i])]
                    if len(temp_df_a) > 0:
                        matched_rows_b = temp_df_a.index.tolist()[0]
                        df_a.drop(index=matched_rows_b, inplace=True)
                        matched_rows_a.append(matched_rows_b)
                matched_rows_a = list(set(matched_rows_a))
                if len(matched_rows_a) == 0:
                    pass
                else:
                    macheds = [str(i + 2) for i in matched_rows_a]
                    df_a2.loc[int(int(macheds[0]) - 2), "匹配结果"] = "5对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[1]) - 2), "匹配结果"] = "5对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[2]) - 2), "匹配结果"] = "5对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[3]) - 2), "匹配结果"] = "5对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[4]) - 2), "匹配结果"] = "5对一匹配：与kyriba表第" + str(index_b + 2) + "行匹配"
                    df_b2.loc[index_b, "匹配结果"] = (f"一对5匹配：与ERP表第{macheds}行匹配")
                    matched_rows_a_str = ", ".join(macheds)
                    # print(f"五对一匹配：a表第{matched_rows_a_str}行与b表第{index_b + 2}行匹配")
                    df_b.drop(index=index_b, inplace=True)

        #
        # 2对2匹配-->3对3-->4对4匹配--》5对5匹配--》6对6匹配
        for more_x in range(2, 7):
            # 用于存储匹配结果的列表
            matches = []
            for a_comb in itertools.combinations(df_a['金额'], more_x):
                hash_b = {sum(pair): pair for pair in itertools.combinations(df_b['金额'], more_x)}
                if sum(a_comb) in hash_b:
                    b_comb = hash_b[sum(a_comb)]
                    matched_rows_a = []
                    matched_rows_b = []
                    for i in range(more_x):
                        temp_df_a = df_a.loc[(df_a['金额'] == a_comb[i])]
                        temp_df_b = df_b.loc[(df_b['金额'] == b_comb[i])]
                        if len(temp_df_a) > 0 and len(temp_df_b) > 0:
                            matched_rows_a.append(temp_df_a.index.tolist()[0])
                            matched_rows_b.append(temp_df_b.index.tolist()[0])

                    if len(matched_rows_a) == more_x and len(matched_rows_b) == more_x:
                        matched_rows_a_str = ", ".join([str(i + 2) for i in matched_rows_a])
                        matched_rows_b_str = ", ".join([str(i + 2) for i in matched_rows_b])
                        matches.append((matched_rows_a_str, matched_rows_b_str))
                        for w in matched_rows_a:
                            df_a2.loc[w, "匹配结果"] = f"{more_x}对{more_x}匹配：与kyriba表第{matched_rows_b_str}行匹配"
                            df_a.drop(index=w, inplace=True)
                        for WS in matched_rows_b:
                            df_b2.loc[WS, "匹配结果"] = f"{more_x}对{more_x}匹配：与ERP表第{matched_rows_a_str}行匹配"
                            df_b.drop(index=WS, inplace=True)

            return df_a2, df_b2





def process_rows_parallel(df_a, df_b):
    grouped_a = df_a.groupby("日期")
    grouped_b = df_b.groupby("日期")
    with concurrent.futures.ProcessPoolExecutor(max_workers=16) as executor:
        results = [
            executor.submit(runs, group_a.copy(), group_b.copy())
            for date, group_a in grouped_a
            for b_date, group_b in grouped_b
            if date == b_date
        ]
        for r in concurrent.futures.as_completed(results):
            res_a, res_b = r.result()
            if res_a is not None:
                df_a.update(res_a)
            if res_b is not None:
                df_b.update(res_b)
    return df_a, df_b

def Run_Files():
    try:
        df_a = pd.read_excel("D:\RPADATA\GBS\P33_BankRecon_CKCN\example.xlsx", sheet_name="A")
        df_b = pd.read_excel("D:\RPADATA\GBS\P33_BankRecon_CKCN\example.xlsx", sheet_name="B")
        df_a_match, df_b_match = process_rows_parallel(df_a, df_b)

        # 保存df_a匹配结果到sheet1
        # 保存df_a匹配结果到sheet1
        df_a_match.to_excel(r"D:/RPADATA/GBS/P33_BankRecon_CKCN/result1.xlsx", sheet_name="Sheet1", index=False)

        # 保存df_b匹配结果到sheet2
        df_b_match.to_excel(r"D:/RPADATA/GBS/P33_BankRecon_CKCN/result2.xlsx", sheet_name="Sheet2", index=False)

        # writer.save()

        return "完成"
    except:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        return exc_value

if __name__ == '__main__':

    start_time = time.time()
    result = Run_Files()
    print(result)
    end_time = time.time()
    run_time = end_time - start_time
    print(f"程序运行时间为：{run_time:.2f}秒")
