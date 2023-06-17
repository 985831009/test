import datetime
import itertools
import os.path
import os
import concurrent.futures
import numpy as np
import pandas as pd
import win32com.client
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
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "一对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对一匹配：与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    # print(f"一对一匹配：a表第{index_a + 2}行与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)
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
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "2对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_b2.loc[int(macheds[1]) - 2, "匹配结果"] = "2对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对2匹配：与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    # print(f"一对二匹配：a表第{index_a + 2}行与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)

        # 一对3匹配
        #
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
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "3对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_b2.loc[int(macheds[1]) - 2, "匹配结果"] = "3对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_b2.loc[int(macheds[2]) - 2, "匹配结果"] = "3对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对3匹配：与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    # print(f"一对三匹配：a表第{index_a + 2}行与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)
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
                    df_b2.loc[int(macheds[0]) - 2, "匹配结果"] = "4对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_b2.loc[int(macheds[1]) - 2, "匹配结果"] = "4对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_b2.loc[int(macheds[2]) - 2, "匹配结果"] = "4对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_b2.loc[int(macheds[3]) - 2, "匹配结果"] = "4对一匹配：与残高照合表第" + str(index_a + 12) + "行匹配"
                    df_a2.loc[index_a, "匹配结果"] = (f"一对4匹配：与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    # print(f"一对四匹配：a表第{index_a + 2}行与97524 SAP勘定明細表第{matched_rows_b_str}行匹配")
                    df_a.drop(index=index_a, inplace=True)

            #     # 2对一匹配
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
                    df_a2.loc[int(int(macheds[0]) - 2), "匹配结果"] = "2对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[1]) - 2), "匹配结果"] = "2对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_b2.loc[index_b, "匹配结果"] = (f"一对2匹配：与残高照合表第{macheds}行匹配")
                    matched_rows_a_str = ", ".join(macheds)
                    # print(f"二对一匹配：a表第{matched_rows_a_str}行与97524 SAP勘定明細表第{index_b + 2}行匹配")
                    df_b.drop(index=index_b, inplace=True)

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
                    df_a2.loc[int(int(macheds[0]) - 2), "匹配结果"] = "3对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[1]) - 2), "匹配结果"] = "3对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[2]) - 2), "匹配结果"] = "3对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_b2.loc[index_b, "匹配结果"] = (f"一对3匹配：与残高照合表第{macheds}行匹配")
                    matched_rows_a_str = ", ".join(macheds)
                    # print(f"三对一匹配：a表第{matched_rows_a_str}行与97524 SAP勘定明細表第{index_b + 2}行匹配")
                    df_b.drop(index=index_b, inplace=True)
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
                    df_a2.loc[int(int(macheds[0]) - 2), "匹配结果"] = "4对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[1]) - 2), "匹配结果"] = "4对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[2]) - 2), "匹配结果"] = "4对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_a2.loc[int(int(macheds[3]) - 2), "匹配结果"] = "4对一匹配：与97524 SAP勘定明細表第" + str(
                        index_b + 2) + "行匹配"
                    df_b2.loc[index_b, "匹配结果"] = (f"一对4匹配：与残高照合表第{macheds}行匹配")
                    matched_rows_a_str = ", ".join(macheds)
                    # print(f"四对一匹配：a表第{matched_rows_a_str}行与97524 SAP勘定明細表第{index_b + 2}行匹配")
                    df_b.drop(index=index_b, inplace=True)
        # 2对2匹配-->3对3
        for more_x in range(2, 4):
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
                        matched_rows_a_str = ", ".join([str(i + 12) for i in matched_rows_a])
                        matched_rows_b_str = ", ".join([str(i + 2) for i in matched_rows_b])
                        matches.append((matched_rows_a_str, matched_rows_b_str))
                        for w in matched_rows_a:
                            df_a2.loc[w, "匹配结果"] = f"{more_x}对{more_x}匹配：与97524 SAP勘定明細表第{matched_rows_b_str}行匹配"
                            df_a.drop(index=w, inplace=True)
                        for WS in matched_rows_b:
                            df_b2.loc[WS, "匹配结果"] = f"{more_x}对{more_x}匹配：与残高照合表第{matched_rows_a_str}行匹配"
                            df_b.drop(index=WS, inplace=True)

        print(str(df_a2['Date'].iloc[len(df_a2) - 1])+f"Finishtime:{start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        return df_a2, df_b2

def process_rows_parallel(df_a, df_b):
    grouped_a = df_a.groupby('Date')
    grouped_b = df_b.groupby("転記日付")
    max_workers = os.cpu_count()  # 获取CPU核心数
    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
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
#
# def process_rows_parallel(df_a, df_b):
#     grouped_a = df_a.groupby('Date')
#     grouped_b = df_b.groupby("転記日付")
#
#     common_dates = set(grouped_a.groups.keys()) & set(grouped_b.groups.keys())  # 计算日期交集
#     tasks = [(date, grouped_a.get_group(date).copy(), grouped_b.get_group(date).copy()) for date in common_dates]  # 创建任务列表
#     print(common_dates)
#     max_workers = os.cpu_count()  # 获取CPU核心数
#     with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
#         futures = [executor.submit(runs, group_a, group_b) for date, group_a, group_b in tasks]  # 异步提交任务
#
#         for future in concurrent.futures.as_completed(futures):
#             res_a, res_b = future.result()
#             if res_a is not None:
#                 df_a.update(res_a)
#             if res_b is not None:
#                 df_b.update(res_b)
#             del futures[future]  # 删除对future的引用
#     return df_a, df_b
#

# 定义一个函数来转换字符串
def convert_string(s):
    if '-' in s:
        return -int(s.replace(',', '').replace('-', ''))
    else:
        return int(s.replace(',', ''))

def convert_date(num_date: int) -> str:
    str_date = str(num_date)
    year = str_date[:1]  # 修改为只取一个字符作为年份
    if "5" in year:
        year="3"
    month = str_date[1:3]  # 修改为取两个字符作为月份
    day = str_date[3:5]  # 修改为取两个字符作为日期
    return f"202{year}/{month}/{day}"


def format_date(df, col_names):
    """
    将 DataFrame 中指定列的日期格式转换为 YYYY/MM/DD 格式

    :param df: DataFrame 对象
    :param col_names: 需要转换格式的列名列表
    :return: 转换格式后的 DataFrame 对象
    """
    for col_name in col_names:
        # 将指定列转换为日期格式
        df[col_name] = pd.to_datetime(df[col_name])

        # 将日期格式转换为 YYYY/MM/DD 格式
        df[col_name] = df[col_name].dt.strftime('%Y/%m/%d')

    return df
def Run_Files():
    # 读取Excel文件
    # try:
        file_path = r"D:\RPADATA\GBS\P35_Bank Recon-Japan\Bot01\temp\tempnew_template2.xlsx"  # 请将此处替换为你的Excel文件路径
        sheet_name = '残高照合'  # 请将此处替换为你的工作表名称

        # 使用pywin32读取Excel文件
        excel = win32com.client.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(file_path, ReadOnly=False)  # 打开文件并设置为可写
        worksheet = workbook.Worksheets(sheet_name)

        # 获取工作表的行数和列数
        row_count = worksheet.UsedRange.Rows.Count
        col_count = worksheet.UsedRange.Columns.Count

        # 新增匹配结果列
        worksheet.Cells(10, col_count + 1).Value = '匹配结果'
        for i in range(11, row_count + 1):
            worksheet.Cells(i, col_count + 1).Value = ''

        # 读取工作表数据
        data = [[worksheet.Cells(row, col).Value for col in range(1, col_count + 1)] for row in range(11, row_count + 1)]

        # 将数据转换为pandas DataFrame
        headers = data.pop(0)
        df_a = pd.DataFrame(data, columns=headers)

        # 读取另一个工作表
        df_b = pd.read_excel(file_path, sheet_name="97524 SAP勘定明細")

        # 新增匹配结果列
        worksheet_b = workbook.Worksheets("97524 SAP勘定明細")
        worksheet_b.Cells(1, col_count + 1).Value = '匹配结果'
        for i in range(2, df_b.shape[0] + 2):
            worksheet_b.Cells(i, col_count + 1).Value = ''

        # 找到Current列值为None的第一行的索引
        first_none_index = df_a[df_a['Currency'].isnull()].index[0]

        # 删除从该行开始的所有行
        df_a = df_a.drop(df_a.index[first_none_index:])

        df_a["Date"] = df_a["Date"].apply(convert_date)
        df_a = df_a.drop(columns=[col for col in df_a.columns if col is None])
        #
        df_a['匹配结果'] = np.nan

        df_b['匹配结果'] = np.nan
        # df_b['国内通貨'] = df_b['国内通貨'].apply(convert_string)
        df_a_split = df_a[['Date', 'Amount', '匹配结果']].copy()
        df_b_split = df_b[["転記日付", "国内通貨", '匹配结果']].copy()
        df_b_split = format_date(df_b_split, ["転記日付"])
        df_b_split = df_b_split.rename(columns={'国内通貨': '金额'})
        df_a_split = df_a_split.rename(columns={'Amount': '金额'})
        df_a_match, df_b_match = process_rows_parallel(df_a_split, df_b_split)

        df_a["匹配结果"] = df_a_match["匹配结果"]
        df_b['匹配结果'] = df_b_match["匹配结果"]
        df_a = df_a.drop(columns=[col for col in df_a.columns if col is None])
        df_a["匹配结果"] = df_a["匹配结果"].fillna('')
        df_b["匹配结果"] = df_b["匹配结果"].fillna('')


        # df_b["国内通貨"]=df_b_match["金额"]
        # 将df_a和df_b的匹配结果写入原始表格
        worksheet.Cells(11,15).Value="匹配结果"
        for i, r in df_a.iterrows():
            worksheet.Cells(i + 12, 15).Value = r['匹配结果']

        worksheet_b.Cells(1, 34).Value = "匹配结果"
        for i, r in df_b.iterrows():
            worksheet_b.Cells(i + 2, 34).Value = r['匹配结果']
            worksheet_b.Cells(i + 2, 28).Value = df_b.loc[i, '国内通貨']
        worksheet.Columns.AutoFit()
        worksheet_b.Columns.AutoFit()
        # 保存并关闭工作簿并退出Excel
        workbook.Save()
        workbook.Close()
        excel.Quit()
    #
    # except:
    #     exc_type, exc_value, exc_traceback = sys.exc_info()
    #     return exc_value
if __name__ == '__main__':

    start_time = datetime.datetime.now()
    print(f"程序开始时间为：{start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("开始运行4对4的匹配，多进程:20")
    x = Run_Files()
    end_time = datetime.datetime.now()
    print(f"程序结束时间为：{end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"程序运行时间为：{(end_time - start_time).total_seconds():.2f}秒")
