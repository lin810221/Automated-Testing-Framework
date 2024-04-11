# -*- coding: utf-8 -*-
import os
import pandas as pd
import datetime


# 匯出 Excel 測試結果
def save_to_excel(df, excel_file, folder_path='record', end_datetime = datetime.datetime.now()):
    try:
        print('Saving Excel file...')
        # 檢查文件夾是否存在，若不存在則創建
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        
        base_name, extension = os.path.splitext(excel_file)
        print(base_name)
        excel_fname = f'{base_name}_{end_datetime.strftime("%Y-%m-%d-%H-%M-%S")}.xlsx'
        excel_file_path = os.path.join(folder_path, excel_fname)
        df.to_excel(excel_file_path, index=False)
    except Exception as e:
        print(e)

# QA 內部報告
def report_QA(test_plan, excel_file, start_time, end_time = datetime.datetime.now(), folder_path = 'QA_Report'):
    try:
        print('Saving QA Report...')
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        
        # 匯入 dataframe
        df_Result = test_plan[['用例編號', '測試標題', '測試項目', '測試結果']]
        
        # 計算良率與不良率
        pass_count = df_Result['測試結果'].value_counts().get('Pass', 0)
        fail_count = df_Result['測試結果'].value_counts().get('Fail', 0)
        total_count = pass_count + fail_count
        
        pass_rate = (pass_count/total_count) * 100 if total_count > 0 else 0
        fail_rate = (fail_count/total_count) * 100 if total_count > 0 else 0
        
        title, extension = os.path.splitext(excel_file)
        head = '''
        <head>
            <meta charset="UTF-8">
            <title>QA Report</title>
            <style>
                table {
                    border-collapse: collapse;
                }
                th, td {
                    border: 1px solid black;
                    padding: 5px;
                }
                .pass {
                    background-color: green;
                    color: white;
                }
                .fail {
                    background-color: red;
                    color: white;
                }
            </style>
        </head>
        '''
        
        title_html = f'<h1>{title}</h1>'
        
        rate_html = f'<h3>良率：{pass_rate:.2f}%</h3>\
            <h3>不良率：{fail_rate:.2f}%</h3>'
        
        time_html = f'<p>開始時間：{start_time.strftime("%Y-%m-%d %H:%M:%S")}</p>\
            <p>結束時間：{end_time.strftime("%Y-%m-%d %H:%M:%S")}</p>'
        
        # 將 'Result' 欄位的值轉換為帶有對應 CSS class 的 HTML 標記
        df_Result['測試結果'] = df_Result['測試結果'].apply(lambda x: f'<td class="pass">{x}</td>' if x == 'Pass' else f'<td class="fail">{x}</td>')
        print(df_Result)
        df_html = df_Result.to_html(index=False, escape=False)
        
        df_html = df_html.replace('<td><td class="fail">Fail</td></td>', '<td class="fail">Fail</td>').replace('<td><td class="pass">Pass</td></td>', '<td class="pass">Pass</td>')
        
        html_content = f'{head}{title_html}{rate_html}{time_html}{df_html}'
        
        html_fname = f'{title}_{end_time.strftime("%Y-%m-%d-%H-%M-%S")}.html'
        html_file_path = os.path.join(folder_path, html_fname)
        with open(html_file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
    
    except Exception as e:
        print(e)















