# -*- coding: utf-8 -*-
import datetime
import json
import time
import warnings

import func_setup
import func_behavior
import func_reporting

warnings.filterwarnings('ignore')

# 主程式
def main():
    try:
        # 宣告變數
        # global raw_data
        # global excel_file
        # global sheet_data
        # global df_test_plan
        # global df_element
        # global test_steps
        # global step
        # global check_dict
        # global img_dict
        # global excel_fname

        # 讀取參數配置
        config = func_setup.get_config()

        # 讀取指令 -> python <AutoTest.py> <web> <browser> --file <Excel>
        (device, platform, file_name) = func_setup.get_device_setup(config)
        
        # 讀取測案
        raw_data = func_setup.get_test_files(file_name = file_name)
        
        # 執行測試
        # 取得檔案名稱
        for excel_file, sheet_data in raw_data.items():
            print('\033[46m 測項開始 \033[0m')
            print(f'【檔案名稱】：{excel_file}')
            df_test_plan, df_element = list(sheet_data.values())[:2]                                     # 取得前 2 個 Sheet 內容
            df_test_plan = df_test_plan[df_test_plan['測試分類'] == '自動']                               # 選取測試分類設定為自動的項目
            check_dict = dict(zip(df_element[list(df_element)[0]], df_element[list(df_element)[1]]))     # 建立元件對照表
            start_datetime = datetime.datetime.now()                                                     # 測試案例開始時間
            # 進行多個列的賦值
            df_test_plan = df_test_plan.assign(
                平台 = device,
                裝置 = platform,
                開始時間 = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                結束時間 = None,
                測項開始時間 = None,
                測項結束時間 = None,
                花費時間_秒 = None,
                用例編號 = lambda x: x['用例編號'].fillna('-').astype(str)
            )

            # 執行 <測試項目>
            for index, row in df_test_plan.iterrows():
                # 宣告變數
                issue_track = []
                img_dict = {}       # 截圖需要紀錄功能時使用
                start_time = datetime.datetime.now()
                df_test_plan.at[index, '測項開始時間'] = start_time.strftime('%Y-%m-%d %H:%M:%S')
                test_title = row['測試標題']
                test_item = row['測試項目']
                print(f'【測試標題】：{test_title}')
                print(f'【測試項目】：{test_item}')
                test_steps = row['操作步驟']
                driver = func_setup.setup_driver(device, platform)
                interval = int(config['Setup']['interval'])
                
                # 執行 <操作步驟>
                for step in test_steps.split('\n'):
                    print(f'\t{step}')
                    func_behavior.action(driver, step, check_dict = check_dict, issue_track = issue_track, img_dict = img_dict, interval = interval)
                    time.sleep(interval)
                driver.quit()
                end_time = datetime.datetime.now()
                df_test_plan.at[index, '測項結束時間'] = end_time.strftime('%Y-%m-%d %H:%M:%S')
                elapsed_time = end_time - start_time
                df_test_plan.at[index, '測試結果'] = 'Fail' if issue_track else 'Pass'  # 如果 issue_track 裡有東西，表示 Fail；反之，Pass
                df_test_plan.at[index, '問題追蹤'] = '\n'.join(issue_track)
                df_test_plan.at[index, '花費時間 (秒)'] = round(elapsed_time.total_seconds(), 2)
                img_json = json.dumps(img_dict, ensure_ascii = False)
                df_test_plan.at[index, '圖片檔名'] = img_json
            print('\033[45m 測項結束 \033[0m')
            end_datetime = datetime.datetime.now()    # 測試案例結束時間
            df_test_plan['結束時間'] = end_datetime.strftime('%Y-%m-%d %H:%M:%S')
            
            if len(df_test_plan):
                df_report = df_test_plan[['用例編號', '平台', '裝置', '開始時間', '結束時間', '測試標題', 
                                          '測試項目', '測試結果', '問題追蹤', '測項開始時間', '測項結束時間', 
                                          '花費時間 (秒)', '圖片檔名']]
                excel_fname = func_reporting.save_to_excel(df = df_report, excel_file = excel_file, end_datetime = end_datetime)
                func_reporting.report_QA(test_plan = df_test_plan, excel_file = excel_file, start_time = start_datetime, end_time = end_datetime)

    except Exception as e:
        print(e)            
        if df_test_plan is not None and excel_file:
            excel_fname = func_reporting.save_to_excel(df = df_test_plan, excel_file = excel_file)
            func_reporting.report_QA(test_plan = df_test_plan, excel_file = excel_file, start_time = start_datetime)

    else:
        print('Test completed !')
        report = func_reporting.TestReport()
        report.create_cover_page(df = df_test_plan, device = device, platform = platform, test_time = start_datetime)
        report.create_content_page(df_test_plan)
        report.save_report(excel_file = excel_file, end_time = end_datetime)

    finally:
        total_sec = 3
        for sec in range(total_sec):
            print(f'{total_sec - sec}s')
            time.sleep(1)
        print('Exit the script.')

if __name__ == '__main__':
    main()