# -*- coding: utf-8 -*-
import argparse
import configparser
import os
import sys

import pandas as pd
from selenium import webdriver

# 取得裝置參數
def get_device_setup(config, args = None):
    parser = argparse.ArgumentParser(description = 'Device setup for testing.')
    parser.add_argument('device', type = str, nargs = '?', help = 'Platform for testing.')
    parser.add_argument('platform', type = str, nargs = '?', help = 'Browser for testing.')
    parser.add_argument('--file', type = str, nargs = '?', help = 'Direct Excel file for testing.')

    if args is None:
        args = sys.argv[1:]

    args = parser.parse_args(args)

    # 如果命令行提供了參數，使用命令行的值；否則使用 config.ini 中的值
    device = args.platform.lower() if args.device else config['Product']['device'].lower()
    platform = args.browser.lower() if args.platform else config['Product']['platform'].lower()
    file_name = args.file if args.file else None

    return device, platform, file_name

# 讀取 ini 設定檔
def get_config():
    config = configparser.ConfigParser()
    config.read('config.ini', encoding = 'utf-8')
    return config

# 獲取測試案例
def get_test_files(file_name = None):
    # 讀取當下路徑的所有測案檔案
    try:
        if not file_name:
            raw_data = {fname: pd.read_excel(fname, sheet_name=None) for fname in os.listdir('./') if fname.endswith('.xls') or fname.endswith('.xlsx')}
        else:
            raw_data = {file_name: pd.read_excel(file_name, sheet_name=None)}
        return raw_data
    except Exception as e:
        print(e)

# 確認參數有無支援
def check_device_support(device, platform, browser_options):
    if device != 'web' and platform not in browser_options:
        raise Exception('Device and Platform not supported.')
    elif device != 'web':
        raise Exception('Device not supported.')
    elif platform not in browser_options:
        raise Exception('platform not supported.')

# 設定裝置參數
def setup_driver(device, platform, implicit_wait = 2, script_timeout = 2):
    browser_options = {
        'chrome': webdriver.ChromeOptions(),    # 創建 ChromeOptions 對象
        'edge': webdriver.EdgeOptions(),        # 創建 EdgeOptions 對象
        'firefox': webdriver.FirefoxOptions()   # 創建 FirefoxOptions 對象           
        }

    if device == 'web' and platform in browser_options:
        options = browser_options[platform]
        options.add_argument('--ignore-certificate-errors')     # 無視 SSL 錯誤。添加一些選項，根據需要進行更改
        driver = getattr(webdriver, platform.capitalize())(options = options)
        driver.maximize_window()                                # 設定全螢幕
        driver.implicitly_wait(implicit_wait)                   # 設定等待時間為 n 秒
        driver.set_script_timeout(script_timeout)               # 設定執行 JavaScript 的超時時間為 n 秒
        return driver
    else:
        check_device_support(device, platform, browser_options)


