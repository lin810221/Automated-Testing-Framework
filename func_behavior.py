# -*- coding: utf-8 -*-
import re

import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import func_setup


# 測試步驟 - 正規表達式
def find_condition(text):
    try:
        pattern = r'(\S+)「([^」]+)」(\S*)'
        matches = re.findall(pattern, text)
        result = [match for match in matches[0]]
        return result
    except:
        raise Exception('Test step format error !')


# 搜尋標籤
def find_element_by_tag(driver, des, locators=[By.XPATH, By.ID, By.CSS_SELECTOR, By.CLASS_NAME]):
    if des.startswith("/") or des.startswith("("):
        element = driver.find_element(By.XPATH, des)
        return element

    xpath_options = [f'//span[text()="{des}"]', f'//div[text()="{des}"]', 
                     f'//input[@placeholder="{des}"]', f'//input[@value="{des}"]']
    
    for locator in locators:
        xpath = " | ".join(xpath_options) if locator == By.XPATH else des
        try:
            print(f'\t\tLOCATOR: {locator}')
            #print(f'XPATH: {xpath}')
            element = driver.find_element(locator, xpath)
            if element:
                return element
        except:
            pass
    return None

def action(driver, step, interval = 1.5,
           config = None, check_dict = {}, cache = {}, issue_track = None):
    try:
        step_origin = step
        step = find_condition(text = step)
        event = step[0]
        value = step[1:3]
        #print(f'\t\tEVENT: {event}')
        #print(f'\t\tVALUE: {value}')
        aim = check_dict.get(value[0], value[0])    # 如果對照表有該項目, 返回對應的值
        #print(f'AIM: {aim}')
        
        if '點擊' in event:
            value_list = value[0].split('#')
            for value in value_list:
                try:
                    aim = cache.get(value, check_dict.get(value, value))
                    print(f'\t點擊「{aim}」')
                    element = find_element_by_tag(driver, aim)
                    element.click()
                    time.sleep(interval)
                except:
                    print(f'Error: {aim}')
        
        elif '輸入' in event:
            element = find_element_by_tag(driver, aim)
            element.click()
            element.clear()
            element.send_keys(value[1])
        
        elif '植入' in event:
            element = find_element_by_tag(driver, aim)
            element.click()
            element.clear()
            element.send_keys(value[1])
        
        elif '鍵盤' in event:
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            
        elif '前往' in event:
            driver.get(aim)
        
        elif '等待' in event:
            time.sleep(int(value[0]))
        
        elif '登入' in event:
            aim_split = aim.split('#')
            config = func_setup.get_config()
            account = config['Account'][aim_split[0]]
            password = config['Account'][aim_split[1]]

            action(driver, step = f'輸入「username」{account}'); time.sleep(interval)
            action(driver, step = f'輸入「password」{password}'); time.sleep(interval)
            action(driver, step = '點擊「kc-login」')
            
        # 暫存資料
        elif '暫存' in event:
            element = find_element_by_tag(driver, aim)
            element = element.text
            print('暫存了一筆資料： ', element)
            cache[value[1]] = element
            print(value[1] + ':' + element)

        elif '驗證' in event:
            test_criteria = cache.get(value[1]) or value[1]
            element = find_element_by_tag(driver, aim).text
            print(f'讀取訊息：{element}')
            print(f'判讀標準：{test_criteria}')
            if element == test_criteria:
                print('\033[42m\033[97m \t\tPASS\t\t \033[0m')
            else:
                print('\033[41m\033[97m \t\tFAIL\t\t \033[0m')
                issue_track.append(step_origin)

        elif '比對' in event:
            page_content = driver.page_source
            verify_items = aim.split('#')
            expected_result = value[1] == '否'   # True or False
            cache = cache
            print(cache)
            for verify_item in verify_items:
                item = cache.get(verify_item, verify_item)
                if (item not in page_content) == expected_result:
                    print('\033[42m\033[97m \t\tPASS\t\t \033[0m')
                else:
                    print('\033[41m\033[97m \t\tFAIL\t\t \033[0m')
                    issue_track.append(step_origin)
        else:
            print('Event not supported')
    except:
        print('Error')