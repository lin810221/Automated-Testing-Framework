- [自動化框架](#自動化框架)
    - [自動化測試工具的選擇](#自動化測試工具的選擇)
    - [INI 參數設定](#ini-參數設定)
    - [測試案例架構](#測試案例架構)
    - [自動化語法](#自動化語法)
    - [語法定義說明](#語法定義說明)
    - [執行自動化框架](#執行自動化框架)
    - [自動化測試報告和結果分析](#自動化測試報告和結果分析)


<!-------------------------------  內 容   --------------------------------->

# 自動化框架
## 自動化測試工具的選擇
1. 程式語言：
    > Python3
1. 第三方套件：
    > pandas  
    > selenium  
1. 配置文件：
    > INI  
1. 其他
    > Excel  
    > TXT  

## INI 參數設定
```ini
[Setup]
interval = 1        ; 每操作之間的時間間隔(秒)

[Product]
device = web        ; 測試平台：web
platform = chrome   ; 瀏覽器：chrome、edge、firefox

[Account]
admin = admin
password = adminp@ssw0rd
QA_acc = qa
QA_pwd = 0000

[Report]
word = True
```

##  測試案例架構
### Excel 測試案例設計與撰寫
|用列編號|測試分類|前置條件|測試標題|測試項目|操作步驟      |預期結果  |測試結果|問題追蹤|  故事劇情  |
|-------|:-----:|:------:|-------|-------|--------------|---------|:-----:|-------|-----------|
|1      |自動   |有該帳號 |系統    |登入   |1.登入「帳號1」|1.成功登入|PASS   |      | 該帳號能登入|

### 測試案例管理
- 用列編號：唯一值，利於後續追蹤測試項目。
- 測試分類：手動、自動、或其他用途。
- 前置條件：測試前需準備的項目。
- 測試標題：根據產品來定義。
- 測試項目：根據產品來定義。
- 操作步驟：詳細描述操作的步驟，若要運行自動化測試，參考自動化撰寫語法。
- 預期結果：描述操作該步驟後所預期會看到的結果。
- 測試結果：紀錄通過與失敗。
- 問題追蹤：若有該步驟被判定為異常時，會記錄於此。
- 故事劇情：該測試的目的。

### 分頁作用
- 分頁 1 (測試案例設計頁面)
- 分頁 2 (元件對照表)
    - 功能：可以自訂義名稱，於分頁 1 的測試步驟進行引用，「目標」可放入自訂義名稱來呼叫。
    - 參數：可放入網頁元素，包含：id、name、css 選擇器、xpath等。




## 自動化語法
命名規則：<動作>「目標」<參數>
> - 動作：點擊、輸入、截圖、驗證...等。
> - 目標：元件、鍵盤、滑鼠。
> - 參數：數字、浮點數、自串。

語法定義：
|    動作    |      命名規則    |描述
|:----------:|:---------------:|:----:
|[點擊](#點擊)|點擊「目標」      |點擊介面上的元件
|[輸入](#輸入)|輸入「目標」<參數>|將參數輸入於介面上的輸入格內
|[清空](#清空)|清空「目標」      |將該元件內容清空
|[鍵盤](#鍵盤)|鍵盤「目標」      |模擬鍵盤按鍵
|[右鍵](#右鍵)|右鍵「目標」      |針對介面上的元件點擊右鍵
|[前往](#前往)|前往「目標」      |跳轉至指定的網址
|[等待](#等待)|等待「目標」      |等待 $n$ 秒
|[登入](#登入)|登入「目標」      |登入流程模組化
|[暫存](#暫存)|暫存「目標」<參數>|將介面上指定的元件名稱儲存於參數
|[驗證](#驗證)|驗證「目標」<參數>|將介面上指定的元件與參數進行相等判斷
|[比對](#比對)|比對「目標」      |確認目標的字串是否出現在該頁面上
<!--|[](#)||-->

## 語法定義說明：
### 點擊
- 語法：點擊「登入#開始」
- 定義：點擊元件。
- 說明：可單一或連續作業的點擊動作，連許點擊作業可用「#」做區隔。適用於按鈕、輸入框等元件。

### 輸入
- 語法：輸入「account」admin
- 定義：點擊「account」元件 → 清空內容 → 輸入 admin 字串。
- 說明：適用於輸入框的元件。

### 清空
- 語法：清空「帳號」
- 定義：點擊「帳號」 元件 → 清空內容
- 說明：適用於輸入框元件。

### 鍵盤
- 語法：鍵盤「delete」
- 定義：鍵盤的「Del」按鍵。
- 說明：目前按鍵有「Enter」、「Delete」、「Tab」。適用於網頁測試。

### 右鍵
- 語法：右鍵「account」
- 定義：點擊「account」元件 → 點擊滑鼠右鍵。
- 說明：適用於網頁測試。

### 前往
- 語法：前往「x網址」
- 定義：網頁跳轉到該網址。
- 說明：適用於瀏覽器裝置。

### 等待
- 語法：等待「15」
- 定義：於此刻停留 15 秒。
- 說明：適用於需等待運算的頁面。

### 登入
- 語法：登入「admin#123456」
- 定義：輸入帳號為 admin → 輸入密碼為 123456 → 點擊登入按鈕。
- 說明：此為特例模組，適用於以上流程的功能。

### 暫存
- 語法：暫存「//*[@id="app"]//nav/ol/li[1]/span」存一下
- 定義：讀取該元件的名稱 → 儲存於 <存一下>
- 說明：將特定元件的文字儲存於自定義的名稱內，可在該測項拿出來做使用。適用於搭配驗證或比對功能。

### 驗證
- 語法：驗證「x台達電代號」2308
- 定義：該元件的名稱與 2308 做相等比較。
- 說明：會判斷測試結果與截圖。

### 比對
- 語法：比對「台達電#2308」
- 定義：比對「台達電」與「2308」是否出現在該網頁內容。
- 說明：會判斷測試結果與截圖。可單一比對或多項式比對，多項式比對時請用「#」做區隔。

## 執行自動化框架
可點擊「AutoTest.py」運行，會將當前資料夾下所有的測案都進行測試；也可運用「命令提示字元」來運行，指令規則如下

- 執行以下指令，參數會依照 ini 檔帶入，測試案例 excel 會全部被執行。
```bash
python AutoTest.py
```
- 執行以下指令來自行設定參數  
    - \<device>：測試平台 (web)。
    - \<platform>：瀏覽器 (chrome、edge、firefox)。
    - --\<file>：運行指定單一檔名；若沒設定此參數，就是運行當前所有測試案例。

```bash
python AutoTest.py <device> <platform> --<file>
```

## 自動化測試報告和結果分析
運行完自動化測試後，會於該資料夾產出一份「軟體測試報告書」 word 檔、「QA_Report」 資料夾、「img」 資料夾與「record」資料夾  

- 測試報告書：紀錄測試結果、驗證項目與截圖。
- QA_Report：簡易的測試報告內容。
- img：紀錄截圖。
- record：將原本讀入的測案，填寫測試結果與問題追蹤



