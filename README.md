# 自動化出缺席系統

這是一個自動化的出缺席系統，利用電子郵件進行學生點名。該系統自動生成隨機密碼，需要教師在課堂上製作紙條或分配給教室內學生一人一組密碼。並由系統發送點名通知給已知學生，學生需要在限定時間內回覆正確的密碼。系統會根據回覆情況標記學生為出席、遲到或缺席。

## 目錄

- [介紹](#介紹)
- [安裝](#安裝)
- [使用方法](#使用方法)
- [執行流程](#執行流程)
- [文件結構](#文件結構)
- [注意事項](#注意事項)

## 介紹

此程式將自動向已知學生發送點名郵件，學生需在指定時間內回覆正確的點名密碼。系統將根據學生的回覆記錄出席、遲到或缺席狀況。
此系統利用SMTP和IMAP來處理電子郵件的發送和接收，並使用Python語言來實現自動化處理，包括隨機生成出席密碼、確認學生是否正確回覆並將結果整理成Excel表格。

## 安裝

1. 複製這個資料夾到您的本地環境：
    ```bash
    git clone https://github.com/183cm-man/Python-Final-Project.git
    cd Python-Final-Project
    ```

2. 安裝所需的依賴包：
    ```bash
    pip install -r requirements.txt
    ```

3. 確保您在同一目錄下有一個名為 `known_students.xlsx` 的Excel文件，該文件應包含一個 "Number" 和 "Email" 列，存儲所有學生的編號與電子郵件地址。

## 使用方法

1. 配置您的SMTP和IMAP設置，在 `checkAttendance.py` 文件中更新以下變數：
    ```python
    SMTP_USER = "您的電子郵件地址"
    SMTP_PASSWORD = "您的電子郵件密碼"
    SMTP_SERVER = "smtp.gmail.com"
    IMAP_SERVER = "imap.gmail.com"
    ```

2. 執行程式：
    ```bash
    python checkAttendance.py
    ```

3. 程式將自動發送點名郵件，並在3分鐘內檢查學生的回覆，最終結果將記錄在 `Attendance.xlsx` 文件中。

## 執行流程：


1. 初始化 Excel 文件。

2. 生成隨機點名密碼(因為無法實際操作發紙條給每位學生，因此將密碼顯示於終端)

3. 發送點名通知給所有已知學生。

4. 在限定時間內定期檢查郵件，處理學生的回覆。

5. 記錄並保存點名結果，標記未回覆或密碼錯誤的學生為缺席。

6. 將點名結果記錄在 `Attendance.xlsx` 文件中。

## 文件結構

- `checkAttendance.py`：主程式文件，包含點名系統的所有邏輯。
- `known_students.xlsx`：包含已知學生電子郵件地址的Excel文件。
- `Attendance.xlsx`：點名結果的記錄文件。

## 注意事項

- 請確保已知學生列表文件 `known_students.xlsx` 路徑正確。
- 請使用有效的 Gmail 帳戶和應用程式密碼。
- 生成應用程式密碼請參考： https://help.url.com.tw/5423.html
- 系統依賴網絡連接來發送和接收電子郵件，請確保網絡連接正常。