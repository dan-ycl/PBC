# Meeeeeting 時間調查工具


## README 内容列表

- [背景](#背景)
- [安装](#安装)
- [執行步驟](#執行步驟)
- [建置可能發生的錯誤](#建置可能發生的錯誤)
- [雲端部署環境](#雲端部署環境)
- [貢獻者](#貢獻者)



## 背景

本專案的目標是：
透過Tkinter 介面，串接輸入資訊於 Google Sheet 與 Google Calendar API，讓多人參與開會時間投票，並在選擇交集時段後寄送線上會議邀請。


## 安装

1. 這個專案使用了以下套件，請確保在本地(cmd)已經進行安裝。

```sh
pip install google-api-core
pip install google-api-python-client
pip install google-auth
pip install google-auth-httplib2
pip install google-auth-oauthlib
pip install googleapis-common-protos
pip install pygsheets
pip install oauthlib
pip install pandas
pip install tkinter
pip install tkinter.font
pip install csv
pip install os
pip install tkcalendar
pip install datetime 
```

2. 請確保程式檔案儲存階層 (以 windows OS 為例)

> 主程式檔：
```sh
# Windows version
GUI_10.2_windows.py
cal_setup.py
create_event.py
def_read_write_Gsheet.py
formulas.py
token.json
```

> API 金鑰：
```sh
# Windows version
\Google Calendar API\credentials_oauth.json
\Google Sheet\credentials_ServiceAccount.json
```

> 使用命令提示字元呼喚時，請調整預設路徑為本地儲存路徑
  例: cd C:\Users\user\Documents\GitHub\PBC


# 執行步驟

1. 執行 cal_setup.py ，連線雲端產生 token.pickle
2. 依作業系統選擇 GUI_10.2_mac 或 GUI_10.2_windows 執行，啟動 tkinter 視窗
3. 選擇身分「會議發起者」，輸入會議資訊，點擊「建立會議!」
4. 選擇身分「會議發起者」，輸入會議名稱「查詢會議」
5. 當前身分為會議參與者，輸入各列開會時間段「加入時間!」，填寫完畢後「送出!」
6. 點擊「查看會議結果」，選擇下拉列交集會議時段，點擊「建立 Google Calendar!」

* 各步驟間會印出命令提示供使用者確認執行狀態


## 建置可能發生的錯誤

1. 發起行事曆錯誤: ‘Token has been expired or revoked.’
   * 處理方式: 請移除資料夾內應自動產生的 token.pickle，chrome登入 PBC testing 工作帳戶，再次點擊「查看會議結果」

3. tkinter 視窗程式:「沒有回應」
   * 處理方式: 請關閉後重新執行主程式 GUI_10.2


## 雲端部署環境

已啟用 Google Sheet API、Calendar API

1. 試算表:
* Meeting (Host): 記錄已建立之會議資訊清單 https://docs.google.com/spreadsheets/u/1/d/18THYdAFqANCRORZkBp7tDgGgjXGeIWdLFsJK_0ez0qM/edit#gid=0
* Schedule (Attendees): 記錄會議參與者的可行時間段 https://docs.google.com/spreadsheets/d/1h9r2dgUJD5Fze36lkCy0sbpeQAw4V0PJlcc0cVLU4iw/edit#gid=1344869664

2. 公用信箱: PBC testing
* 帳號：PBCtesting2023@gmail.com 
* 密碼：qwer3434


### 貢獻者

感谢參與專案的 Group21 組員 [GitHub] (https://github.com/dan-ycl/PBC)
