# 多使用者權限控制

當需要多人共同編輯時，不免俗的會看到別人檔案，所以本專案希望建立一EXCEL檔案，利用一權限表決定該用戶能擁有查閱那些工作部的權限。本專案是採用各幹部評定考績來當作範本<br>



## 啟用巨集

在 EXCEL/檔案/選項 裡，<br>
在"自訂功能區"，啟用"開發工具" <br>
詳細教程：<a href="https://officeguide.cc/excel-show-developer-tab-tutorial/" target="_blank">Excel 啟用「開發人員」工具教學</a><br><br>
點擊開發人員/Visual Basic, 或者按快捷鍵"alt+F11"開啟VBA

記得另存成.xlsm檔，才能將巨集程式保存起來<br>
存完後記得啟用內容<br>
<img src="/picture/img_open_marco.png" width="375" />

## VBA(設定環境)

#### 多行註解
在 檢視/工具列/開啟編輯， 可以開啟多行註解<br>
<img src="/picture/img_comment.png" width="375" />
#### 開啟密碼
為了保護撰寫的程式被別人更改，可以將VBA的程式上鎖<br>
在 工具/VBAProject屬性/保護/檢視專案屬性的密碼<br>
<img src="/picture/img_password.png" width="600" />
#### 文字風格
在 工具/選項/選寫風格<br>
可更改文字字體與大小<br>
<img src="/picture/img_font.png" width="375" />

## VBA(程式)

#### Thisworkbook
開啟Workbook<br>
<img src="/picture/img_thisworkbook.png" width="375" /><br>
將"workbook.vba"的程式複製到右方裡<br>

#### UserForm
在 插入/自訂表單<br>
<img src="/picture/img_userform1.png" width="375" /><br>
利用工具列可將表單布置成下方形式<br>
<img src="/picture/img_userform2.png" width="375" /><br>
點擊按鈕或文字進入編寫程式<br>
一樣將"userform.vba"的程式複製到右方裡<br>

## EXCEL 工作簿

#### 幹部讀取員工名單
為了日後增減員工名單，所以只讓最高權限的編輯名稱，其他幹部去讀取名單<br>
'''
=IF(廠長!A4="","",廠長!A4)
'''
#### 鎖定編輯範圍
為了防止一些檔案方程式被不經意刪掉，以下是如何鎖定編輯範圍
點擊 校閱/允許使用者編輯範圍

詳細教程 <a href="https://bp6ru8.pixnet.net/blog/post/35656420-42.-excel:-%E5%87%BA%E5%B7%AE%E8%B2%BB%E7%94%A8%E8%A1%A8%EF%BC%8C%E5%8F%AA%E8%83%BD%E5%A1%AB%E5%AF%AB%E7%89%B9%E5%AE%9A%E5%84%B2%E5%AD%98%E6%A0%BC-%E5%85%81" target="_blank">Excel: 出差費用表，只能填寫特定儲存格/允許使用者編輯範圍</a><br><br>
