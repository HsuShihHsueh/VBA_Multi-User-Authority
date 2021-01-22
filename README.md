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
在 檢視/工具列/開啟編輯， 可以開啟多行註解
<img src="/picture/img_comment.png" width="375" />
#### 開啟密碼
為了保護撰寫的程式被別人更改，可以將VBA的程式上鎖
在 工具/VBAProject屬性/保護/檢視專案屬性的密碼
<img src="/picture/img_password.png" width="375" />



