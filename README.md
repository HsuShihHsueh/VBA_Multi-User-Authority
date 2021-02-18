# 多使用者權限控制

當需要多人共同編輯時，不免俗的會看到別人檔案，所以本專案希望建立一EXCEL檔案，利用一權限表決定該用戶能擁有查閱那些工作部的權限。本專案是採用各幹部評定考績來當作範本<br><br>
以下範例程式的密碼皆是 0 <br>
#
👇 youtube link<br>
var image = "https://i9.ytimg.com/vi/giUrLR0bFrg/maxresdefault.jpg?time=1613615100000&sqp=CPyft4EG&rs=AOn4CLDCiUrWZAh5dSdnRy-ul-jOxtkNEA"
var link = "http://www.youtube.com/watch?v=giUrLR0bFrg"
<a href="http://www.youtube.com/watch?v=giUrLR0bFrg"><img src="https://i9.ytimg.com/vi/giUrLR0bFrg/maxresdefault.jpg?time=1613615100000&sqp=CPyft4EG&rs=AOn4CLDCiUrWZAh5dSdnRy-ul-jOxtkNEA" border="0"></a><br>
<a href=link><img src=image border="0"></a><br>
[![](https://i9.ytimg.com/vi/giUrLR0bFrg/maxresdefault.jpg?time=1613615100000&sqp=CPyft4EG&rs=AOn4CLDCiUrWZAh5dSdnRy-ul-jOxtkNEA)](http://www.youtube.com/watch?v=giUrLR0bFrg "")

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

#### 權限表
可以決定個幹部可以看到的權限<br>
<img src="/picture/img_access.png" width="375" /><br>

#### 幹部讀取員工名單
為了日後增減員工名單，所以只讓最高權限的編輯名稱，其他幹部去讀取名單<br>
```
=IF(廠長!A4="","",廠長!A4)
```
#### 鎖定編輯範圍
為了防止一些檔案方程式被不經意刪掉，以下是如何鎖定編輯範圍<br>
點擊 校閱/允許使用者編輯範圍<br>
在 "允許使用者編輯範圍" 視窗中，點擊修改<br>
在 參考儲存格中，可以編輯範圍或是直接拖曳想要範圍<br>
這邊我們只希望幹部能編輯自己的密碼及員工分數<br>
這邊也可以讓下屬不能打上司的成績(不選取範圍即可)<br>
<img src="/picture/img_range.png" width="600" /><br>
點擊校閱/保護工作表，這邊也可以設定密碼<br>
如果編輯未選取範圍，會跳出以下警告<br>
<img src="/picture/img_protect.png" width="600" /><br>

詳細教程 <a href="https://bp6ru8.pixnet.net/blog/post/35656420-42.-excel:-%E5%87%BA%E5%B7%AE%E8%B2%BB%E7%94%A8%E8%A1%A8%EF%BC%8C%E5%8F%AA%E8%83%BD%E5%A1%AB%E5%AF%AB%E7%89%B9%E5%AE%9A%E5%84%B2%E5%AD%98%E6%A0%BC-%E5%85%81" target="_blank">Excel: 出差費用表，只能填寫特定儲存格/允許使用者編輯範圍</a><br>

#### 讀取工作表名稱
這邊要注意一定要打$A$1或者別的範圍，才會鎖定在自身的工作表，如果沒打，在編輯別的工作表時，名稱會跑掉<br>
前面加" "&是因為11月及12月跟別的月份長度不一樣，所以一律讀取兩位數<br>
【月份】<br>
```
=" "&MID(CELL("filename",$A$1),FIND("]",CELL("filename",$A$1))+1,LEN(CELL("filename",$A$1))-FIND("]",CELL("filename",$A$1)))
```
【月份轉換】(在J1)<br>
```
=CHOOSE(MID($A$1,LEN($A$1)-2,2),"B","C","D","E","F","G","H","I","J","K","L","M")
```
【從幹部那邊抓成績】<br>
```
=IF(INDIRECT(B$1&"!"&$J$1&ROW()+1)="","",INDIRECT(B$1&"!"&$J$1&ROW()+1))
```
總分：為了怕有幹部未評分導致某些成績失真，所以先計算出有評分的幹部滿分，然後將加完權重的成績除以滿分，即可得分數<br>
```
=IF(OR($B5="",$B5="x"),0,$B$2)+IF(OR($C5="",$C5="x"),0,$C$2)+IF(OR($D5="",$D5="x"),0,$D$2)+IF(OR($E5="",$E5="x"),0,$E$2)+IF(OR($F5="",$F5="x"),0,$F$2)+IF(OR($G5="",$G5="x"),0,$G$2)
```
<img src="/picture/img_score.png" width="600" /><br>
不會因為沒被評分而被拉低平均<br>

