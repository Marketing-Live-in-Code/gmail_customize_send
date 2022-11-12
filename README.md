# Python實現Gmail客製化信件大量發送 — 【基礎篇】

在mail中cc別人，是一個非常方便的功能，但有時候每個信件必須客製化，譬如在信件開頭必須打上「親愛的XXX小姐您好」，姓名的部分就必須要手動替換，非常麻煩，更別說會員還有上百，甚至上千個。其實用Python程式可以完全自動化這個過程，我們將以Gmail為範例帶領您走向自動化寄信的航道。

本文章會以以下流程來為各位命苦的小編、行銷人員、秘書解惑：
1. Gmail設定與申請
2. Python寄信基礎教學
3. 客製化寄件案例實戰，這會在「Python實現Gmail客製化信件大量發送 — 【案例實戰篇】」中詳細說明。

## Gmail設定與申請
本文將使用Gmail為範例，因此我們必須先將Gmail帳號開通，許可Python可以使用自己的帳號發送信件。

### 進入Google帳號管理
請到Chrome中開啟分頁，點選您的帳號圖片，並點選「管理你的Google帳號」，即可進入帳號管理畫面。
![進入Google帳號管理](https://i.imgur.com/E20lrrY.png)

### 進入安全性管理
在左邊的選單中，選擇「安全性」，並將兩步驟驗證打開，打開時會要求一些身分驗證。
![進入安全性管理](https://i.imgur.com/Q4zlSP9.png)

### 設定專案名稱
完成後還必須設定一個專案密碼，您的Gmail就可以透過Python發送信件了。應用程式直接選取「其他」，然後輸入自己自訂的專案名稱。
![設定專案名稱](https://i.imgur.com/vEnWEMY.png)

### 取得專案密碼
完成後您就會得到一串密碼，為了保護作者資訊安全，這裡將密碼置換成範例「abcdefghijkalmno」。請將這組密碼保護好，否則有心人士就可以以您的名義濫發信件，進行非法用途。
![取得專案密碼](https://i.imgur.com/ETQryH6.png)

## Python寄信基礎教學
接下來的Python寄信基礎教學，會切分成「附件」與「文字」兩個區塊進行教學。「附件」是平時我們在信件中可以看到的附件檔案，通常都會是圖片、PDF、Word、Excel格式。另外「文字」則是分享如何客製化信中的文字樣式，譬如人名下方需要底線，重點需要粗體等等。

### 寄送檔案
![完整程式碼](https://i.imgur.com/S6PzDrz.png)
在程式碼中可以對照後方的註解，您便會了解該行程式碼在作什麼。首先先創建一個MIMEMultipart物件，您可以把他想成準備好一張信紙，當然一開始整張紙都是白的。
```python
content = MIMEMultipart() #建立MIMEMultipart物件
```
設定信件的標題，身為小編的您，聳動的標題就下在這裡。
```python
ccontent["subject"] = "測試寄信"  #郵件標題
```
您必須將程式碼16行改成前個步驟，開啟Python可以控制的Gmail帳號。程式碼17行則改成要接收信的Email帳號。
```python
content["from"] = "寄件者的信箱"  #寄件者
content["to"] = "收件者的信箱" #收件者
```
![請不要輸入作者我的信箱](https://i.imgur.com/Nl1KeL2.png)
放置信件的內文，內文的客製化會在後面的文章詳細講解。
```python
content.attach(MIMEText("Ivan的測試寄信，寄信處女作品～～"))  #郵件內容
```
增加一個圖片附件，只需要將圖片與執行的程式碼放在一起，便可以將圖片包入信件一並寄送。
```python
content.attach(MIMEImage(Path("夕陽.jpg").read_bytes()))  # 郵件圖片內容
```
![這張照片是在台南漁光島所拍下的美景](https://i.imgur.com/VvIc8V1.png)
利用MIMEApplication這個方法，能包附非圖片的檔案。在Email中最常使用的PDF就可以利用這個方法進行傳送。
```python
#寄送PDF檔案
fileName = 'test.pdf'
pdfload = MIMEApplication(open(fileName,'rb').read()) 
pdfload.add_header('Content-Disposition', 
                   'attachment', 
                   filename=fileName) 
content.attach(pdfload) 
```
當然常見的word檔案也是可以進行傳送的，只需要設定好檔案名稱，其他程式碼內容都不需要修改。
```python
#寄送Word檔案
fileName = 'test.docx'
pdfload = MIMEApplication(open(fileName,'rb').read()) 
pdfload.add_header('Content-Disposition', 
                   'attachment', 
                   filename=fileName) 
content.attach(pdfload) 
```
為了示範各種格式檔案，這裡也將下一篇文章才會用到的「顧客訂單.csv」檔案，包進信件進行傳送了。
```python
#寄送csv檔案
fileName = '顧客訂單.csv'
pdfload = MIMEApplication(open(fileName,'rb').read()) 
pdfload.add_header('Content-Disposition', 
                   'attachment', 
                   filename=fileName) 
content.attach(pdfload) 
```
以上的動作都是在「寫信」，接下來要將寫好的信件，寄送到另一個自己的帳號進行測試（這個帳號就不需要開啟Google帳號安全性）。

這裡必須將程式碼49行的「寄件者信箱」，換成自己的寄件信箱，與程式碼16行的信箱相同。而密碼並不是登入該Gmail的密碼，是我們在第1–4步驟所取得的密碼。
```python
with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
    try:
        smtp.ehlo()  # 驗證SMTP伺服器
        smtp.starttls()  # 建立加密傳輸
        smtp.login("寄件者的信箱", "寄件者的寄件密碼")  # 登入寄件者gmail
        smtp.send_message(content)  # 寄送郵件
        print("成功傳送")
    except Exception as e:
        print("Error message: ", e)
```
![請不要輸入作者的信箱及密碼，密碼也是假的](https://i.imgur.com/YwCWJDG.png)
執行後在接收的信箱中，就會得到剛剛所寫的信了。到這裡若您對程式碼還是有些不了解，可以將程式碼與信件傳送的結果作對比，也可以了解那些程式碼控制信件的那些區域。
![](https://i.imgur.com/xJjnVH4.png)
### 文字樣式客製化
![](https://i.imgur.com/LTW8Wlx.png)
在程式碼的13、14、30行，也都需要如上方修改成自己的信箱與密碼，因此這裡不再重複。客製化的文字主要都在MIMEText這個方法內，而裡面的內容則是用HTML（網頁）的程式碼進行編寫。
```python
content.attach(
                MIMEText("""
                        親愛的 <u>Ivan</u>您好：<br><br>
                        您寫這篇文章真是辛苦了，<b>不如您自己追蹤自己吧！</b> <br>
                        趕快手刀點擊<a href="https://medium.com/@ivanyang0606">Ivan的Medium文章</a>。
                        """
                        , "html"))  #郵件內容
每個標籤都有他的開始與結束，可以發現標籤的結束，都會有一個「/」來做為結尾，中間包附您想要呈現的文字。
```
![](https://i.imgur.com/PnSxELz.png)
HTML的標籤其實有非常多種類，這裡列出比較常用的幾種供您作參考，這些標籤也在範例中都有實作。
![](https://i.imgur.com/uoyOONK.png)
可以看到呈現的結果，我們的期待，甚至超連結（a標籤）的按鈕進行點選，一樣可以連結到作者的文章部落格喔！
![](https://i.imgur.com/9C9KN25.png)
如何每封信件進行客製化呢？

在下一篇文章[「Python實現Gmail客製化信件大量發送 — 【案例實戰篇】」]()中，我們將以實戰的方式，將本篇文章所學的技能，成功應用到Email廣宣上，從此行銷人員只需要一鍵即可完成往日兩三個小時的工作了。
