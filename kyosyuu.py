from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import datetime
import time
import smtplib
from email.mime.text import MIMEText
from email.utils import formatdate

from sqlalchemy import false, true
"""
教習所の予約状況を監視し、空きがあればGmailにメールを送って知らせるプログラム。
既に予約を入れてある日は避ける（一日一時限）。
ngday_listに手動で入力すれば任意の日を避けることもできる。ただし、日付が変わるとズレるので毎日入力し直す必要がある。
day_rangeを変更すれば今日から何日以内の予約を狙うか決められる。
"""

#メールの生成
def create_message(from_addr, to_addr, subject, body):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Date'] = formatdate()
    return msg

#スマホにメールを送って空きがあることを知らせる
def send_gmail(from_addr, to_addr, body):
    smtpobj = smtplib.SMTP('ホスト', 'ポート番号')
    smtpobj.ehlo()
    smtpobj.starttls()
    smtpobj.ehlo()
    smtpobj.login("メアド", "パスワード")
    smtpobj.sendmail(from_addr, to_addr, body.as_string())
    smtpobj.close()

#予約を取りたくない日。今日が0、明日が1、明後日が2と続いていく
ngday_list = [0]

#今日から何日以内の予約を取るか
day_range = 6

#chromeとchromedriverのバージョンを合わせる
driver = webdriver.Chrome(ChromeDriverManager().install())
#接続
driver.get("リンク")
time.sleep(2)
#ここを押してログインというボタンオブジェクトをクリック
driver.find_element(by=By.ID, value="lnkToLogin").click()
time.sleep(2)
#サインイン画面はframeが複数ある構成だったので、入力を受けるframeに移動
iframe = driver.find_element(by=By.ID, value="frameMenu")
driver.switch_to.frame(iframe)
time.sleep(2)
#ログイン
driver.find_element(by=By.ID, value="txtKyoushuuseiNO").send_keys("ユーザ番号")
driver.find_element(by=By.ID, value="txtPassword").send_keys("パスワード")
driver.find_element(by=By.ID, value="btnAuthentication").click()
time.sleep(2)
#持ち予約画面に移動
driver.find_element(By.ID, "btnMenu_YoyakuItiran").click()
time.sleep(2)
#持ち予約の親要素idがlst_lc1, lst_lc2, ...であることを利用し、xpathで全ての持ち予約をリストで取得
bookings = driver.find_elements(By.XPATH, "//*[@id=\"lst_lc\"]/div/div[@class=\"blocks\"]/div[1]/span")

#一つ以上の持ち予約があれば
if bookings:
    #今日の日付などの情報を取得
    now = datetime.datetime.now()
    #持ち予約を格納するリスト
    bookings_list = []
    for i in bookings:
        #要素をstr型で取得
        temp = i.get_attribute("textContent")
        #取得してきた持ち予約は「2022/3/1（担当者：hogehoge）」のようになっているので、2022/3/1の部分だけを取り出す
        target = '('
        idx = temp.find(target)
        temp = temp[:idx]
        #[年, 月, 日]のリストにする
        temp = temp.split('/')
        #int型に変換
        temp = [int(i) for i in temp]
        #bookings_listに追加。例：bookings_list=[[2022, 3, 1], [2022, 3, 3], ...]
        bookings_list.append(temp)
    for i in bookings_list:
        #持ち予約が現在から何日後かを計算し、ngday_listに追加
        temp = datetime.datetime(i[0], i[1], i[2], now.hour, now.minute, now.second, now.microsecond)
        ngday_list.append((temp-now).days)

#ホーム画面に戻る
driver.find_element(By.ID, "layout_header_back").click()
time.sleep(2)
#空き予約状況の画面に移動
driver.find_element(by=By.ID, value="btnMenu_Kyoushuuyoyaku").click()
time.sleep(2)

for i in range(1000):
    #空きなら「〇」、そうでなければ「×」と表示されているので、その要素をクラス名で指定してすべて取得
    for day in driver.find_elements(by=By.CLASS_NAME, value="badge"):
        if day.text != "×":
            #〇×要素の親要素に割り振られたidが、明日の予約なら「id=hoge1」、明後日なら「id=hoge2」となっていて都合がいいので、id名の番号だけをxpathで取得してくる
            oya_id = day.find_element(By.XPATH, "(../../../preceding-sibling::input)[last()]")
            oya_id = oya_id.get_attribute("id")[-1]
            #ng_flagはその予約がday_rangeを超えていたり、ngday_listに該当したりするとTrue。
            ng_flag = False
            if int(oya_id) > day_range:
                ng_flag = True
            for ng_day in ngday_list:
                if ng_day == int(oya_id):
                    ng_flag = True
            #ng_flagがTrueならこの空き予約を無視
            if ng_flag:
                continue
            #その日の何時が空いているか調べる。時限は一日を13分割しており、1~13と数字で表示されているため、この数字の要素をxpathで取得したい。
            #お目当ての数字要素は日付を表示する要素の中にあり、その日付を表示する要素のid名がlstDetail_0_lc、lstDetail_1_lc、...と連番になっているので、「id='lstDetail_' + oya_id + '_lc'」とすれば指定できる
            #後は/div[1]/div[1]/...と子要素を辿っていくと数字の要素が取得できる
            zigen_id = 'lstDetail_' + oya_id + '_lc'
            zigen_path = "//div[@id=\""+zigen_id+"\"]/div[1]/div[1]/div[1]/div[1]/span[1]"
            zigen = driver.find_element(By.XPATH, zigen_path)
            #早朝や夜間を避ける
            if (int(zigen.get_attribute("textContent")) < 11) and (int(zigen.get_attribute("textContent")) > 2):
                #driverを閉じる
                driver.quit()
                #メールをスマホに送る
                msg = create_message("メアド", "メアド", oya_id + "番空きあり", "")
                send_gmail("メアド", "メアド", msg)
                #プログラムを終了
                exit()

    #画面更新のために別の画面を開いてから戻ってくる
    time.sleep(60)
    driver.find_element(by=By.CLASS_NAME, value="next").click()
    time.sleep(2)
    driver.find_element(by=By.CLASS_NAME, value="previous").click()

