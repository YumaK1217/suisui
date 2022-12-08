# -*- coding: utf-8 -*-
"""
Created on Thu Dec  8 08:11:38 2022

@author: ykmnn
"""

""" ライブラリ読込 """
import imapclient
from backports import ssl
from OpenSSL import SSL 
import pyzmail
import pandas as pd
import numpy as np
import datetime
import csv
import pandas as pd
import sys
import itertools
import math
import smtplib
from email.mime.text import MIMEText
 

pd.options.display.max_columns = None
pd.options.display.max_rows = None

atraction_name = ["美女と野獣_魔法の物語", "スプラッシュマウンテン", "ベイマックスのハッピーランド",
                  "ビッグサンダーマウンテン", "ホーンテッドマンション"]


""" (要編集) 引数指定 """
#ログイン情報
my_mail = "suisui20241207@gmail.com"
app_password = "ssdexwefazewwlui"

#メール検索条件
FolderName = "INBOX"
Search_KWD = ["SEEN"] #UNSEEN→未読, SEEN→既読

#メールアドレスの設定
from_email ='suisui20241207@gmail.com'  #最適巡回路を送信するメールアドレス
to_email = 'ykmnnk1217@outlook.jp' #ユーザーのメールアドレス


""" get mail information """
def Get_Mail(my_mail, app_password, FolderName, Search_KWD):
    
    """ ① IMAPサーバー接続 & SSL化 """
    context = ssl.SSLContext(SSL.TLSv1_2_METHOD)
    # context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    imap = imapclient.IMAPClient("imap.gmail.com", ssl=True, ssl_context=context)
    
    """ ② IMAPログイン """
    imap.login(my_mail,app_password)
    
    
    """ ③ メール検索 """
    #対象の受信メールフォルダを指定
    imap.select_folder(FolderName, readonly=True)
    
    #検索キーワードを設定 & 検索キーワードに紐づくメールID検索
    KWD = imap.search(Search_KWD)
    
    #メールID→メール本文取得
    raw_message = imap.fetch(KWD,["BODY[]"])
    
    """  ④ 解析結果保存 """
    #検索結果保存用
    From_list     = []
    Cc_list       = []
    Bcc_list      = []
    Subject_list  = []
    Body_list     = []
    
    #メール解析
    for j in range(len(KWD)):
        #特定メール取得
        message = pyzmail.PyzMessage.factory(raw_message[KWD[j]][b"BODY[]"])

        #宛先取得
        From = message.get_addresses("from")
        From_list.append(From)

        Cc = message.get_addresses("cc")
        Cc_list.append(Cc)

        Bcc = message.get_addresses("bcc")
        Bcc_list.append(Bcc)

        #件名取得
        Subject = message.get_subject()
        Subject_list.append(Subject)

        #本文
        Body = message.text_part.get_payload().decode(message.text_part.charset)
        Body_list.append(Body)


    #Pandas データフレーム
    table = pd.DataFrame({"From":From_list,
                          "Cc":Cc_list,
                          "Bcc":Bcc_list,
                          "件名":Subject_list,
                          "本文":Body_list,
                         })
    
    #IMAPログアウト
    #imap.logout()
    
    return table

#本文から行きたいアトラクションの抽出
def mail_to_algorithm():
    # 関数実行
    table = Get_Mail(my_mail, app_password, FolderName, Search_KWD)
    
    #単語の抽出
    body = table['本文'][0]
    words = body.splitlines()
    goals = []
    goal_id = []
    
    for atraction in atraction_name: #全てのアトラクション
        for word in words: #メールの単語
            if(atraction == word):
              goals.append(word)
    print('行きたい場所の抽出完了')
    
    #場所を番号に変換
    for i in range(len(atraction_name)):
        for goal in goals:
            if(atraction_name[i] == goal):
               goal_id.append(i)
    
    #現在時刻
    dt = datetime.datetime.now()  # ローカルな現在の日付と時刻を取得
    hour = dt.hour #時
    minute = dt.minute #分
    
    #df = pd.DataFrame({'atraction':goal_id,'hour':hour,'minute':minute})
    df = pd.DataFrame({'atraction':goals,'hour':hour,'minute':minute})
    
    return df

#最適な巡回路の計算
def algorithm():    
    go_atraction = [] #ユーザーが行きたいアトラクション
    user_goal = mail_to_algorithm()
    hour = user_goal['hour'][0]
    minute = user_goal['minute'][0]
    df = pd.read_csv('time.csv')
    df_atraction = df.loc[:, ["time"]]
    df_map = pd.read_csv("map.csv", index_col = 0)
    
    """
    print("現在の時刻を入力してね")
    hour = input("何時? ")
    minute = input("何分? ")
    print("乗りたいアトラクションを3つ選んでね")
    print("美女と野獣_魔法の物語  1")
    print("スプラッシュマウンテン　2")
    print("ベイマックスのハッピーランド　3")
    print("ビッグサンダーマウンテン  4")
    print("ホーンテッドマンション　5")
    """
    
    for i in range(len(user_goal['atraction'])):
        go_atraction.append(user_goal['atraction'][i])
    
    for atraction in go_atraction:
        df_atraction.loc[:,atraction] = df[atraction].to_list()
    
    df_atraction_permutation = list(itertools.permutations(go_atraction, len(go_atraction)))
    
    current_time = int(hour) + float(minute) / 60.0
    
    min_wait = 10000
    best_atraction = 0
    
    time_index = df_atraction.index[(df_atraction["time"] - current_time).abs().argsort()][0].tolist()
    
    for i in df_atraction_permutation:
        store_wait = 0.0
        current_time = int(hour) + float(minute) / 60.0
        map_number = 0
        map_point = 0.0
        store_map = go_atraction[0]
        sum = 0.0
    
        for j in i:
            map_number += 1
            current_time += store_wait / 60.0
            time_index = df_atraction.index[(df_atraction["time"] - current_time).abs().argsort()][0].tolist()
            store_wait += df_atraction.loc[time_index][j]
            
            if map_number > 1:
                map_point += math.sqrt((df_map.loc["x", j] - df_map.loc["x", store_map]) ** 2 + (df_map.loc["y", j] - df_map.loc["y", store_map]) ** 2)
            store_map = j
    
        sum = store_wait + map_point
        if(sum < min_wait):
            min_wait = sum
            best_atraction = i
        
        
        # print(store_wait," ",  map_point)
        # print
        # print(i)
        
    text_number = []
    text_wait = []
    text_endtime = []
    current_time = int(hour) + float(minute) / 60.0
    print("現在時刻 " + str(int(current_time)) + ":" + str(int(((current_time - int(current_time)) * 60))))
    text_time = "現在時刻 " + str(int(current_time)) + ":" + str(int(((current_time - int(current_time)) * 60))) + '\n'
    print("回る順番")
    text_label = "回る順番\n"
    number = 1
    for i in best_atraction:
        time_index = df_atraction.index[(df_atraction["time"] - current_time).abs().argsort()][0].tolist()
        print(str(number) + "番目 " + str(i), end = " ")
        text_number.append(str(number) + "番目 " + str(i) + " ")
        print("待ち時間" +  str(int(df_atraction.loc[time_index][i])) + "分" + " ")
        text_wait.append("待ち時間" +  str(int(df_atraction.loc[time_index][i])) + "分" + " ")
        print("アトラクション終了時刻" +str(int(current_time + df_atraction.loc[time_index][i] / 60.0)) + ":" + str(int((current_time + df_atraction.loc[time_index][i] / 60.0 - int(current_time + df_atraction.loc[time_index][i] / 60.0)) * 60)) + "分")
        text_endtime.append("アトラクション終了時刻" +str(int(current_time + df_atraction.loc[time_index][i] / 60.0)) + ":" + str(int((current_time + df_atraction.loc[time_index][i] / 60.0 - int(current_time + df_atraction.loc[time_index][i] / 60.0)) * 60)) + "分")
        current_time += df_atraction.loc[time_index][i] / 60.0
        number += 1
    
    #メールに表示する文章の作成
    text = text_time
    text+=text_label
    for i in range(len(text_number)):        
        text_number[i] += text_wait[i] + '\n' + text_endtime[i] + '\n'
        text+=text_number[i]
        
    return text



def mail_output():
    text =  algorithm()
    cc_mail = ''
    mail_title = '最適巡回路'
    message = text
     
    # PRG3: MIMEオブジェクトでメールを作成
    msg = MIMEText(message, 'plain')
    msg['Subject'] = mail_title
    msg['To'] = to_email
    msg['From'] = from_email
    msg['cc'] = cc_mail
     
    # PRG4: サーバを指定してメールを送信する
    smtp_host = 'smtp.gmail.com'
    smtp_port = 587
    smtp_password = 'ssdexwefazewwlui'
     
    server = smtplib.SMTP(smtp_host, smtp_port)
    server.starttls()
    server.login(from_email, smtp_password)
    server.send_message(msg)
    server.quit()

#メールの送信
mail_output()
