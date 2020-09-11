import requests
from apiclient.discovery import build
import config
import urllib.request
import urllib.error

import re
import os
import sys
import random

import wmi
import subprocess
import time


# 動画情報取得
class YoutubeVideoGet():
    # Youtube Data API 認証
    def __init__(self, out_path):
        yt_api_key = config.yt_api_key
        self.yt_api = build('youtube', 'v3', developerKey=yt_api_key)
        self.out_path = out_path

    # 動画情報取得
    def getVideo(self, search_list):
        # 動画取得
        search_response = self.yt_api.search().list(
            part='snippet',
            q=random.choice(search_list),
            type='video'
        ).execute()

        # ランダムピックアップ
        video = search_response['items'][random.randint(0, len(search_response['items'])-1)]

        # 動画の長さ取得
        video_content = self.yt_api.videos().list(
            part = 'contentDetails', 
            id = video['id']['videoId']
        ).execute()
        duration = video_content['items'][0]['contentDetails']['duration'] # 動画の長さ
        
        # ISO-8601 Duration -> Seconds
        video_length = self.calcTime_duration2sec(duration)

        # サムネイルダウンロード
        thumb_url = video['snippet']['thumbnails']['high']['url'] # サムネイル
        self.download_thumb(thumb_url, self.out_path)

        return video, video_length

    # サムネイルダウンロード
    def download_thumb(self, thumb_url, out_path):
        try:
            # 画像URLを開き、読み込む
            with urllib.request.urlopen(thumb_url) as web_file:
                thumb_data = web_file.read()

                # 書き込み
                with open(out_path, mode='wb') as thumb:
                    thumb.write(thumb_data)

        except urllib.error.URLError as e:
            print(e)

    # ISO-8601 Duration -> Seconds
    def calcTime_duration2sec(self, duration):
        time_list = re.findall('[0-9]+[D|H|M|S]+', duration) # 時間情報のみ抽出

        video_length = 0
        # 計算
        for i in range(len(time_list)):
            # Days
            if 'D' in time_list[i]:
                d = time_list[i].replace('D', '')
                video_length += int(d) * 86400
                continue
            # Hours
            if 'H' in time_list[i]:
                h = time_list[i].replace('H', '')
                video_length += int(h) * 3600
                continue
            # Minutes
            if 'M' in time_list[i]:
                m = time_list[i].replace('M', '')
                video_length += int(m) * 60
                continue
            # Seconds
            if 'S' in time_list[i]:
                s = time_list[i].replace('S', '')
                video_length += int(s)
                continue

        return video_length

# Notify送信
class LineNotifySend():
    # Line Notify API セットアップ
    def __init__(self, out_path):
        line_token = config.line_token
        self.url = "https://notify-api.line.me/api/notify"
        self.line_headers = {'Authorization': 'Bearer ' + line_token}
        self.out_path = out_path

    # メッセージ送信
    def sendMessage(self, video, length):
        video_title = video['snippet']['title'] # タイトル
        video_url = 'https://www.youtube.com/watch?v=' + video['id']['videoId'] # 動画URL
        
        # サムネイル読み込み
        with open(self.out_path, 'rb') as thumb:
            thumb_data = thumb.read()

        # 時間計算
        time = self.calcTime_sec2time(length)

        # メッセージ内容
        message = "\n" + video_title + "\n[" + time + "]\n" + video_url
        payload = {'message': message}
        # サムネイル
        thumb = {'imageFile' : thumb_data}
        
        # Line送信
        requests.post(self.url, headers=self.line_headers, params=payload, files=thumb)

        # サムネイル削除
        os.remove(self.out_path)

    # ML完了時メッセージ送信
    def sendFinishMessage(self):
        message = "機械学習が完了しました！！作業に戻ってください。"
        payload = {'message' : message}
        requests.post(self.url, headers=self.line_headers, params=payload)

    # Seconds -> 時刻
    def calcTime_sec2time(self, length):
        h, m, s = 0, 0, 0
        h = int(length / (60 * 60))
        length = length % (60 * 60)
        m = int(length / 60)
        s = length % 60
        
        time = '{:02}:{:02}:{:02}'.format(h, m, s)
        return time

def main():

    # 検索ワードリスト
    search_list = ['機械学習', '猫', 'ディープラーニング', 'アンパンマン']

    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),'YTthumb_temp.jpg')
    
    # Youtube動画取得
    YTget = YoutubeVideoGet(out_path)
    
    # LINE Notify送信
    LNsend = LineNotifySend(out_path)

    c = wmi.WMI()
    cmd = 'nvidia-smi'
    cre_process_watcher = c.Win32_Process.watch_for("creation") 

    while True:
        delay = 0

        new_cre_process = cre_process_watcher()
        time.sleep(5)

        # 新規でPythonが起動されたか
        if new_cre_process.Caption == "python.exe":
            try:
                # 'nvidia-smi'コマンド実行
                res_b = subprocess.run(cmd, stdout=subprocess.PIPE)
                # 実行結果デコード
                res = res_b.stdout.decode(encoding='utf-8')

                # 'Anacon3\python.exe'が含まれている -> 機械学習実行中
                if res.find("Anaconda3\\python.exe") != -1:
                    print("Running Machine learning.")

                    while True:
                        time.sleep(5)

                        # 'nvidia-smi'コマンド実行
                        res_b = subprocess.run(cmd, stdout=subprocess.PIPE)
                        # 実行結果デコード
                        res = res_b.stdout.decode(encoding='utf-8')
                        # 'Anacon3\python.exe'が含まれていない -> 機械学習完了
                        if res.find("Anaconda3\\python.exe") == -1:
                            print("Machine learning has been completed.")
                            LNsend.sendFinishMessage() # 完了を通知
                            break
                        else: # 'Anacon3\python.exe'が含まれている -> 機械学習実行中
                            if delay == 0:
                                video, length = YTget.getVideo(search_list) # 動画取得
                                LNsend.sendMessage(video, length) # 動画URL送信
                            delay += 1

                            # 動画が見終わる頃合いになったら -> 動画URL送信
                            if delay > int(int(length*0.95)/5):
                                delay = 0
                else:
                    continue
            except subprocess.CalledProcessError as e:
                print(e)

if __name__ == '__main__':
    main()
