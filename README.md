# Project Name

VBA_PortFolio_1

## Table of Contents
- [Introduction video](#introduction-video)
- [diagram](#diagram)
- [description](#description)

## Introduction Video

### 1. スクレイピングマクロの実行
- [Webサイト](https://www.net-japan.co.jp/precious_metal/kakaku_past.html)より（株式会社ネットジャパン様のWebサイトのデータをお借りしております。書類面接終了後、直ちに公開を停止いたします。）jsonファイルを取得し、DBに挿入します。実運用では繰り返し実行される処理ではないと考えられますが、技術の実証として作成しています。
### 2. PowerQueryを用いたデータの更新と表示
- PowerQueryを使用してDBに接続し、指定した期間のピボットテーブルとグラフを生成・表示します。
### 3. スクレイピングをバッチファイルから実行
- タスクスケジューラを利用して指定した時間にスクレイピングを実行し、最新のデータのみをDBに挿入します。
### 4. 取得したデータがDBに保存されているかの確認
- スクレイピングで取得したデータは、`market_table`、`nj_buy_table`、および`nj_sell_table`に保存され、それらを利用します。

[![紹介動画](/docs/サムネイル.png)](https://youtu.be/OihXIm_BcHs)

## diagram

![構成図](/docs/portfolio_1.drawio.png)

## description

製作時間: 20 ~ 30 時間

私がこれまで経験してきたVBAを活用したプロジェクトの中で、特に需要が高かったものをベースに本プロジェクトを作成しました。

以下のプロセスを自動で実行します：
1. 特定のWebサイトから必要なデータを取得。
2. 取得したデータを適切に加工。
3. 加工したデータをDBに保存。
4. PowerQueryを用いてDBに接続し、必要なデータを取得。
5. 取得したデータをピボットテーブルやグラフとしてユーザーに表示。

この自動化により、手間のかかる作業をプログラムに委ねることで、データの分析に専念することが可能となります。
