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

私が今まで経験してきたVBAを用いたプロジェクトの中で、需要の大きかったプロジェクトを元に作成いたしました。

1. 特定のWebサイトから必要なデータを取得
2. 1.にて取得したデータを加工
3. 2.にて加工したデータをDBに保存
4. DBにPowerQueryを用いて接続し必要データを取得
5. 5.にて取得したデータをピボットテーブル・グラフにしてユーザーに表示

上記プロセスを指定した時間に自動で実行します。

これにより煩雑な作業をプログラムに任せ、データの分析作業に集中することができます。

