# Project Name

VBA_PortFolio_1

## Table of Contents
- [diagram](#diagram)
- [description](#description)
- [Introduction video](#introduction-video)

## Introduction video

### 1. スクレイピングマクロの実行
-> [Webサイト](https://www.net-japan.co.jp/precious_metal/kakaku_past.html)(株式会社ネットジャパン様のWebサイトのデータをお借りしました。書類面接の終了と同時に公開を停止します。)からjsonファイルを取得してDBに挿入します。実運用では繰り返し使われる処理ではないと想定されますが、技術の証明として作成しております。

### 2. PowerQueryを用いたデータの更新と表示
-> PowerQueryを用いてDBに接続し、指定した期間のピボットテーブルとグラフを表示する

### 3. スクレイピングをバッチファイルから実行
-> 指定の時間にタスクスケジューラから実行され、最新データのみをDBに挿入する

### 4. 取得したデータがDBに保存されているかの確認
-> スクレイピングによって取得したデータはmarket_table、nj_buy_table、nj_sell_tableに保存され利用する

[![紹介動画](docs\サムネイル.png)](https://youtu.be/OihXIm_BcHs)

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

