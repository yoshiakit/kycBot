# kycBot
Automatically finds if a signed up company has credibility or not and sends message to a slack channel.

## 概要
-法人による自社サービス登録通知がSlackで流れた際、当該会社の法人登記情報についてBotが返信して教えてくれる。

## 前提
- サービスに登録した会社が、社内Slackの特定チャンネルにてリアルタイムで通知(メッセージ送信)される環境にあった
- 筆者にサーバー構築の経験が皆無だったため、まずはサーバーレスなものをつくりたかった
- 自動的に返信を返すようにした

## 使用したツール
- GAS
- Slack API
- Slack Events API
- 国税庁API

## 作成時期
- 2019年1月~2020年2月
