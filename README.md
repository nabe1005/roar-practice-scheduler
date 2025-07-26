# roar-practice-scheduler
ROAR練習日程管理用のGASプログラムです
claspを使用して管理しています

## 実行方法

① `clasp login` などでclaspにログインしておきます

ローカルにある `.clasprc-${name}.json` を使用して、`-A` オプションを指定しても大丈夫です

```bash
clasp push -A ~/.clasprc-${hoge}.json	
```

② `clasp push` で変更をプロジェクトに反映します

必要に応じて `-A` オプションを使用してください

## 参考文献

`-A` オプション等については[こちら](https://zenn.dev/ptna/articles/bc49c1d61f6dd7)を参考にしてください

