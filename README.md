# 福岡おすすめスポットMAP

FREECAMP学生チームで、福岡のおすすめスポットを地図で共有するWebアプリです。

- `index.html`：アプリ本体（GitHub Pagesで公開）。地図はGoogleマップ。使えないときは自動でOpenStreetMapに切り替わる
- `Code.gs`：Googleスプレッドシートの「拡張機能 → Apps Script」に貼る保存係のプログラム

合言葉とスポットのデータはスプレッドシート側にあり、このリポジトリには入っていません。

GoogleマップのAPIキーは `index.html` の `GOOGLE_KEY` にあります。Google Cloud側で「このサイトのみ」「Maps JavaScript APIのみ」「1日300回まで」の制限をかけて使います。
