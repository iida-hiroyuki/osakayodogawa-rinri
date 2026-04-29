# osakayodogawa-rinri

大阪淀川倫理法人会の会員リストを、Googleフォーム回答（スプレッドシート）から反映して公開するサイトです。

## 公開URL

- `https://iida-hiroyuki.github.io/osakayodogawa-rinri/`

## ファイル構成

- `index.html`: 会員一覧UI（3列カード + 詳細モーダル）
- `members.json`: 会員データ本体（GASが上書き更新）
- `gas/Code.gs`: スプレッドシートに貼り付けるGoogle Apps Script

## ローカル編集

静的サイトのため、`index.html` と `members.json` の編集結果はそのまま GitHub Pages に反映されます。

## GASセットアップ手順（初回のみ）

1. GitHubでPersonal Access Tokenを作成（Fine-grained推奨）
   - Repository access: `osakayodogawa-rinri` のみ
   - Permissions: `Contents` = `Read and write`
2. スプレッドシートで `拡張機能` → `Apps Script` を開く
3. `gas/Code.gs` の内容を Apps Script に貼り付けて保存
4. `プロジェクトの設定` → `スクリプトプロパティ` に以下を追加
   - `GITHUB_TOKEN`: 上記PAT
   - `GITHUB_REPO`: `iida-hiroyuki/osakayodogawa-rinri`
   - `GITHUB_BRANCH`: `main`（任意、未設定時はmain）
5. スプレッドシートを再読み込み
6. メニュー `会員サイト` → `公開・更新（members.json）` を実行

## スプレッドシート列（1行目ヘッダー）

- 単会役職
- 会社名
- 氏名
- 業務内容
- 倫理に入ったきっかけ
- 顔写真

## 運用フロー

1. Googleフォーム回答がスプレッドシートに追加される
2. 必要時に `会員サイト` → `公開・更新` を押す
3. 数秒で `members.json` が更新され、公開サイトに反映

## 補足

- 顔写真URLは、Googleドライブの `/file/d/{id}/view` 形式でも `uc?id={id}` に変換して表示します。
- `=IMAGE("...")` 形式のセル値もURL抽出して対応します。
