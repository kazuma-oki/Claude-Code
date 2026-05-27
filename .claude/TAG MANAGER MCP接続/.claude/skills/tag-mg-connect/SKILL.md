---
name: tag-mg-connect
description: Google Tag Manager MCPサーバーをClaude Code に接続するセットアップ手順。「GTMに接続したい」「Tag ManagerのMCPを設定したい」「gtm-mcpが動かない」などの場合に使用する。
---

## Claude への指示

このスキルが起動したら、**1ステップずつ対話形式**でユーザーを案内する。  
次のステップへ進む前に必ず確認コマンドを実行し、成功を確認する。  
コマンドはユーザーのターミナルで実行してもらう（Claude自身はツールで実行しない）。

---

## 最初にユーザーへ確認する

スキル起動時に以下を順番に質問する:

1. **OSを確認する**  
   「Windows / Mac / Linux のどれを使っていますか？」

2. **Node.js がインストール済みか確認する**  
   「ターミナルで `node -v` を実行してください。バージョンが表示されますか？」  
   - 表示される → そのまま進む  
   - 表示されない → [nodejs.org](https://nodejs.org/) からインストールしてもらう

3. **サービスアカウントキーの有無を確認する**  
   「GCPのサービスアカウントキー（.json ファイル）はこの端末にありますか？」  
   - ある → パスを教えてもらう  
   - ない → 「別の端末からコピーできますか？それとも新規作成が必要ですか？」

4. **GCPプロジェクトIDを確認する**  
   「GCPのプロジェクトIDを教えてください」  
   （参考: このアカウントは `iron-axon-162706`）

5. **GTMの初期設定が済んでいるかを確認する**  
   「GTMアカウントへのサービスアカウントの追加はすでに設定済みですか？」  
   - 済 → STEP 1（キー準備）から開始  
   - 未 → STEP 0（GCP/GTM設定）から開始

確認が取れたら「では順番に進めましょう」と伝え、STEP を開始する。

---

## STEP 0: GCP / GTM の初期設定（初回のみ・既に設定済みならスキップ）

以下をユーザーに案内する（ブラウザ操作のため、完了したら次へ進むよう伝える）:

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセスし、対象プロジェクトを開く

2. **Google Tag Manager API** を有効化する:  
   「APIとサービス」→「ライブラリ」→ `Tag Manager API` を検索して「有効にする」

3. サービスアカウントを作成しキーをダウンロードする:
   - 「IAMと管理」→「サービスアカウント」→「作成」
   - 任意の名前（例: `gtm-mcp-server`）で作成
   - 作成後、「キー」タブ → 「鍵を追加」→「新しい鍵を作成」→ JSON 形式でダウンロード

4. GTM でアクセスを許可する:
   - [Google Tag Manager](https://tagmanager.google.com/) を開く
   - 対象アカウントの右上「⋮」→「ユーザー管理」
   - サービスアカウントのメールアドレスを **管理者** 権限で追加  
   （アカウントレベルで追加すること。コンテナレベルのみでは不十分な場合がある）

完了したら確認: 「JSONキーのダウンロードと GTM のアクセス設定は完了しましたか？」

---

## STEP 1: サービスアカウントキーをこの端末に配置する

### キーが別端末にある場合

ユーザーに伝える:  
「キーファイル（.json）を別端末からこの端末にコピーしてください。  
パスに**日本語・スペースが含まれない場所**に置いてください」

推奨配置場所:
```
C:\Users\<ユーザー名>\.claude\MY PROJECT\MCP Setting\TAG MANEGER MCP接続\.claude\tag-mg-key\<キーファイル名>.json
```

### 配置後の確認（ファイルが読めるかチェック）

ユーザーがキーのパスを教えたら、以下のコマンドを実行してもらう:

**Windows:**
```powershell
Get-Content "ここにキーのパス" | ConvertFrom-Json | Select-Object project_id, client_email
```

**Mac/Linux:**
```bash
python3 -c "import json; d=json.load(open('ここにキーのパス')); print(d['project_id'], d['client_email'])"
```

`project_id` と `client_email` が表示されれば OK。

---

## STEP 2: MCPパッケージをインストールする

```bash
npm install -g @weppa-cloud/mcp-google-tag-manager
```

インストール後、JSファイルのパスを確認してもらう:

**Windows:**
```powershell
(Get-Command node -ErrorAction SilentlyContinue).Source
```
```powershell
Test-Path "$env:APPDATA\npm\node_modules\@weppa-cloud\mcp-google-tag-manager\bin\mcp-google-tag-manager.js"
```

**Mac/Linux:**
```bash
which node
ls $(npm root -g)/@weppa-cloud/mcp-google-tag-manager/bin/mcp-google-tag-manager.js
```

→ `node.exe` のパスを `[NODE_PATH]`、JSファイルのパスを `[JS_PATH]` として控えてもらう。

---

## STEP 3: claude コマンドのパスを確認する

**Windows:**
```powershell
(Get-Command claude -ErrorAction SilentlyContinue).Source
```
表示されない場合:
```powershell
Get-ChildItem "$env:LOCALAPPDATA\Packages\Claude_*" -Recurse -Filter "claude.exe" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName
```

**Mac/Linux:**
```bash
which claude
```

→ このパスを `[CLAUDE_PATH]` として控えてもらう。

---

## STEP 4: MCP サーバーを登録する

**重要:** MCP の設定は必ず `~/.claude.json`（ホームディレクトリ直下）に書き込まれる。  
`~/.claude/settings.json` に書いても MCP ツールとして認識されない。  
`claude mcp add` コマンドを使えば自動で正しいファイルに書き込まれる。

**Windows:**
```powershell
& "[CLAUDE_PATH]" mcp add --scope user google-tag-manager "[NODE_PATH]" `
  --args "[JS_PATH]" `
  -e GOOGLE_APPLICATION_CREDENTIALS="[キーのパス]"
```

**Mac/Linux:**
```bash
"[CLAUDE_PATH]" mcp add --scope user google-tag-manager "[NODE_PATH]" \
  --args "[JS_PATH]" \
  -e GOOGLE_APPLICATION_CREDENTIALS="[キーのパス]"
```

実行後、エラーメッセージが出た場合は内容を教えてもらい対処する。

---

## STEP 5: 登録を確認する

```powershell
# Windows
& "[CLAUDE_PATH]" mcp list
```
```bash
# Mac/Linux
"[CLAUDE_PATH]" mcp list
```

`google-tag-manager` が一覧に表示されれば OK。  
表示されない場合は STEP 4 を再実行する。

---

## STEP 6: VSCode を再起動して動作確認する

ユーザーに伝える:  
「VSCode を完全に閉じて、再度開いてください。  
`Developer: Reload Window` ではなく、アプリ自体の再起動が必要です」

再起動後、チャットで以下を試してもらう:
```
GTMのアカウント一覧を表示して
```

GTM のデータが返ってきたら**セットアップ完了**。

---

## トラブルシューティング

| 症状 | 確認すること | 対処 |
|------|------------|------|
| `node` が見つからない | `node -v` を実行 | Node.js をインストール |
| `mcp-google-tag-manager.js` が見つからない | `npm list -g` で確認 | `npm install -g @weppa-cloud/mcp-google-tag-manager` を再実行 |
| `mcp list` で何も表示されない | `~/.claude.json` を確認 | `settings.json` ではなく `mcp add` で登録 |
| 認証エラー（credentials） | キーのパスを再確認 | 日本語・スペースのないパスに移動 |
| 権限エラー | GTM のユーザー管理を確認 | サービスアカウントを管理者権限・アカウントレベルで追加 |
| ツールがチャットに出てこない | Reload Window を使っていないか | VSCode を完全再起動 |
| `mcp add` でエラー | エラーメッセージを確認 | `[NODE_PATH]` と `[JS_PATH]` の値を再確認 |

---

## 参考: このアカウントの設定値

- GCPプロジェクトID: `iron-axon-162706`
- GTMアカウント: `n2p.co.jp`（ID: 3108186510）
- キーファイルのパスはスキル起動時にユーザーへ確認する（「サービスアカウントキー (.json) のフルパスを教えてください」）
