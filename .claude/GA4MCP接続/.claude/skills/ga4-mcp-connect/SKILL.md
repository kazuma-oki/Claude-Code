---
name: ga4-mcp-connect
description: GA4 MCPサーバーをClaude Code に接続するセットアップ手順。「GA4に接続したい」「analytics-mcpを設定したい」「GA4のMCPが動かない」などの場合に使用する。
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

2. **サービスアカウントキーの有無を確認する**  
   「GCPのサービスアカウントキー（.json ファイル）はこの端末にありますか？」  
   - ある → パスを教えてもらう  
   - ない → 「別の端末からコピーできますか？それとも新規作成が必要ですか？」

3. **GCPプロジェクトIDを確認する**  
   「GCPのプロジェクトIDを教えてください」  
   （参考: このアカウントは `iron-axon-162706`）

4. **GCPが未設定かどうかを確認する**  
   「GA4のプロパティへのアクセス権はすでに設定済みですか？」  
   - 済 → STEP 1（キー準備）から開始  
   - 未 → STEP 0（GCP設定）から開始

確認が取れたら「では順番に進めましょう」と伝え、STEP を開始する。

---

## STEP 0: GCP の初期設定（初回のみ・既に設定済みならスキップ）

以下をユーザーに案内する（ブラウザ操作のため、完了したら次へ進むよう伝える）:

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセスし、対象プロジェクトを開く

2. 以下の2つのAPIを有効化する:
   - **Google Analytics Admin API**
   - **Google Analytics Data API**  
   （「APIとサービス」→「ライブラリ」から検索して有効化）

3. サービスアカウントを作成しキーをダウンロードする:
   - 「IAMと管理」→「サービスアカウント」→「作成」
   - 役割は「閲覧者」でOK
   - 作成後、「キー」タブ → 「鍵を追加」→「新しい鍵を作成」→ JSON 形式でダウンロード

4. GA4でアクセスを許可する:
   - GA4管理画面 →「プロパティのアクセス管理」
   - サービスアカウントのメールアドレスを「閲覧者」として追加

完了したら確認: 「JSONキーのダウンロードと GA4 のアクセス設定は完了しましたか？」

---

## STEP 1: サービスアカウントキーをこの端末に配置する

### キーが別端末にある場合

ユーザーに伝える:  
「キーファイル（.json）を別端末からこの端末にコピーしてください。  
パスに**日本語・スペースが含まれない場所**に置いてください（例: `C:\Users\<名前>\.claude\ga4-key\` など）」

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
エラーが出た場合はパスを確認する（バックスラッシュの数、スペースの有無など）。

---

## STEP 2: Python / pipx の確認

以下を実行してもらい、結果を確認する:

**Windows:**
```powershell
python --version
pipx --version
```

**Mac/Linux:**
```bash
python3 --version && pipx --version
```

- `python` が見つからない → Python をインストールしてもらう（python.org）
- `pipx` が見つからない → 以下を実行してもらう:

```bash
python -m pip install --user pipx
python -m pipx ensurepath
```

その後**新しいターミナルを開いて** `pipx --version` を再確認する。

---

## STEP 3: analytics-mcp をインストールする

```bash
pipx install analytics-mcp
pipx ensurepath
```

**新しいターミナルを開いてから**インストールパスを確認してもらう:

**Windows:**
```powershell
(Get-Command analytics-mcp -ErrorAction SilentlyContinue).Source
```
表示されない場合:
```powershell
Test-Path "$env:USERPROFILE\.local\bin\analytics-mcp.exe"
```

**Mac/Linux:**
```bash
which analytics-mcp
```

パスが表示されれば OK。表示されない場合は `pipx list` を実行してもらい `analytics-mcp` が含まれるか確認する。  
→ このパスを以降 `[MCP_PATH]` として使う。控えてもらう。

---

## STEP 4: claude コマンドのパスを確認する

**Windows（PATHにあるか確認→なければデスクトップアプリから探す）:**
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

パスが表示されれば OK。  
→ このパスを以降 `[CLAUDE_PATH]` として使う。控えてもらう。  
表示されない場合は Claude Code がインストールされていないので、インストールを案内する。

---

## STEP 5: MCP サーバーを登録する

STEP 1 のキーパス、STEP 3 の `[MCP_PATH]`、STEP 4 の `[CLAUDE_PATH]`、確認済みのプロジェクトIDを使ってコマンドを組み立てて提示する。

**Windows:**
```powershell
& "[CLAUDE_PATH]" mcp add --scope user google-analytics "[MCP_PATH]" `
  -e GOOGLE_APPLICATION_CREDENTIALS="[キーのパス]" `
  -e GOOGLE_PROJECT_ID="[プロジェクトID]"
```

**Mac/Linux:**
```bash
"[CLAUDE_PATH]" mcp add --scope user google-analytics "[MCP_PATH]" \
  -e GOOGLE_APPLICATION_CREDENTIALS="[キーのパス]" \
  -e GOOGLE_PROJECT_ID="[プロジェクトID]"
```

**重要な注意点をユーザーに伝える:**
- `~/.claude/settings.json` に直接書いても読まれない
- 必ず `mcp add` コマンドで登録する（`~/.claude.json` に書き込まれる）
- `--scope user` を使うと全プロジェクトで有効になる（推奨）

実行後、エラーメッセージが出た場合は内容を教えてもらい対処する。

---

## STEP 6: 登録を確認する

```powershell
# Windows
& "[CLAUDE_PATH]" mcp list
```
```bash
# Mac/Linux
"[CLAUDE_PATH]" mcp list
```

`google-analytics` が一覧に表示されれば OK。  
`✓ Connected` と表示されれば接続まで完了。  
表示されない場合は STEP 5 を再実行する。

---

## STEP 7: VSCode を再起動して動作確認する

ユーザーに伝える:  
「VSCode を完全に閉じて、再度開いてください。  
`Developer: Reload Window` ではなく、アプリ自体の再起動が必要です」

再起動後、チャットで以下を試してもらう:
```
GA4のプロパティ一覧を見せて
```

GA4のデータが返ってきたら**セットアップ完了**。  
ツールが出てこない場合は VSCode の完全再起動を再度試す。

---

## トラブルシューティング

問題が起きたらユーザーに症状を確認し、以下を参考に対処する:

| 症状 | 確認すること | 対処 |
|------|------------|------|
| `pipx` が見つからない | `python -m pip show pipx` | `python -m pip install --user pipx` |
| `analytics-mcp` が見つからない | 新しいターミナルを開いたか | `pipx list` で確認、再インストール |
| `mcp list` で何も表示されない | `~/.claude.json` を確認 | `settings.json` ではなく `mcp add` で登録 |
| 認証エラー（credentials） | キーのパスを再確認 | 日本語・スペースのないパスに移動 |
| ツールがチャットに出てこない | Reload Window を使っていないか | VSCode を完全再起動 |
| `claude` コマンドが見つからない | STEP 4 を再実行 | Claude Code がインストールされているか確認 |
| `mcp add` でエラー | エラーメッセージを確認 | `[CLAUDE_PATH]` と `[MCP_PATH]` の値を再確認 |

---

## 参考: このアカウントの設定値

- GCPプロジェクトID: `iron-axon-162706`
- GA4プロパティ: `n2p.co.jp_GA4` (ID: 305939297)
- キーファイルのパスはスキル起動時にユーザーへ確認する（「サービスアカウントキー (.json) のフルパスを教えてください」）
