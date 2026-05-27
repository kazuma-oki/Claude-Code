---
name: gsc-mcp-connect
description: Google Search Console MCPサーバーをClaude Code に接続するセットアップ手順。「GSCに接続したい」「Search ConsoleのMCPを設定したい」「gsc-mcpが動かない」などの場合に使用する。
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

4. **GSCの初期設定が済んでいるかを確認する**  
   「GSCのAPIアクセス権（サービスアカウントをGSCプロパティに追加）はすでに設定済みですか？」  
   - 済 → STEP 1（キー準備）から開始  
   - 未 → STEP 0（GCP/GSC設定）から開始

確認が取れたら「では順番に進めましょう」と伝え、STEP を開始する。

---

## STEP 0: GCP / GSC の初期設定（初回のみ・既に設定済みならスキップ）

以下をユーザーに案内する（ブラウザ操作のため、完了したら次へ進むよう伝える）:

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセスし、対象プロジェクトを開く

2. **Google Search Console API** を有効化する:  
   「APIとサービス」→「ライブラリ」→ `Google Search Console API` を検索して「有効にする」

3. サービスアカウントを作成しキーをダウンロードする:
   - 「IAMと管理」→「サービスアカウント」→「作成」
   - 役割は「閲覧者」でOK
   - 作成後、「キー」タブ → 「鍵を追加」→「新しい鍵を作成」→ JSON 形式でダウンロード

4. GSC でアクセスを許可する:
   - [Google Search Console](https://search.google.com/search-console) を開く
   - データにアクセスしたいプロパティを選択
   - 左メニュー下部「設定」→「ユーザーと権限」→「ユーザーを追加」
   - サービスアカウントのメールアドレスを入力し、権限「フル」を選択して追加

完了したら確認: 「JSONキーのダウンロードと GSC のアクセス設定は完了しましたか？」

---

## STEP 1: サービスアカウントキーをこの端末に配置する

### キーが別端末にある場合

ユーザーに伝える:  
「キーファイル（.json）を別端末からこの端末にコピーしてください。  
パスに**日本語・スペースが含まれない場所**に置いてください（例: `C:\Users\<名前>\.claude\gsc-key\` など）」

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

## STEP 2: Python と uv / uvx の確認・インストール

以下を実行してもらい、結果を確認する:

**Windows:**
```powershell
python --version
uvx --version
```

**Mac/Linux:**
```bash
python3 --version && uvx --version
```

- `python` が見つからない → Python をインストールしてもらう（python.org）
- `uvx` が見つからない → 以下を実行してもらう:

```bash
pip install uv
```

インストール後、uvx のパスを確認する:

**Windows:**
```powershell
python -c "import sysconfig; print(sysconfig.get_path('scripts'))"
```

表示されたパス（例: `C:\Users\<名前>\AppData\Local\Programs\Python\Python312\Scripts`）に `uvx.exe` があるか確認:

```powershell
Get-ChildItem "<上で表示されたパス>" | Where-Object { $_.Name -match "^uvx" }
```

**Mac/Linux:**
```bash
which uvx
```

パスが確認できたら OK。→ このパスを以降 `[UVX_PATH]` として控えてもらう。

---

## STEP 3: claude コマンドのパスを確認する

> **⚠️ 重要:** VSCode 拡張機能として Claude Code を使っている場合、`claude` コマンドが CLI として使えないことがある。その場合は STEP 4-A をスキップして **STEP 4-B（直接編集）** に進む。

**Windows:**
```powershell
(Get-Command claude -ErrorAction SilentlyContinue).Source
```
表示されない場合:
```powershell
Get-ChildItem "$env:LOCALAPPDATA\AnthropicClaude\" -Filter "claude.exe" -Recurse -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName
```

**Mac/Linux:**
```bash
which claude
```

- パスが表示された → `[CLAUDE_PATH]` として控え、**STEP 4-A** へ  
- 表示されない → **STEP 4-B** へ（`~/.claude.json` を直接編集する方法）

---

## STEP 4-A: `claude mcp add` コマンドで登録する（CLI が使える場合）

STEP 1 のキーパス、STEP 2 の `[UVX_PATH]`、STEP 3 の `[CLAUDE_PATH]` を使ってコマンドを組み立てて提示する。

**Windows:**
```powershell
& "[CLAUDE_PATH]" mcp add --scope user google-search-console "[UVX_PATH]" `
  --args "mcp-search-console" `
  -e GSC_CREDENTIALS_PATH="[キーのパス]" `
  -e GSC_SKIP_OAUTH="true"
```

**Mac/Linux:**
```bash
"[CLAUDE_PATH]" mcp add --scope user google-search-console "[UVX_PATH]" \
  --args "mcp-search-console" \
  -e GSC_CREDENTIALS_PATH="[キーのパス]" \
  -e GSC_SKIP_OAUTH="true"
```

このコマンドは `~/.claude.json` の `mcpServers` セクションに自動で書き込む。  
実行後、STEP 5 の確認に進む。

---

## STEP 4-B: `~/.claude.json` を直接編集する（CLI が使えない場合）

> **⚠️ 絶対に間違えてはいけないこと:**  
> MCP サーバーの設定は `~/.claude/settings.json` **ではなく** `~/.claude.json`（ホームディレクトリ直下）に書く。  
> `settings.json` に書いてもツールとして認識されない。（実証済みの落とし穴）

`~/.claude.json`（Windows: `C:\Users\<名前>\.claude.json`）をテキストエディタで開き、  
ファイル内の `"mcpServers"` セクションを探して以下を追記する:

```json
"google-search-console": {
  "type": "stdio",
  "command": "[UVX_PATH]",
  "args": ["mcp-search-console"],
  "env": {
    "GSC_CREDENTIALS_PATH": "[キーのパス]",
    "GSC_SKIP_OAUTH": "true"
  }
}
```

**追記後のイメージ（既存の google-analytics の直後に追加）:**
```json
"mcpServers": {
  "google-analytics": {
    ...既存の設定...
  },
  "google-search-console": {
    "type": "stdio",
    "command": "C:\\Users\\<名前>\\AppData\\Local\\Programs\\Python\\Python312\\Scripts\\uvx.exe",
    "args": ["mcp-search-console"],
    "env": {
      "GSC_CREDENTIALS_PATH": "C:\\Users\\<名前>\\path\\to\\key.json",
      "GSC_SKIP_OAUTH": "true"
    }
  }
}
```

**注意点:**
- `"type": "stdio"` を必ず含める（`mcp add` コマンドが自動で付与するフィールド）
- Windows のパスはバックスラッシュを `\\` にエスケープする
- JSON の末尾カンマに注意（最後の要素にカンマを付けない）
- `~/.claude.json` は Claude Code の内部設定も含む重要ファイルのため、編集前にバックアップを推奨

保存後、STEP 5 の確認に進む。

---

## STEP 5: 登録を確認する

**CLI が使える場合:**
```powershell
# Windows
& "[CLAUDE_PATH]" mcp list
```
```bash
# Mac/Linux
"[CLAUDE_PATH]" mcp list
```
`google-search-console` が一覧に表示されれば OK。

**CLI が使えない場合:**  
`~/.claude.json` を開き、`"mcpServers"` に `"google-search-console"` が追記されているか目視確認する。

---

## STEP 6: VSCode を再起動して動作確認する

ユーザーに伝える:  
「VSCode を完全に閉じて、再度開いてください。  
`Developer: Reload Window` ではなく、アプリ自体の再起動が必要です」

再起動後、チャットで以下を試してもらう:
```
Google Search Consoleに接続して、登録されているプロパティの一覧を教えて
```

GSC のデータが返ってきたら**セットアップ完了**。  
ツールが出てこない場合は VSCode の完全再起動を再度試す。

---

## トラブルシューティング

問題が起きたらユーザーに症状を確認し、以下を参考に対処する:

| 症状 | 確認すること | 対処 |
|------|------------|------|
| ツールがチャットに出てこない（最重要） | 設定を書いたファイルが `~/.claude.json` か確認 | `settings.json` ではなく `~/.claude.json` のルート `mcpServers` に書く（STEP 4-B） |
| `uvx` が見つからない | `pip install uv` を実行したか | 新しいターミナルを開いて再確認 |
| `uvx.exe` のパスが分からない | `python -c "import sysconfig; print(sysconfig.get_path('scripts'))"` | 表示されたパスを `[UVX_PATH]` に使う |
| `claude` コマンドが見つからない | VSCode 拡張として使っているか | STEP 4-B（`~/.claude.json` 直接編集）に切り替える |
| `mcp list` で何も表示されない | `~/.claude.json` の `mcpServers` を確認 | `"type": "stdio"` が含まれているか確認、なければ追記 |
| 認証エラー（credentials） | キーのパスを再確認 | 日本語・スペースのないパスに移動 |
| OAuth の認証画面が開く | `GSC_SKIP_OAUTH` が設定されているか | `"true"` (文字列) になっているか確認 |
| プロパティが見つからない | STEP 0 の GSC ユーザー追加が完了しているか | サービスアカウントメールを「フル」権限で追加 |
| 再起動してもツールが出ない | `~/.claude.json` の JSON 構文エラー | エディタで開き、カンマ漏れ・括弧ミスを確認 |
| VSCode Reload Window 後もダメ | Reload Window は不十分 | VSCode アプリ自体を完全終了して再起動 |

### 特に注意: settings.json と .claude.json の違い

```
~/.claude/settings.json   ← ❌ MCP登録には使えない（権限設定などに使う）
~/.claude.json            ← ✅ MCP登録はここに書く（ホームディレクトリ直下）
```

この違いを間違えると、設定しても一切ツールが認識されない。（実際に発生した問題）

---

## 参考: このアカウントの設定値

- GCPプロジェクトID: `iron-axon-162706`
- サービスアカウント: `ga4-mcp-server@iron-axon-162706.iam.gserviceaccount.com`
- キーファイルのパスはスキル起動時にユーザーへ確認する（「サービスアカウントキー (.json) のフルパスを教えてください」）
