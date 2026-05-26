---
name: ga4-mcp-connect
description: GA4 MCPサーバーをClaude Code に接続するセットアップ手順。「GA4に接続したい」「analytics-mcpを設定したい」「GA4のMCPが動かない」などの場合に使用する。
---

## 進め方

**各ステップで確認コマンドを実行し、成功を確認してから次に進む。**  
環境（OS・パス・インストール済みツール）はステップ0で把握してから作業する。

---

## STEP 0: 環境を確認する

### OSとツールの確認

**Windows:**
```powershell
$PSVersionTable.OS
python --version
pipx --version
```

**Mac/Linux:**
```bash
uname -s && python3 --version && pipx --version
```

### `claude` コマンドの場所を確認

**Windows（まずPATHを確認、なければデスクトップアプリを探す）:**
```powershell
$CLAUDE = (Get-Command claude -ErrorAction SilentlyContinue).Source
if (-not $CLAUDE) {
    $CLAUDE = Get-ChildItem "$env:LOCALAPPDATA\Packages\Claude_*" -Recurse -Filter "claude.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}
Write-Host "Claude path: $CLAUDE"
```

**Mac/Linux:**
```bash
CLAUDE=$(which claude) && echo "Claude path: $CLAUDE"
```

→ 見つかったパスを控えておく。見つからない場合は Claude Code をインストールする。

### `analytics-mcp` の確認（インストール済みか）

**Windows:**
```powershell
(Get-Command analytics-mcp -ErrorAction SilentlyContinue).Source
Test-Path "$env:USERPROFILE\.local\bin\analytics-mcp.exe"
```

**Mac/Linux:**
```bash
which analytics-mcp 2>/dev/null || ls ~/.local/bin/analytics-mcp 2>/dev/null || echo "not found"
```

---

## STEP 1: サービスアカウントキーを準備する

### 2台目以降（GCPは設定済みの場合）

既存のJSONキーファイルをこの端末にコピーするだけでOK。  
パスを控えておく。パスに日本語・スペースが含まれる場合は英数字のパスに置く。

**キーファイルの内容を確認（読み取れるか検証）:**

**Windows:**
```powershell
Get-Content "C:\path\to\key.json" | ConvertFrom-Json | Select-Object project_id, client_email
```

**Mac/Linux:**
```bash
python3 -c "import json,sys; d=json.load(open('/path/to/key.json')); print(d['project_id'], d['client_email'])"
```

`project_id` と `client_email` が表示されればOK。

### 初回（GCP未設定の場合）

1. [Google Cloud Console](https://console.cloud.google.com/) で以下のAPIを有効化:
   - Google Analytics Admin API
   - Google Analytics Data API
2. IAM > サービスアカウント > 作成 → JSONキーをダウンロード
3. GA4管理画面 > プロパティのアクセス管理 → サービスアカウントのメールを「閲覧者」で追加

---

## STEP 2: analytics-mcp をインストールする

### インストール（OS共通）

```bash
pipx install analytics-mcp
pipx ensurepath
```

**インストール後、新しいターミナルを開いてから**パスを確認:

**Windows:**
```powershell
$MCP = (Get-Command analytics-mcp -ErrorAction SilentlyContinue).Source
if (-not $MCP) { $MCP = "$env:USERPROFILE\.local\bin\analytics-mcp.exe" }
Write-Host "analytics-mcp: $MCP"
Test-Path $MCP
```

**Mac/Linux:**
```bash
MCP=$(which analytics-mcp 2>/dev/null || echo "$HOME/.local/bin/analytics-mcp")
echo "analytics-mcp: $MCP" && ls "$MCP"
```

→ `True` または `ls` でファイルが表示されればOK。  
→ 見つからない場合は `pipx list` で `analytics-mcp` が含まれるか確認する。

---

## STEP 3: MCP サーバーを登録する

STEP 0 の `$CLAUDE`、STEP 2 の `$MCP`、STEP 1 のキーパスを使う。

**重要:** `~/.claude/settings.json` に直接書いても読まれない。  
必ず `mcp add` コマンドで登録する（設定は `~/.claude.json` に書き込まれる）。

**Windows:**
```powershell
& $CLAUDE mcp add --scope user google-analytics $MCP `
  -e GOOGLE_APPLICATION_CREDENTIALS="C:\path\to\key.json" `
  -e GOOGLE_PROJECT_ID="your-project-id"
```

**Mac/Linux:**
```bash
$CLAUDE mcp add --scope user google-analytics "$MCP" \
  -e GOOGLE_APPLICATION_CREDENTIALS="/path/to/key.json" \
  -e GOOGLE_PROJECT_ID="your-project-id"
```

スコープ:
- `--scope user` — 全プロジェクトで有効（推奨）
- `--scope local` — カレントディレクトリのみ

---

## STEP 4: 登録を確認する

**Windows:**
```powershell
& $CLAUDE mcp list
```

**Mac/Linux:**
```bash
$CLAUDE mcp list
```

`google-analytics` が表示されればOK。`✓ Connected` であれば接続成功。  
表示されない場合 → STEP 3 を再実行。

---

## STEP 5: VSCode を再起動して動作確認する

VSCode を**完全に閉じて**再度開く。  
（`Developer: Reload Window` では不十分な場合がある）

再起動後、チャットで確認:
```
GA4のプロパティ一覧を見せて
```

以下のツールが応答すればセットアップ完了:
- `mcp__google-analytics__get_account_summaries`
- `mcp__google-analytics__run_report`
- `mcp__google-analytics__run_realtime_report` など

---

## トラブルシューティング

| 症状 | 確認コマンド | 対処 |
|------|------------|------|
| `pipx` が見つからない | `python -m pip show pipx` | `python -m pip install --user pipx` |
| `analytics-mcp` が見つからない | `pipx list` | 新しいターミナルを開く。それでもなければ再インストール |
| `mcp list` で「No servers configured」 | — | `settings.json` ではなく `mcp add` で登録する |
| 認証エラー（credentials） | キーパスを確認 | パスに日本語・スペースがあれば英数字パスに移動 |
| ツールがチャットに出てこない | — | VSCode を完全に再起動（Reload Window は不可） |
| `claude` が見つからない | STEP 0 を再実行 | デスクトップアプリのexeをフルパスで指定 |
| `mcp add` でエラー | — | `$CLAUDE`・`$MCP` の値を `Write-Host` で確認してから再実行 |

---

## このアカウントの設定値（参考）

- GCPプロジェクトID: `iron-axon-162706`
- GA4プロパティ: `n2p.co.jp_GA4` (ID: 305939297)
- キーファイル（メイン端末）: `C:\Users\kazug\.claude\GA4MCP接続\.claude\ga4-key\iron-axon-162706-69c5fb3aef9f.json`
